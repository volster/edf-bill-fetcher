"""One durable job per child process, using the existing application pipeline."""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
from collections.abc import Callable
from dataclasses import replace
from datetime import date, datetime, timezone
from functools import partial
from pathlib import Path
from threading import Event
from typing import cast
from zipfile import ZipFile, is_zipfile

from edf_bill_fetcher.application.run_service import (
    RunCancelled,
    bundle_engine_sources,
    collect_engine,
    run,
)
from edf_bill_fetcher.application.workbook_review import bundle_review_sources
from edf_bill_fetcher.domain.analysis import AnalysisSnapshot
from edf_bill_fetcher.domain.json_values import freeze_json, object_value, rows_value
from edf_bill_fetcher.domain.requests import RunRequest, SourceSelection
from edf_bill_fetcher.io.persistence.result_codec import migrate_legacy
from edf_bill_fetcher.io.persistence.result_store import load_result, save_result
from edf_bill_fetcher.io.persistence.run_manifest import artifact_record
from edf_bill_fetcher.io.review_workbook import read_workbook
from edf_bill_fetcher.models.config import ConfigDict
from edf_bill_fetcher.models.defaults import EXTRACTION_DEFAULTS
from edf_bill_fetcher.models.report_view import RenderSelection
from edf_bill_fetcher.web.contracts import Job, Source
from edf_bill_fetcher.web.storage import Store, contained


def copy_accepted_sources(
    snapshot: AnalysisSnapshot, source_root: Path, directory: Path, allowed: Callable[[str], bool]
) -> tuple[dict, tuple[str, ...]]:
    """Keep promised relative paths while copying verified sources for an unchanged result."""
    import pandas as pd

    from edf_bill_fetcher.io.writers.evidence_bundle import (
        build_bundle_index,
        build_bundle_manifest,
        resolve_bundle_paths,
    )

    state = object_value(snapshot.prepared.extraction.legacy_state)
    paths, warnings = resolve_bundle_paths(state, str(source_root))
    notes = list(warnings)
    saved = {}
    manifest = state.get("source_manifest", {})
    for identity, metadata in manifest.items():
        original = paths.get(identity)
        if not original and not metadata.get("Saved Path"):
            candidate = metadata.get("Original Path")
            if (
                candidate
                and allowed(candidate)
                and Path(candidate).is_file()
                and metadata.get("Content Hash") == artifact_record(candidate)["sha256"]
            ):
                original = candidate
        if not original:
            notes.append("Source unavailable: " + str(metadata.get("Original Filename", identity)))
            continue
        relative = (
            metadata.get("Saved Path")
            or "evidence_files/"
            + hashlib.sha256(identity.encode()).hexdigest()
            + Path(original).suffix
        )
        target = contained(directory, relative)
        if target.parent == directory and target.name in {
            "result.json",
            "run.json",
            "worker.log",
            "Evidence.xlsx",
            "Evidence_Report.html",
            "Evidence_Report.pdf",
            "Evidence_Report.docx",
        }:
            raise ValueError("A preserved source path collides with a report destination")
        target.parent.mkdir(parents=True, exist_ok=True)
        if (
            target.exists()
            and artifact_record(str(target))["sha256"] != artifact_record(original)["sha256"]
        ):
            raise ValueError("Conflicting preserved source paths")
        shutil.copyfile(original, target)
        saved[metadata["Attachment Name"]] = str(target)
        saved[identity] = str(target)
    records = pd.DataFrame(rows_value(snapshot.prepared.records) + list(manifest.values()))
    build_bundle_manifest(records, saved, str(directory))
    build_bundle_index(records, saved, str(directory / "evidence_index.docx"))
    return saved, tuple(dict.fromkeys(notes))


class Cancellation(Event):
    """Adapt persisted cancellation to the collector's Event interface."""

    def __init__(self, store: Store, identity: str):
        """Capture the job identity, not a live browser request."""
        super().__init__()
        self.store, self.identity = store, identity

    def is_set(self) -> bool:
        """Observe cancellation across processes and OCR polling loops."""
        return super().is_set() or bool(self.store.get(self.identity)["data"].get("cancel"))


def source_path(store: Store, source: Source, roots: dict[str, str]) -> Path:
    """Resolve uploaded inputs or explicitly allowed server roots."""
    if source.intake_id:
        entry = store.get(source.intake_id, "job")
        if entry["data"]["status"] != "complete" or entry["data"]["request"]["mode"] != "intake":
            raise ValueError("Wait for a completed intake before selecting its files")
        directory = store.directory(entry)
        if source.relative == "__all__":
            return directory / "intake.json"
        rows = json.loads((directory / "intake.json").read_text(encoding="utf-8"))
        row = next((row for row in rows if row["id"] == source.relative), None)
        if row is None or row["status"] != "needs_choice":
            raise ValueError("Choose a saved analysis or workbook from the completed intake")
        target = Path(row["path"]).resolve()
        if not store.path_allowed(entry.get("workspace_id"), str(target), roots):
            raise ValueError("Intake source is outside the configured folders")
        return target
    if source.asset_id:
        asset = store.get(source.asset_id, "asset")
        if not asset["data"].get("sealed") or asset["data"]["kind"] != source.kind:
            raise ValueError("Upload is incomplete or has the wrong source type")
        directory = store.directory(asset)
        if source.kind == "auto" and source.relative:
            return contained(directory, source.relative)
        return (
            directory
            if source.kind == "auto" or source.kind.endswith("_dir")
            else contained(directory, asset["data"]["files"][0])
        )
    if source.root_id not in roots:
        raise ValueError("Unknown source folder")
    if source.relative in {".", "./"}:
        return Path(roots[cast(str, source.root_id)]).resolve()
    return contained(Path(roots[cast(str, source.root_id)]), cast(str, source.relative))


def unpack_bundle(
    path: Path,
    target: Path,
    max_bytes: int = 64 * 1024**3,
    *,
    cancelled: Callable[[], bool] = lambda: False,
) -> Path:
    """Expand a portable pack with traversal, symlink and expansion checks."""
    from edf_bill_fetcher.io.adapters.intake import validate_zip_members

    with ZipFile(path) as archive:
        infos = archive.infolist()
        total = sum(info.file_size for info in infos)
        if len(infos) > 100_000 or total > max_bytes:
            raise ValueError("Evidence pack exceeds the import limit")
        validate_zip_members(infos)
        target.mkdir(parents=True, exist_ok=True)
        if total > shutil.disk_usage(target).free:
            raise ValueError("Not enough free space to import this evidence pack")
        expanded = 0
        for info in infos:
            if cancelled():
                raise RunCancelled("Evidence pack import cancelled")
            destination = contained(target, info.filename.rstrip("/"))
            if info.is_dir():
                destination.mkdir(parents=True, exist_ok=True)
            else:
                destination.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(info) as source, destination.open("xb") as output:
                    for block in iter(lambda: source.read(1024 * 1024), b""):
                        if cancelled():
                            raise RunCancelled("Evidence pack import cancelled")
                        expanded += len(block)
                        if expanded > max_bytes:
                            raise ValueError("Evidence pack exceeds the import limit")
                        output.write(block)
    results = list(target.rglob("result.json"))
    if len(results) != 1:
        raise ValueError("An imported pack must contain exactly one result.json")
    return results[0]


def input_path(
    store: Store,
    job: Job,
    directory: Path,
    roots: dict[str, str],
    *,
    cancelled: Callable[[], bool] = lambda: False,
) -> Path:
    """Locate a parent result or materialize a portable imported pack."""
    if job.parent_id:
        parent = store.get(job.parent_id, "job")
        return store.directory(parent) / "result.json"
    if job.input is None:
        raise ValueError("No input selected")
    path = source_path(store, job.input, roots)
    if (
        job.input.kind == "bundle"
        or job.input.kind == "auto"
        and job.mode != "workbook"
        and is_zipfile(path)
    ):
        return unpack_bundle(path, directory / "imported", cancelled=cancelled)
    return path


def execute(store: Store, identity: str, roots: dict[str, str]) -> None:
    """Publish truthful status and preserve valid artifacts on every failure."""
    entry = store.get(identity, "job")
    workspace_id = entry.get("workspace_id")
    if workspace_id is not None:
        roots = store.accounts.roots(workspace_id, roots)
    directory = store.directory(entry)
    token = Cancellation(store, identity)
    engine = None
    warnings: tuple[str, ...] = ()
    renderers = None

    def allowed(path):
        return store.path_allowed(workspace_id, path, roots)

    def resource_config(config: dict, *, resolve: bool) -> dict:
        value = config.get("ofgem_caps_path")
        if not value:
            return config
        path = str(value)
        prefix, separator, relative = path.partition(":")
        if resolve and separator and prefix in roots:
            path = str(contained(Path(roots[prefix]), relative))
        if not allowed(path):
            raise ValueError(
                "Custom cap data must be inside a configured source folder; select NAME:relative/path.csv or explicitly reanalyse an imported result with a permitted location"
            )
        return {**config, "ofgem_caps_path": path}

    try:
        job = Job.model_validate(entry["data"]["request"])
        if token.is_set():
            raise RunCancelled("Run cancelled before starting")
        for source in [*job.sources, *([job.input] if job.input else [])]:
            source_identity = source.asset_id or source.intake_id
            if source_identity and store.get(source_identity)["workspace_id"] != workspace_id:
                raise ValueError("Input belongs to another workspace")
        for related_identity in (job.parent_id, job.other_id):
            if (
                related_identity
                and store.get(related_identity, "job")["workspace_id"] != workspace_id
            ):
                raise ValueError("Run belongs to another workspace")
        if actor_id := entry["data"].get("actor_id"):
            if workspace_id is None:
                raise ValueError("Queued run has no workspace")
            if not store.accounts.user(actor_id)["active"]:
                raise ValueError("The account that queued this run was disabled")
            store.accounts.require(actor_id, workspace_id, "editor")
        config = (
            {**EXTRACTION_DEFAULTS, **job.config}
            if job.mode in {"extract", "intake"}
            else dict(job.config)
        )
        config.setdefault("as_of", date.today().isoformat())
        request = RunRequest(
            identity,
            datetime.now(timezone.utc),
            date.fromisoformat(config["as_of"]),
            (),
            freeze_json(config),
        )
        accepted = None
        collect = None
        copy_sources = None
        if job.mode == "intake":
            from edf_bill_fetcher.application.intake_execution import inspect_sources

            inspect_sources(
                [str(source_path(store, source, roots)) for source in job.sources],
                str(directory),
                cancelled=token.is_set,
                progress=lambda message: store.log(identity, message),
            )
            if token.is_set():
                raise RunCancelled("Intake cancelled")
            store.patch(identity, status="complete")
            return
        if job.mode == "extract":
            from edf_bill_fetcher.collectors.engine import EvidenceEngine

            config = resource_config(config, resolve=True)
            request = replace(request, config=freeze_json(config))
            sources = tuple(
                SourceSelection(
                    "intake" if s.intake_id and s.relative == "__all__" else s.kind,
                    str(source_path(store, s, roots)),
                )
                for s in job.sources
            )
            request = RunRequest(
                identity, request.created_at, request.as_of, sources, request.config
            )
            engine = EvidenceEngine(
                cast(ConfigDict, config),
                lambda message: store.log(identity, message),
                cancel_event=token,
                allowed_source=allowed,
            )

            def collect():
                return collect_engine(engine, request)

            def copy_sources(prepared):
                return bundle_engine_sources(
                    engine,
                    prepared,
                    str(directory),
                    legacy_paths=(str(directory / "records.json"),),
                )
        else:
            original = input_path(store, job, directory, roots, cancelled=token.is_set)
            if job.mode == "workbook":
                records, state, as_of = read_workbook(str(original))
                config = {**state.get("review_config", config), **job.config}
                config["as_of"] = job.config.get("as_of") or as_of or config["as_of"]
            else:
                accepted = load_result(original, legacy_request=request)
                if job.mode == "diff":
                    from edf_bill_fetcher.helpers.artifacts import write_json
                    from edf_bill_fetcher.io.writers.diff import write_diff_workbook
                    from edf_bill_fetcher.processors.run_diff import diff_records

                    other = load_result(
                        store.directory(store.get(cast(str, job.other_id), "job")) / "result.json"
                    )
                    diff = diff_records(
                        rows_value(accepted.prepared.records), rows_value(other.prepared.records)
                    )
                    write_json(diff, str(directory / "comparison.json"))
                    write_diff_workbook(diff, str(directory / "comparison.xlsx"))
                    if token.is_set():
                        raise RunCancelled("Comparison cancelled")
                    store.patch(identity, status="complete")
                    return
                if job.mode in {"review", "preview"}:
                    from edf_bill_fetcher.application.review_execution import (
                        preview_review,
                        publish_review,
                    )

                    resource_config(object_value(accepted.prepared.effective_config), resolve=False)
                    if job.mode == "preview":
                        snapshot = preview_review(
                            str(original), cast(dict, job.draft), run_id=identity
                        )
                        if token.is_set():
                            raise RunCancelled("Preview cancelled")
                        save_result(snapshot, directory / "result.json", result_id=identity)
                        store.patch(identity, status="complete")
                    else:
                        outcome = publish_review(
                            str(original),
                            cast(dict, job.draft),
                            str(directory),
                            formats=tuple(job.formats),
                            sections=tuple(job.sections) if job.sections is not None else None,
                            run_id=identity,
                            cancelled=token.is_set,
                            progress=lambda event: store.log(identity, event.message),
                            allowed_source=allowed,
                        )
                        store.patch(identity, status=outcome.status, error=outcome.error)
                    return
                else:
                    # Originals accompany rerendered reports, while the accepted
                    # result and its calculations remain byte-for-byte semantic peers.
                    from edf_bill_fetcher.application.run_service import default_renderers
                    from edf_bill_fetcher.io.writers.export import render_xlsx

                    previous = object_value(accepted.prepared.effective_config)
                    changed = {
                        key for key, value in job.config.items() if previous.get(key) != value
                    }
                    config = {**previous, **job.config}
                    if changed - {"report_sections", "report_account_ref"}:
                        config = resource_config(config, resolve=True)
                        state = object_value(accepted.prepared.extraction.legacy_state)
                        request = RunRequest(
                            identity,
                            request.created_at,
                            date.fromisoformat(config["as_of"]),
                            (SourceSelection("records_json", str(original)),),
                            freeze_json(config),
                        )
                        accepted = migrate_legacy(
                            {
                                **state,
                                "version": 2,
                                "records": rows_value(accepted.prepared.records),
                                "original_records": [
                                    object_value(row.raw)
                                    for row in accepted.prepared.extraction.observations
                                ],
                            },
                            request,
                        )
                    elif changed:
                        accepted = replace(
                            accepted,
                            prepared=replace(
                                accepted.prepared, effective_config=freeze_json(config)
                            ),
                        )

                    saved, warnings = copy_accepted_sources(
                        accepted, original.parent, directory, allowed
                    )
                    for warning in warnings:
                        store.log(identity, warning)
                    renderers = default_renderers()
                    renderers["xlsx"] = partial(render_xlsx, evidence_paths=saved)
            if accepted is None:
                config = resource_config(config, resolve=True)
                request = RunRequest(
                    identity,
                    request.created_at,
                    date.fromisoformat(config["as_of"]),
                    (SourceSelection("records_json", str(original)),),
                    freeze_json(config),
                )

                def collect():
                    return (records, state)

                def copy_sources(prepared):
                    return bundle_review_sources(
                        prepared, str(directory), str(original), allowed_source=allowed
                    )

        formats = tuple(dict.fromkeys(["xlsx", *job.formats]))
        outcome = run(
            request,
            {
                fmt: str(
                    directory / ("Evidence" + ("_Report" if fmt != "xlsx" else "") + "." + fmt)
                )
                for fmt in formats
            },
            result_path=str(directory / "result.json"),
            manifest_path=str(directory / "run.json"),
            collect=collect,
            copy_sources=copy_sources,
            cleanup=engine.cleanup if engine else lambda: None,
            cancelled=token.is_set,
            progress=lambda value: store.log(identity, value.message or value.stage),
            accepted_result=accepted,
            warnings=warnings,
            renderers=renderers,
            selection=RenderSelection(
                formats, tuple(job.sections) if job.sections is not None else None
            ),
            allowed_source=allowed,
        )
        store.patch(identity, status=outcome.status, error=outcome.error)
    except Exception as error:
        store.log(identity, str(error))
        store.patch(
            identity,
            status="cancelled" if isinstance(error, RunCancelled) else "failed",
            error=f"{type(error).__name__}: {error}",
        )
    finally:
        if engine is not None:
            engine.cleanup()


def main() -> None:
    """Run a single server-owned job; never start an HTTP server in the child."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data", required=True)
    parser.add_argument("--job", required=True)
    options = parser.parse_args()
    store = Store(options.data)
    roots_file = store.root / "roots.json"
    execute(store, options.job, json.loads(roots_file.read_text()) if roots_file.exists() else {})


if __name__ == "__main__":
    main()
