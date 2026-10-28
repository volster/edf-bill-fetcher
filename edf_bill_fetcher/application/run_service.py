"""Shared run orchestration without UI state or implicit output destinations."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any

from edf_bill_fetcher.application.analysis_service import analyse_prepared
from edf_bill_fetcher.application.legacy_adapter import capture_legacy
from edf_bill_fetcher.application.preparation_service import prepare_snapshot
from edf_bill_fetcher.domain.analysis import AnalysisSnapshot
from edf_bill_fetcher.domain.evidence import PreparedSnapshot
from edf_bill_fetcher.domain.json_values import FrozenJSON, freeze_json, object_value
from edf_bill_fetcher.domain.requests import RunRequest
from edf_bill_fetcher.io.persistence.result_store import save_result
from edf_bill_fetcher.io.persistence.run_manifest import artifact_record, save_manifest
from edf_bill_fetcher.models.report_view import RenderSelection, build_report_view
from edf_bill_fetcher.models.snapshot_adapter import to_snapshot_json
from edf_bill_fetcher.processors.claim_reconciliation import has_active_claims


class RunCancelled(Exception):
    """Cancellation observed at a stage boundary."""


@dataclass(frozen=True)
class ProgressEvent:
    """A stage update safe for GUI queue delivery."""

    stage: str
    message: str


@dataclass(frozen=True)
class RunOutcome:
    """Truthful terminal result, retaining valid analysis after renderer failure."""

    status: str
    snapshot: AnalysisSnapshot | None
    manifest: FrozenJSON
    error: str | None = None


def check_cancelled(cancelled: Callable[[], bool]) -> None:
    """Stop between stages without discarding already published artifacts."""
    if cancelled():
        raise RunCancelled("Run cancelled")


def check_source_destinations(
    request: RunRequest, destinations: list[Path], state: dict | None = None
) -> None:
    """Protect selected inputs and original/copy paths captured by source collectors."""
    protected = {Path(source.path).resolve() for source in request.sources}
    source_extensions = {"pdf_dir": ".pdf", "eml_dir": ".eml", "msg_dir": ".msg"}
    for source in request.sources:
        extension = source_extensions.get(source.kind)
        if extension:
            root = Path(source.path).resolve()
            if any(
                path.is_file() and path.is_relative_to(root) and path.suffix.lower() == extension
                for path in destinations
            ):
                raise ValueError(
                    "Artifact destination would overwrite a source in a selected folder"
                )
    state = state or {}
    for key in ("source_paths", "saved_evidence_paths"):
        paths = state.get(key, {})
        if isinstance(paths, dict):
            protected.update(
                Path(path).resolve() for path in paths.values() if isinstance(path, str) and path
            )
    manifest = state.get("source_manifest", {})
    bundle_root = state.get("source_bundle_path")
    for metadata in manifest.values() if isinstance(manifest, dict) else ():
        if not isinstance(metadata, dict):
            continue
        original = metadata.get("Original Path")
        if isinstance(original, str) and original:
            protected.add(Path(original).resolve())
        saved = metadata.get("Saved Path")
        if isinstance(saved, str) and saved and isinstance(bundle_root, str) and bundle_root:
            protected.add((Path(bundle_root) / saved).resolve())
    if protected.intersection(destinations):
        raise ValueError("Artifact destination would overwrite a selected or preserved source")


def default_renderers() -> dict[str, Callable]:
    """Resolve the four snapshot renderers at the I/O boundary."""
    from edf_bill_fetcher.io.reporters.docx_report import render_docx
    from edf_bill_fetcher.io.reporters.html_report import render_html
    from edf_bill_fetcher.io.reporters.pdf_report import render_pdf
    from edf_bill_fetcher.io.writers.export import render_xlsx

    return {"xlsx": render_xlsx, "pdf": render_pdf, "docx": render_docx, "html": render_html}


def run(
    request: RunRequest,
    outputs: dict[str, str],
    *,
    result_path: str,
    manifest_path: str,
    collect: Callable[[], tuple[list[dict], dict]] | None = None,
    copy_sources: Callable[[PreparedSnapshot], dict] | None = None,
    cleanup: Callable[[], None] = lambda: None,
    cancelled: Callable[[], bool] = lambda: False,
    progress: Callable[[ProgressEvent], None] = lambda event: None,
    renderers: dict[str, Callable] | None = None,
    store: Callable = save_result,
    accepted_result: AnalysisSnapshot | None = None,
    selection: RenderSelection | None = None,
    warnings: tuple[str, ...] = (),
) -> RunOutcome:
    """Collect once, preserve one analysis and render requested artifacts independently."""
    snapshot = accepted_result
    manifest: dict = {
        "manifest_version": 1,
        "run_id": request.run_id,
        "status": "running",
        "stage": "captured",
        "requested": dict(outputs),
        "artifacts": {},
        "result": None,
        "error": None,
        "warnings": list(warnings),
    }
    error = None
    destinations_valid = False

    def stage(name: str) -> None:
        check_cancelled(cancelled)
        manifest["stage"] = name
        if destinations_valid:
            save_manifest(manifest, manifest_path)
        progress(ProgressEvent(name, name.replace("_", " ").title()))

    try:
        destinations = [
            Path(path).resolve() for path in [result_path, manifest_path, *outputs.values()]
        ]
        if len(set(destinations)) != len(destinations):
            raise ValueError("Artifact destinations must be distinct")
        check_source_destinations(
            request,
            destinations,
            object_value(snapshot.prepared.extraction.legacy_state)
            if snapshot is not None
            else None,
        )
        destinations_valid = True
        render_selection = selection or RenderSelection(
            tuple(outputs),
            tuple(object_value(request.config)["report_sections"])
            if "report_sections" in object_value(request.config)
            else None,
        )
        if set(outputs) != set(render_selection.formats):
            raise ValueError("Requested destinations disagree with selected formats")
        if snapshot is None:
            if collect is None:
                raise ValueError("Extraction requires a collector")
            # An existing status destination may itself be an original that
            # the collector has not identified yet. Fresh run folders can
            # publish progress immediately; existing files wait for discovery.
            destinations_valid = not Path(manifest_path).exists()
            stage("collecting")
            records, state = collect()
            check_cancelled(cancelled)
            try:
                check_source_destinations(request, destinations, state)
            except ValueError:
                destinations_valid = False
                raise
            destinations_valid = True
            evidence = capture_legacy(records, state, request)
            stage("preparing")
            prepared = prepare_snapshot(evidence)
            if copy_sources is not None:
                stage("copying_sources")
                updated_state = copy_sources(prepared)
                prepared = replace(
                    prepared,
                    extraction=replace(
                        evidence,
                        legacy_state=freeze_json(
                            to_snapshot_json(updated_state, path="$.legacy_state")
                        ),
                    ),
                )
            stage("analysing")
            snapshot = analyse_prepared(prepared)
        stage("saving_result")
        store(snapshot, result_path, result_id=request.run_id + "-result")
        manifest["result"] = artifact_record(result_path)
        state = object_value(snapshot.prepared.extraction.legacy_state)
        no_evidence = not snapshot.prepared.extraction.observations and not any(
            state.get(key) for key in ("sap_contract_rows", "sap_meter_rows", "sap_financial_rows")
        )
        if no_evidence:
            manifest["status"] = "empty"
        else:
            registry = renderers if renderers is not None else default_renderers()
            view = build_report_view(snapshot, render_selection)
            uncertain = sum(bool(row.get("Transaction Identity Review")) for row in view.records)
            if uncertain:
                manifest["warnings"].append(
                    f"{uncertain} transaction observations have ambiguous or conflicting cross-source identities; review them before relying on aggregate totals."
                )
            claims = view.claims
            if has_active_claims(claims) and claims is not None and not claims["ready"]:
                manifest["warnings"].append(
                    "Claim reconciliation is incomplete; review overlaps, allocations and changed evidence before asserting one total."
                )
            for fmt, path in outputs.items():
                stage("rendering_" + fmt)
                registry[fmt](view, path)
                manifest["artifacts"][fmt] = artifact_record(path)
                save_manifest(manifest, manifest_path)
            from edf_bill_fetcher.application import complaint_pack

            if complaint_pack.requested(view):
                stage("building_complaint_pack")
                pack_paths, pack_notes = complaint_pack.publish(
                    view,
                    str(Path(result_path).parent / "complaint_pack"),
                    source_roots=tuple(
                        str(Path(source.path).parent)
                        for source in request.sources
                        if source.kind == "records_json"
                    ),
                    cancelled=cancelled,
                )
                manifest["warnings"].extend(pack_notes)
                for key, path in pack_paths.items():
                    manifest["artifacts"][key] = artifact_record(path)
                save_manifest(manifest, manifest_path)
            check_cancelled(cancelled)
            manifest["status"] = (
                "review_required"
                if state.get("error_log")
                or manifest["warnings"]
                or any(
                    annotation.get("status") == "needs_information"
                    for annotation in state.get("human_review_statuses", {}).values()
                )
                else "complete"
            )
    except RunCancelled as exc:
        error = str(exc)
        manifest["status"] = "cancelled"
    except Exception as exc:
        error = f"{type(exc).__name__}: {exc}"
        manifest["status"] = "failed"
    finally:
        try:
            cleanup()
        except Exception as exc:
            error = (error + "; " if error else "") + f"Cleanup failed: {exc}"
            manifest["status"] = "failed"
        manifest["error"] = error
        if destinations_valid:
            try:
                save_manifest(manifest, manifest_path)
            except Exception as exc:
                error = (error + "; " if error else "") + f"Run status could not be saved: {exc}"
                manifest["error"] = error
                manifest["status"] = "failed"
    return RunOutcome(manifest["status"], snapshot, freeze_json(manifest), error)


def collect_engine(
    engine: Any, request: RunRequest, *, pff_module: Any = None
) -> tuple[list[dict], dict]:
    """Adapt existing source collectors, closing PST resources on every outcome."""
    for source in request.sources:
        cancel_event = getattr(engine, "cancel_event", None)
        if cancel_event is not None and cancel_event.is_set():
            raise RunCancelled("Run cancelled during collection")
        if source.kind in {"pst", "ost"}:
            if pff_module is None:
                import pypff

                pff_module = pypff

            constructor = getattr(pff_module, "file", None) or getattr(pff_module, "File", None)
            if constructor is None:
                raise RuntimeError("pypff has no file constructor")
            archive = constructor()
            previous_origin = getattr(engine, "_archive_origin", "")
            try:
                archive.open(str(Path(source.path).absolute()))
                engine._archive_origin = str(Path(source.path).absolute())
                engine.crawl_pst(archive.get_root_folder())
            finally:
                engine._archive_origin = previous_origin
                archive.close()
        else:
            method = {
                "htm": "process_htm_file",
                "pdf_dir": "crawl_local_pdfs",
                "eml_dir": "process_eml_folder",
                "msg_dir": "process_msg_folder",
            }.get(source.kind)
            if method is None:
                raise ValueError(f"Unsupported collection source: {source.kind}")
            getattr(engine, method)(source.path)
    return engine.records, engine.__getstate__()


def bundle_engine_sources(
    engine: Any, prepared: PreparedSnapshot, bundle_dir: str, *, legacy_paths: tuple[str, ...] = ()
) -> dict:
    """Copy sources and emit explicit v2 compatibility files from existing preparation."""
    import os

    import pandas as pd

    from edf_bill_fetcher.domain.json_values import rows_value
    from edf_bill_fetcher.helpers.artifacts import write_json
    from edf_bill_fetcher.io.writers.evidence_bundle import (
        build_bundle_index,
        build_bundle_manifest,
        save_evidence_files,
    )

    config = object_value(prepared.effective_config)
    engine.config = config
    state = object_value(prepared.extraction.legacy_state)
    records = rows_value(prepared.records)
    bundle_records = pd.concat(
        [
            pd.DataFrame(records),
            pd.DataFrame(rows_value(prepared.duplicates)),
            pd.DataFrame(state.get("filtered_records", [])),
            pd.DataFrame(list(engine.source_manifest.values())),
        ],
        ignore_index=True,
    )
    saved: dict[str, str] = {}
    stats: dict[str, int] = {}
    if config.get("save_evidence_files", True):
        saved = save_evidence_files(
            bundle_records,
            engine.source_paths,
            str(Path(bundle_dir) / "evidence_files"),
            stats=stats,
            log=lambda text: engine.log_error("Evidence bundle", text),
        )
        build_bundle_manifest(bundle_records, saved, bundle_dir)
        build_bundle_index(
            bundle_records,
            saved,
            str(Path(bundle_dir) / "evidence_index.docx"),
            account=str(config.get("acc_num", "")),
        )
    for key in ("missing", "ambiguous"):
        if stats.get(key):
            engine.log_error("Evidence bundle", f"{stats[key]} {key} source files")
    saved_aliases = dict(saved)
    for source_id, metadata in engine.source_manifest.items():
        copied = saved.get(metadata["Attachment Name"])
        metadata["Saved Path"] = (
            os.path.relpath(copied, bundle_dir).replace(os.sep, "/") if copied else None
        )
        if copied:
            saved_aliases[source_id] = copied
    state.update(to_snapshot_json(engine.__getstate__(), path="$.legacy_state"))
    state.update(
        {
            "saved_evidence_paths": saved_aliases,
            "source_bundle_path": str(Path(bundle_dir).absolute()),
            "original_records": [object_value(row.raw) for row in prepared.extraction.observations],
        }
    )
    compatibility = {
        **state,
        "version": 2,
        "extracted_at": prepared.extraction.request.created_at.isoformat(),
        "config": config,
        "records": records,
        "duplicate_records": rows_value(prepared.duplicates),
        "status": "review_required" if engine.error_log else "complete",
    }
    for path in legacy_paths:
        write_json(compatibility, path)
    return state
