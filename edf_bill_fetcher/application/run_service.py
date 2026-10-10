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
    }
    error = None
    destinations_valid = False

    def stage(name: str) -> None:
        check_cancelled(cancelled)
        manifest["stage"] = name
        save_manifest(manifest, manifest_path)
        progress(ProgressEvent(name, name.replace("_", " ").title()))

    try:
        destinations = [
            Path(path).resolve() for path in [result_path, manifest_path, *outputs.values()]
        ]
        if len(set(destinations)) != len(destinations):
            raise ValueError("Artifact destinations must be distinct")
        if any(Path(source.path).resolve() in destinations for source in request.sources):
            raise ValueError("Artifact destination would overwrite a selected source")
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
            stage("collecting")
            records, state = collect()
            check_cancelled(cancelled)
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
            for fmt, path in outputs.items():
                stage("rendering_" + fmt)
                registry[fmt](view, path)
                manifest["artifacts"][fmt] = artifact_record(path)
                save_manifest(manifest, manifest_path)
            check_cancelled(cancelled)
            manifest["status"] = "review_required" if state.get("error_log") else "complete"
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
            save_manifest(manifest, manifest_path)
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
            try:
                archive.open(str(Path(source.path).absolute()))
                engine.crawl_pst(archive.get_root_folder())
            finally:
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
    from edf_bill_fetcher.io.writers.evidence_bundle import build_bundle_index, save_evidence_files

    config = object_value(prepared.effective_config)
    engine.config = config
    state = object_value(prepared.extraction.legacy_state)
    records = rows_value(prepared.records)
    bundle_records = pd.concat(
        [pd.DataFrame(records), pd.DataFrame(list(engine.source_manifest.values()))],
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
        )
        build_bundle_index(
            bundle_records,
            saved,
            str(Path(bundle_dir) / "evidence_index.docx"),
            account=str(config.get("acc_num", "")),
        )
    for key in ("missing", "ambiguous"):
        if stats.get(key):
            engine.log_error("Evidence bundle", f"{stats[key]} {key} source files")
    for metadata in engine.source_manifest.values():
        copied = saved.get(metadata["Attachment Name"])
        metadata["Saved Path"] = os.path.relpath(copied, bundle_dir) if copied else None
    state.update(to_snapshot_json(engine.__getstate__(), path="$.legacy_state"))
    state.update(
        {
            "saved_evidence_paths": saved,
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
