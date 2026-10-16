"""Immutable desktop jobs and a Qt worker around the shared run coordinator."""

from dataclasses import dataclass, replace
from datetime import date
from functools import partial
from pathlib import Path
from threading import Event
from typing import cast

from PySide6.QtCore import QObject, QThread, Signal

from edf_bill_fetcher.application.run_service import (
    RunCancelled,
    RunOutcome,
    bundle_engine_sources,
    check_cancelled,
    collect_engine,
    default_renderers,
    run,
)
from edf_bill_fetcher.domain.json_values import freeze_json, object_value
from edf_bill_fetcher.domain.requests import RunRequest
from edf_bill_fetcher.io.persistence.result_store import load_result
from edf_bill_fetcher.io.persistence.run_manifest import save_manifest
from edf_bill_fetcher.models.config import ConfigDict
from edf_bill_fetcher.models.report_view import RenderSelection


@dataclass(frozen=True)
class DesktopJob:
    """All worker inputs captured before any background work begins."""

    request: RunRequest
    directory: str
    stem: str
    formats: tuple[str, ...]
    sections: tuple[str, ...]
    mode: str = "extract"
    input_path: str = ""

    @property
    def outputs(self) -> dict[str, str]:
        """Use reserved run folders to preserve earlier output artifacts."""
        return {
            fmt: str(
                Path(self.directory)
                / (self.stem + ("_Report" if fmt != "xlsx" else "") + "." + fmt)
            )
            for fmt in self.formats
        }


def read_workbook(path: str) -> tuple[list[dict], dict, str | None]:
    """Use the shared Excel review/legacy import adapter."""
    from edf_bill_fetcher.io.review_workbook import read_workbook as read

    return read(path)


class RunWorker(QThread):
    """Emit Qt signals; never read or mutate widgets from the worker thread."""

    progress = Signal(str)
    percentage = Signal(int)
    outcome_ready = Signal(object)

    def __init__(self, job: DesktopJob, parent: QObject | None = None) -> None:
        """Own one immutable job and a cooperative cancellation token."""
        super().__init__(parent)
        self.job = job
        self.cancel_event = Event()

    def run(self) -> None:
        """Load/extract once and preserve the coordinator's truthful terminal outcome."""
        from edf_bill_fetcher.collectors.engine import EvidenceEngine

        job, request = self.job, self.job.request
        engine = None
        result_path = str(Path(job.directory) / "result.json")
        manifest_path = str(Path(job.directory) / "run.json")
        try:
            check_cancelled(self.cancel_event.is_set)
            accepted = None
            imported = None
            renderers = None
            warnings: tuple[str, ...] = ()
            if job.mode == "saved_result":
                self.progress.emit("Opening saved analysis…")
                accepted = load_result(job.input_path)
                from edf_bill_fetcher.io.writers.evidence_bundle import resolve_bundle_paths
                from edf_bill_fetcher.io.writers.export import render_xlsx

                paths, warnings = resolve_bundle_paths(
                    object_value(accepted.prepared.extraction.legacy_state),
                    str(Path(job.input_path).parent),
                )
                renderers = default_renderers()
                renderers["xlsx"] = partial(render_xlsx, evidence_paths=paths)
                cfg = object_value(accepted.prepared.effective_config)
                request = replace(
                    request,
                    as_of=accepted.prepared.extraction.request.as_of,
                    config=freeze_json(cfg),
                )
            elif job.mode == "workbook":
                self.progress.emit("Reading workbook…")
                records, state, as_of = read_workbook(job.input_path)
                cfg = state.get("review_config", object_value(request.config))
                cfg["as_of"] = as_of or request.as_of.isoformat()
                request = replace(
                    request, as_of=date.fromisoformat(cfg["as_of"]), config=freeze_json(cfg)
                )
                imported = (records, state)

            def collect():
                nonlocal engine
                if imported is not None:
                    return imported
                engine = EvidenceEngine(
                    cast(ConfigDict, object_value(request.config)),
                    self.progress.emit,
                    self._progress,
                    self.cancel_event,
                )
                return collect_engine(engine, request)

            def cleanup():
                if engine is not None:
                    engine.cleanup()

            def bundle(prepared):
                if imported is not None and "human_review" in imported[1]:
                    from edf_bill_fetcher.application.workbook_review import bundle_review_sources

                    return bundle_review_sources(prepared, job.directory, job.input_path)
                return bundle_engine_sources(
                    engine,
                    prepared,
                    job.directory,
                    legacy_paths=(str(Path(job.directory) / "records.json"),),
                )

            outcome = run(
                request,
                job.outputs,
                result_path=result_path,
                manifest_path=manifest_path,
                collect=collect,
                copy_sources=bundle
                if job.mode == "extract" or (imported is not None and "human_review" in imported[1])
                else None,
                cleanup=cleanup,
                cancelled=self.cancel_event.is_set,
                progress=lambda event: self.progress.emit(
                    {
                        "collecting": "Reading your sources…",
                        "preparing": "Organising the evidence…",
                        "copying_sources": "Preserving original files…",
                        "analysing": "Analysing your bills…",
                        "saving_result": "Saving the analysis…",
                    }.get(
                        event.stage,
                        "Creating " + event.stage.removeprefix("rendering_").upper() + "…",
                    )
                ),
                accepted_result=accepted,
                selection=RenderSelection(job.formats, job.sections),
                renderers=renderers,
                warnings=warnings,
            )
        except Exception as error:
            status = "cancelled" if isinstance(error, RunCancelled) else "failed"
            manifest: dict = {
                "manifest_version": 1,
                "run_id": request.run_id,
                "status": status,
                "stage": "loading_input",
                "requested": job.outputs,
                "artifacts": {},
                "result": None,
                "error": str(error),
            }
            try:
                save_manifest(manifest, manifest_path)
            except OSError as publication_error:
                manifest["error"] += f"; Run status could not be saved: {publication_error}"
            outcome = RunOutcome(status, None, freeze_json(manifest), str(manifest["error"]))
        self.outcome_ready.emit(outcome)

    def _progress(self, current, total, text=None):
        if text:
            self.progress.emit(str(text))
        self.percentage.emit(int(max(0, min(100, current / total * 100))) if total else 0)
