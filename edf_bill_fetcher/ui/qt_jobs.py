"""Immutable desktop jobs and a Qt worker around the shared run coordinator."""

from dataclasses import dataclass, replace
from datetime import date
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
    """Import workbook rows with an explicit warning about lost extraction provenance."""
    import pandas as pd

    from edf_bill_fetcher.models.snapshot_adapter import to_snapshot_json

    with pd.ExcelFile(path) as workbook:
        records = pd.read_excel(workbook, sheet_name="EDF Evidence Report").to_dict("records")
        metadata = {}
        if "Provenance" in workbook.sheet_names:
            frame = pd.read_excel(workbook, sheet_name="Provenance", header=None)
            metadata = {
                str(row[0]): to_snapshot_json(row[1])
                for row in frame.itertuples(index=False, name=None)
                if len(row) > 1 and pd.notna(row[0])
            }
        filtered = (
            pd.read_excel(workbook, sheet_name="Filtered (Below Min)").to_dict("records")
            if "Filtered (Below Min)" in workbook.sheet_names
            else []
        )
    return (
        records,
        {
            "records": records,
            "filtered_records": filtered,
            "pdf_count": 0,
            "email_count": 0,
            "source_manifest": {},
            "imported_workbook": path,
            "imported_workbook_metadata": metadata,
            "provenance": "lossy_workbook_import",
            "error_log": [
                "Workbook import: original observations, source files and analytical decisions are unavailable."
            ],
        },
        str(metadata["as_of"]) if metadata.get("as_of") else None,
    )


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
            if job.mode == "saved_result":
                self.progress.emit("Opening saved analysis…")
                accepted = load_result(job.input_path)
                cfg = object_value(accepted.prepared.effective_config)
                request = replace(
                    request,
                    as_of=accepted.prepared.extraction.request.as_of,
                    config=freeze_json(cfg),
                )
            elif job.mode == "workbook":
                self.progress.emit("Reading workbook…")
                records, state, as_of = read_workbook(job.input_path)
                cfg = object_value(request.config)
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
                copy_sources=bundle if job.mode == "extract" else None,
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
