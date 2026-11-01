"""Preview and publish portable review drafts without a Qt or HTTP dependency."""

from __future__ import annotations

from collections.abc import Callable
from datetime import date, datetime, timezone
from pathlib import Path
from uuid import uuid4

from edf_bill_fetcher.application.analysis_service import analyse_prepared
from edf_bill_fetcher.application.legacy_adapter import capture_legacy
from edf_bill_fetcher.application.preparation_service import prepare_snapshot
from edf_bill_fetcher.application.review_session import ReviewSession
from edf_bill_fetcher.application.run_service import RunOutcome, run
from edf_bill_fetcher.application.workbook_review import bundle_review_sources
from edf_bill_fetcher.domain.analysis import AnalysisSnapshot
from edf_bill_fetcher.domain.json_values import freeze_json
from edf_bill_fetcher.domain.requests import RunRequest, SourceSelection
from edf_bill_fetcher.io.persistence.result_store import load_result
from edf_bill_fetcher.models.report_view import RenderSelection


def review_inputs(
    result_path: str, draft: dict, run_id: str | None = None
) -> tuple[RunRequest, list[dict], dict]:
    """Bind the frozen draft to the entire saved analysis and retain pinned settings."""
    session = ReviewSession.restore(load_result(result_path), draft)
    records, state, as_of = session.inputs(result_path)
    request = RunRequest(
        run_id or uuid4().hex,
        datetime.now(timezone.utc),
        date.fromisoformat(as_of),
        (SourceSelection("records_json", str(Path(result_path).resolve())),),
        freeze_json(state["review_config"]),
    )
    return request, records, state


def preview_review(result_path: str, draft: dict, *, run_id: str | None = None) -> AnalysisSnapshot:
    """Recompute through the ordinary services without publishing or changing originals."""
    request, records, state = review_inputs(result_path, draft, run_id)
    return analyse_prepared(prepare_snapshot(capture_legacy(records, state, request)))


def publish_review(
    result_path: str,
    draft: dict,
    directory: str,
    *,
    formats: tuple[str, ...] = ("xlsx", "pdf", "docx", "html"),
    sections: tuple[str, ...] | None = None,
    run_id: str | None = None,
    cancelled: Callable[[], bool] = lambda: False,
    progress: Callable = lambda event: None,
    allowed_source: Callable[[str], bool] = lambda path: True,
) -> RunOutcome:
    """Create a new complete reviewed pack using the shared audited source-copy workflow."""
    request, records, state = review_inputs(result_path, draft, run_id)
    destination = Path(directory)
    destination.mkdir(parents=True, exist_ok=True)
    if any(path.name != "worker.log" for path in destination.iterdir()):
        raise ValueError("Publish reviews into a fresh empty folder")
    formats = tuple(dict.fromkeys(("xlsx", *formats)))
    return run(
        request,
        {
            fmt: str(destination / ("Evidence" + ("_Report" if fmt != "xlsx" else "") + "." + fmt))
            for fmt in formats
        },
        result_path=str(destination / "result.json"),
        manifest_path=str(destination / "run.json"),
        collect=lambda: (records, state),
        copy_sources=lambda prepared: bundle_review_sources(
            prepared, directory, result_path, allowed_source=allowed_source
        ),
        cancelled=cancelled,
        progress=progress,
        selection=RenderSelection(formats, sections),
        allowed_source=allowed_source,
    )
