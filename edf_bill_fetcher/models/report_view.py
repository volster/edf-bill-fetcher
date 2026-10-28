"""Presentation selection and fresh decoded views of accepted analysis."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timezone
from types import SimpleNamespace
from typing import Any
from uuid import uuid4

import pandas as pd

from edf_bill_fetcher.domain.analysis import AnalysisSnapshot
from edf_bill_fetcher.domain.json_values import object_value, rows_value
from edf_bill_fetcher.models.report_sections import REPORT_SECTIONS, RenderContext
from edf_bill_fetcher.models.snapshot_adapter import from_snapshot_json


@dataclass(frozen=True)
class RenderSelection:
    """Absent sections mean defaults; an explicit empty tuple means no sections."""

    formats: tuple[str, ...] = ("xlsx",)
    sections: tuple[str, ...] | None = None

    def __post_init__(self) -> None:
        """Reject mutable selections, duplicate or unknown choices."""
        if (
            type(self.formats) is not tuple
            or len(set(self.formats)) != len(self.formats)
            or any(value not in {"xlsx", "pdf", "docx", "html"} for value in self.formats)
        ):
            raise ValueError("Invalid output formats")
        allowed = {section.key for section in REPORT_SECTIONS} | {"cover", "toc"}
        if self.sections is not None and (
            type(self.sections) is not tuple
            or len(set(self.sections)) != len(self.sections)
            or any(value not in allowed for value in self.sections)
        ):
            raise ValueError("Invalid report sections")


@dataclass(frozen=True)
class ReportView:
    """No computation: each access returns detached presentation objects."""

    snapshot: AnalysisSnapshot
    selection: RenderSelection

    @property
    def ledger_overview(self) -> list[dict]:
        """Return a stored optional overview; old results never gain recomputed figures."""
        return list(self.model("sap").get("ledger_overview", []))

    @property
    def config(self) -> dict:
        """Return a detached effective configuration with explicit presentation choices."""
        cfg = object_value(self.snapshot.prepared.effective_config)
        cfg["report_sections"] = list(self.context.selected)
        return cfg

    @property
    def context(self) -> RenderContext:
        """Resolve registry numbering for the selected sections."""
        selected = self.selection.sections
        return RenderContext(
            ({s.key for s in REPORT_SECTIONS} | {"cover", "toc"})
            if selected is None
            else list(selected)
        )

    @property
    def records(self) -> list[dict]:
        """Return fresh canonical rows."""
        return rows_value(self.snapshot.prepared.records)

    @property
    def duplicates(self) -> list[dict]:
        """Return fresh duplicate decisions."""
        return rows_value(self.snapshot.prepared.duplicates)

    @property
    def review_notice(self) -> str | None:
        """Disclose human corrections and exclusions in every regenerated report."""
        state = object_value(self.snapshot.prepared.extraction.legacy_state)
        if "human_review" not in state:
            return None
        excluded = len(state.get("review_excluded_records", []))
        changes = len(state.get("review_history", []))
        statuses = list(state.get("human_review_statuses", {}).values())
        verified = sum(row.get("status") == "verified" for row in statuses)
        needs_info = sum(row.get("status") == "needs_information" for row in statuses)
        return (
            f"Human-reviewed analysis: {changes} recorded changes and review decisions; "
            f"{excluded} records currently excluded; {verified} verified against source; {needs_info} need information. The workbook Review history and result.json "
            "preserve reasons and original extraction data. Shared source PDFs remain whole; "
            "record exclusions do not redact pages."
        )

    @property
    def ocr_notice(self) -> str | None:
        """Keep OCR limitations visible even when optional report sections are hidden."""
        manifest = object_value(self.snapshot.prepared.extraction.legacy_state).get(
            "source_manifest", {}
        )
        used_ocr = any(row.get("OCR Review Required") for row in self.records)
        if isinstance(manifest, dict):
            used_ocr = used_ocr or any(
                isinstance(source, dict) and source.get("OCR Page Results")
                for source in manifest.values()
            )
        if not used_ocr:
            return None
        return (
            "OCR review required: some source pages required text recognition. "
            "Verify invoice references, dates, periods and amounts against the original PDFs. "
            "Confidence scores do not establish accuracy; unreadable pages may leave gaps."
        )

    @property
    def engine(self) -> SimpleNamespace:
        """Expose data-only legacy metadata to the compatibility layout builders."""
        state = {
            "pdf_count": 0,
            "email_count": 0,
            "error_log": [],
            "filtered_records": [],
            "source_manifest": {},
            **object_value(self.snapshot.prepared.extraction.legacy_state),
        }
        return SimpleNamespace(**state)

    def model(self, key: str) -> Any:
        """Decode a fixed model without fitting or financial processing."""
        return from_snapshot_json(object_value(self.snapshot.models)[key]["payload"])

    @property
    def revisions(self) -> dict | None:
        """Read the stored revision assessment; older accepted results remain unchanged."""
        result = self.model("sap").get("bill_revisions")
        return result if isinstance(result, dict) else None

    @property
    def rebilling_ledger(self) -> dict | None:
        """Return accepted history and dispute positions without recalculation."""
        result = self.model("sap").get("rebilling_ledger")
        return result if isinstance(result, dict) else None

    @property
    def claims(self) -> dict | None:
        """Read accepted claim reconciliation without recomputing proposals."""
        result = self.model("sap").get("claim_reconciliation")
        return result if isinstance(result, dict) else None

    def frame(self, key: str = "evidence") -> pd.DataFrame:
        """Return a new precomputed frame stored with the analytical payload."""
        return self.model("sap")["frames"][key]


def build_report_view(snapshot: AnalysisSnapshot, selection: RenderSelection) -> ReportView:
    """Build a presentation view without changing its financial fingerprint."""
    return ReportView(snapshot, selection)


def capture_report_view(
    records: list[dict],
    config: dict,
    engine: Any = None,
    filtered: list | None = None,
    *,
    formats: tuple[str, ...],
    sections: tuple[str, ...] | None = None,
) -> ReportView:
    """Compatibility boundary: capture and compute once before delegating to rendering."""
    from edf_bill_fetcher.application.analysis_service import analyse_prepared
    from edf_bill_fetcher.application.legacy_adapter import capture_legacy
    from edf_bill_fetcher.application.preparation_service import prepare_snapshot
    from edf_bill_fetcher.domain.json_values import freeze_json
    from edf_bill_fetcher.domain.requests import RunRequest

    if not records and not any(
        isinstance(getattr(engine, key, None), list) and getattr(engine, key)
        for key in ("sap_contract_rows", "sap_meter_rows", "sap_financial_rows")
    ):
        raise ValueError("No records to report on")
    cfg = dict(config)
    cfg.setdefault("as_of", date.today().isoformat())
    if sections is None and "report_sections" in cfg:
        sections = tuple(cfg["report_sections"])
    state = (
        {}
        if engine is None
        else (engine.__getstate__() if hasattr(engine, "__getstate__") else vars(engine))
    )
    state = dict(state or {})
    state["filtered_records"] = (
        filtered if filtered is not None else state.get("filtered_records", [])
    )
    originals = state.get("original_records", records)
    request = RunRequest(
        uuid4().hex,
        datetime.now(timezone.utc),
        date.fromisoformat(cfg["as_of"]),
        (),
        freeze_json(cfg),
    )
    result = analyse_prepared(prepare_snapshot(capture_legacy(originals, state, request)))
    return build_report_view(result, RenderSelection(formats, sections))
