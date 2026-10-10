"""Report selection and dispatcher contracts exercised through actual renderers."""

from __future__ import annotations

import importlib
import json
from datetime import date, datetime, timezone
from functools import wraps
from pathlib import Path

import pandas as pd
import pytest

from edf_bill_fetcher.domain.json_values import freeze_json
from edf_bill_fetcher.domain.requests import RunRequest
from edf_bill_fetcher.io.persistence.result_codec import migrate_legacy
from edf_bill_fetcher.models.report_sections import REPORT_SECTIONS
from edf_bill_fetcher.models.report_view import RenderSelection, build_report_view

BUILDERS = {
    "exec_summary": "create_executive_summary",
    "key_findings": "create_key_findings_table",
    "evidence_index": "create_evidence_index",
    "detailed_findings": "create_anomaly_detail_section",
    "timeline": "create_timeline_section",
    "ofgem": "create_ofgem_comparison",
    "statistical": "create_statistical_analysis",
    "payment": "create_payment_analysis",
    "forecast": "create_forecast_section",
    "data_quality": "create_data_quality_section",
    "tariff": "create_tariff_impact_section",
    "compensation": "create_compensation_section",
    "appendix_methodology": "create_appendix_methodology",
    "appendix_glossary": "create_appendix_glossary",
    "appendix_full_evidence": "create_appendix_full_evidence",
}


@pytest.mark.parametrize("fmt", ["pdf", "docx", "html"])
@pytest.mark.parametrize("sections", [None, ("exec_summary", "payment", "compensation"), ()])
def test_selected_registry_sections_are_rendered_with_stored_payloads(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, fmt: str, sections: tuple[str, ...] | None
) -> None:
    assert set(BUILDERS) == {section.key for section in REPORT_SECTIONS}
    module = importlib.import_module("edf_bill_fetcher.io.reporters." + fmt + "_report")
    request = RunRequest(
        "dispatch",
        datetime(2024, 4, 1, tzinfo=timezone.utc),
        date(2024, 4, 1),
        (),
        freeze_json({"as_of": "2024-04-01"}),
    )
    snapshot = migrate_legacy(_synthetic_html_records(), request)
    view = build_report_view(snapshot, RenderSelection((fmt,), sections))
    observed = []

    def recorder(key, original):
        @wraps(original)
        def invoke(*args, **kwargs):
            observed.append(key)
            if key in {
                "ofgem",
                "statistical",
                "payment",
                "forecast",
                "data_quality",
                "tariff",
                "compensation",
            }:
                assert "payload" in kwargs
            return original(*args, **kwargs)

        return invoke

    for key, name in BUILDERS.items():
        monkeypatch.setattr(module, name, recorder(key, getattr(module, name)))
    output = tmp_path / ("report." + fmt)
    getattr(module, "render_" + fmt)(view, str(output))
    expected = [
        section.key for section in REPORT_SECTIONS if sections is None or section.key in sections
    ]
    assert observed == expected
    assert output.stat().st_size > 0


def _synthetic_html_records() -> list[dict]:
    """Small self-consistent records fixture for the CLI HTML smoke test.

    Mirrors the fixture ``tests/test_html_report.py`` drives the
    renderer with; all identifiers are fabricated (no real EDF data).
    """
    return [
        {
            "Date": "01/01/2023",
            "Period From": "01/11/2022",
            "Period To": "31/01/2023",
            "Amount (£)": 150.50,
            "Source": "HTM Account History",
            "Entry Type": "New Bill",
            "Invoice #": "INV-001",
            "Details": "Standard bill",
            "Reading": "Actual",
            "Units (kWh)": 500,
            "Period Charge (£)": 150.50,
            "Unit Rate (p/kWh)": 30.10,
            "Tariff": "Standard Variable",
        },
        {
            "Date": "01/02/2023",
            "Period From": "01/01/2023",
            "Period To": "28/02/2023",
            "Amount (£)": 200.00,
            "Source": "PST PDF Attachment",
            "Entry Type": "New Bill",
            "Invoice #": "INV-002",
            "Details": "High usage",
            "Reading": "Estimated",
            "Units (kWh)": 600,
            "Period Charge (£)": 200.00,
            "Unit Rate (p/kWh)": 33.33,
            "Tariff": "Standard Variable",
        },
        {
            "Date": "01/03/2023",
            "Period From": "01/02/2023",
            "Period To": "31/03/2023",
            "Amount (£)": 180.00,
            "Source": "HTM Account History",
            "Entry Type": "New Bill",
            "Invoice #": "INV-003",
            "Details": "Normal bill",
            "Reading": "Smart",
            "Units (kWh)": 550,
            "Period Charge (£)": 180.00,
            "Unit Rate (p/kWh)": 32.73,
            "Tariff": "Standard Variable",
        },
    ]


class TestCliHtmlReportSmoke:
    """``--html-report`` runs the real HTML generator headlessly and
    writes an HTML file — the same smoke surface the PDF/DOCX CLI
    entry points expose."""

    def test_cli_html_report_smoke_produces_file(self, tmp_path: Path) -> None:
        from edf_bill_fetcher.io.cli import run_cli_html_report

        records_json = tmp_path / "records.json"
        records_json.write_text(json.dumps(_synthetic_html_records()), encoding="utf-8")
        out_html = tmp_path / "report.html"

        with pytest.raises(SystemExit) as exc:
            run_cli_html_report(["-i", str(records_json), "-o", str(out_html)])
        assert exc.value.code == 0
        assert out_html.exists()
        rendered = out_html.read_text(encoding="utf-8")
        assert "<html" in rendered
        assert "EDF Energy Ombudsman Evidence Report" in rendered


class TestGuiHtmlFormatPresence:
    """The Report Options dialog's Output Format frame exposes HTML."""

    def test_gui_report_options_dialog_exposes_html_format(self) -> None:
        pytest.importorskip("tkinter")
        import tkinter as tk
        from tkinter import ttk

        from edf_bill_fetcher.ui.app import ReportOptionsDialog

        root = tk.Tk()
        root.withdraw()
        try:
            dlg = ReportOptionsDialog(root)
            dlg.dialog = tk.Toplevel(root)
            dlg.dialog.withdraw()
            dlg._build_ui()

            radio_values: list[str] = []

            def _walk(widget: tk.Misc) -> None:
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Radiobutton):
                        radio_values.append(str(child.cget("value")))
                    _walk(child)

            _walk(dlg.dialog)
            assert "html" in radio_values
        finally:
            root.destroy()


class TestFmtDateParity:
    """Parity test for ``fmt_date`` between the PDF and DOCX generators.

    Phase 1.2: ``edf_report_docx.fmt_date`` is an alias of
    ``edf_report.fmt_date`` so a user's bill is rendered
    identically in both formats.  This test pins that invariant at
    the unit level — any code path that re-introduces a
    DOCX-local ``fmt_date`` will break the alias resolution and
    fail the import parity check.
    """

    @pytest.mark.parametrize(
        "input_val",
        [
            None,
            "N/A",
            "NA",
            "",
            "2023-06-15",
            "15/06/2023",
            "15 Jun 2023",
            "15 June 2023",
            "2026-Q3 2026-07-01",
            pd.NaT,
        ],
    )
    def test_fmt_date_pdf_matches_docx_for_every_input(self, input_val):
        from edf_bill_fetcher.io.reporters.docx_report import fmt_date as docx_fmt_date
        from edf_bill_fetcher.io.reporters.pdf_report import fmt_date as pdf_fmt_date

        # Pandas NaT can't be compared with ``==`` cleanly (raises
        # warnings), so we route through ``isinstance`` + ``pd.isna``
        # to short-circuit the equality branch.
        if isinstance(input_val, type(pd.NaT)) or pd.isna(input_val):
            assert pdf_fmt_date(input_val) == docx_fmt_date(input_val)
            return
        assert pdf_fmt_date(input_val) == docx_fmt_date(input_val), (
            f"PDF and DOCX fmt_date disagree on {input_val!r}: "
            f"PDF={pdf_fmt_date(input_val)!r} "
            f"DOCX={docx_fmt_date(input_val)!r}"
        )

    def test_fmt_date_returns_blank_for_missing(self):
        """The blank-for-missing contract is new in Phase 1.2 — explicitly
        tested so ``"Unknown"`` (the old DOCX-only string) cannot sneak
        back in.
        """
        from edf_bill_fetcher.io.reporters.pdf_report import fmt_date

        for missing in (None, "", "N/A", "NA", pd.NaT):
            assert fmt_date(missing) == "", (
                f"fmt_date({missing!r}) returned {fmt_date(missing)!r}; expected blank string"
            )

    def test_fmt_date_renders_iso_and_uk_to_same_dd_mm_yyyy(self):
        from edf_bill_fetcher.io.reporters.pdf_report import fmt_date

        # Both inputs collapse to the same canonical dd/mm/yyyy form,
        # matching the convention the Excel export already uses.
        assert fmt_date("2023-06-15") == "15/06/2023"
        assert fmt_date("15/06/2023") == "15/06/2023"

    def test_docx_uses_edf_report_fmt_date_not_local_definition(self):
        """A docx-local ``fmt_date`` would diverge from the PDF and break
        this project.  The import in ``docx_report`` must be the
        one in ``pdf_report``.
        """
        from edf_bill_fetcher.io.reporters.docx_report import fmt_date as docx_fmt_date
        from edf_bill_fetcher.io.reporters.pdf_report import fmt_date as pdf_fmt_date

        assert docx_fmt_date is pdf_fmt_date, (
            "DOCX fmt_date is a different object; the two renderers "
            "can drift apart at runtime, which is exactly what Phase 1.2 "
            "was meant to prevent."
        )
