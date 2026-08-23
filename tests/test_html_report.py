"""Tests for the HTML report renderer (``io/reporters/html_report.py``).

The HTML report is a third output surface next to the PDF and DOCX
generators.  It builds its document from the same ``REPORT_SECTIONS``
registry and ``RenderContext``, so numbering and the table of contents
always agree with the other two surfaces.

The acceptance tests below pin the task contract:

* rendering with a small records fixture produces a file;
* the file contains the heading text for every *selected* section;
* the file contains the project disclaimer;
* the file does NOT contain the heading text of an *unselected* section;
* a registry key with no wired HTML builder raises a loud RuntimeError,
  mirroring the PDF dispatcher's failure mode.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd
import pytest

from edf_bill_fetcher.helpers.version import get_package_version
from edf_bill_fetcher.io.reporters.pdf_report import RenderContext, ReportSectionMeta
from edf_bill_fetcher.models.config import ConfigDict


@pytest.fixture
def sample_records() -> list[dict[str, Any]]:
    """A small, self-consistent records fixture (4 rows, 3 sources)."""
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
            "Date": "15/02/2023",
            "Period From": "N/A",
            "Period To": "N/A",
            "Amount (£)": 100.00,
            "Source": "Local PDF Folder",
            "Entry Type": "Payment",
            "Invoice #": "N/A",
            "Details": "Direct debit",
            "Reading": "Unknown",
            "Units (kWh)": None,
            "Period Charge (£)": None,
            "Unit Rate (p/kWh)": None,
            "Tariff": None,
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


@pytest.fixture
def sample_config() -> ConfigDict:
    """Configuration mirroring the PDF/DOCX report tests."""
    return {
        "min_amount": 50.0,
        "analysis_min": 500.0,
        "acc_num": "123456789",
        "report_account_ref": "ACC-123456",
        "use_anchors": True,
        "use_large": True,
        "use_reading_classification": True,
        "use_pdf_fields": True,
        "use_acc_filter": True,
        "filter_below": False,
        "save_filtered": True,
        "use_dedup": True,
        "save_dups": True,
        "use_domain_filter": True,
        "domain_filter": "edfenergy.com",
        "report_sections": [
            "cover",
            "toc",
            "exec_summary",
            "key_findings",
            "evidence_index",
            "timeline",
            "data_quality",
            "appendix_glossary",
        ],
    }


@pytest.fixture
def rendered_html(
    tmp_path: Path, sample_records: list[dict[str, Any]], sample_config: ConfigDict
) -> str:
    """Render an HTML report with the default selected sections."""
    from edf_bill_fetcher.io.reporters.html_report import generate_html_report

    out_path = tmp_path / "report.html"
    generate_html_report(
        records=sample_records,
        output_path=str(out_path),
        config=sample_config,
    )
    return out_path.read_text(encoding="utf-8")


def test_html_report_file_exists(
    tmp_path: Path, sample_records: list[dict[str, Any]], sample_config: ConfigDict
) -> None:
    """Rendering a small fixture produces an .html file on disk."""
    from edf_bill_fetcher.io.reporters.html_report import generate_html_report

    out_path = tmp_path / "report.html"
    generate_html_report(
        records=sample_records,
        output_path=str(out_path),
        config=sample_config,
    )
    assert out_path.exists()
    assert out_path.stat().st_size > 0


def test_html_report_contains_selected_section_headings(rendered_html: str) -> None:
    """Every selected section's heading appears (registry numbering)."""
    # exec_summary, key_findings, evidence_index, timeline, data_quality
    # are the five main sections selected; numbering is derived from the
    # registry at render time.
    for expected in (
        "1. Executive Summary",
        "2. Key Findings Summary",
        "3. Evidence Index",
        "4. Timeline of Events",
        "5. Data Quality Assessment",
        "A. Glossary",
    ):
        assert expected in rendered_html, f"missing heading: {expected}"


def test_html_report_contains_toc_entries(rendered_html: str) -> None:
    """The Table of Contents lists the selected sections with labels."""
    # HTML-escaped ampersand inside the registry title.
    assert "Evidence Index &amp; Source Cross-Reference" in rendered_html
    assert "Table of Contents" in rendered_html


def test_html_report_contains_disclaimer(rendered_html: str) -> None:
    """The project disclaimer is present in the report."""
    assert "provided as-is without warranty" in rendered_html


def test_html_report_omits_unselected_section(rendered_html: str) -> None:
    """A section left out of the selection must not appear."""
    # "statistical" / "tariff" are not in the fixture's report_sections.
    assert "Statistical Analysis" not in rendered_html
    assert "Tariff Impact Analysis" not in rendered_html


def test_html_report_cover_carries_package_version(rendered_html: str) -> None:
    """The cover shows the version declared in pyproject.toml."""
    version = get_package_version()
    assert f"EDF Evidence Collector v{version}" in rendered_html


def test_html_report_is_offline_only(rendered_html: str) -> None:
    """No external stylesheets or remote assets — inline CSS only."""
    assert "<style>" in rendered_html
    assert "<link" not in rendered_html.lower()
    assert re.search(r'src="https?://', rendered_html) is None
    assert re.search(r'href="https?://', rendered_html) is None


def test_report_sections_argument_overrides_config(
    tmp_path: Path, sample_records: list[dict[str, Any]], sample_config: ConfigDict
) -> None:
    """The explicit ``report_sections`` argument wins over config."""
    from edf_bill_fetcher.io.reporters.html_report import generate_html_report

    out_path = tmp_path / "report2.html"
    generate_html_report(
        records=sample_records,
        output_path=str(out_path),
        config=sample_config,
        report_sections=["exec_summary", "tariff"],
    )
    html = out_path.read_text(encoding="utf-8")
    assert "1. Executive Summary" in html
    assert "2. Tariff Impact Analysis" in html
    assert "Key Findings Summary" not in html


def test_placeholder_sections_render_note(
    tmp_path: Path, sample_records: list[dict[str, Any]], sample_config: ConfigDict
) -> None:
    """Chart-heavy sections render the 'not implemented in HTML' note."""
    from edf_bill_fetcher.io.reporters.html_report import generate_html_report

    out_path = tmp_path / "report3.html"
    generate_html_report(
        records=sample_records,
        output_path=str(out_path),
        config=sample_config,
        report_sections=["statistical", "payment", "forecast"],
    )
    html = out_path.read_text(encoding="utf-8")
    assert html.count("not implemented in HTML") >= 3
    assert "Statistical Analysis" in html
    assert "Payment &amp; Credit Analysis" in html
    assert "Forecast &amp; Projection" in html


def test_missing_builder_raises_runtime_error(
    tmp_path: Path,
    sample_records: list[dict[str, Any]],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A registry key with no HTML builder must fail loudly (PDF parity)."""
    import edf_bill_fetcher.io.reporters.html_report as html_module
    from edf_bill_fetcher.io.reporters.html_report import generate_html_report

    original = html_module.REPORT_SECTIONS
    try:
        monkeypatch.setattr(
            html_module,
            "REPORT_SECTIONS",
            [
                *original,
                ReportSectionMeta(key="__test_orphan_section__", title="Orphan Test Section"),
            ],
        )
        with pytest.raises(RuntimeError) as exc_info:
            generate_html_report(
                records=sample_records,
                output_path=str(tmp_path / "report4.html"),
                config={"report_sections": ["__test_orphan_section__"]},
            )
        assert "__test_orphan_section__" in str(exc_info.value)
    finally:
        html_module.REPORT_SECTIONS = original


def test_generate_html_from_gui_success(
    tmp_path: Path, sample_records: list[dict[str, Any]], sample_config: ConfigDict
) -> None:
    """The GUI wrapper returns ``(True, path)`` on success."""
    from edf_bill_fetcher.io.reporters.html_report import generate_html_from_gui

    out_path = tmp_path / "report5.html"
    ok, message = generate_html_from_gui(
        records=sample_records,
        output_path=str(out_path),
        config=sample_config,
    )
    assert ok is True
    assert str(out_path) in message
    assert out_path.exists()


def test_generate_html_from_gui_failure(tmp_path: Path) -> None:
    """The GUI wrapper returns ``(False, message)`` instead of raising."""
    from edf_bill_fetcher.io.reporters.html_report import generate_html_from_gui

    ok, message = generate_html_from_gui(
        records=[],
        output_path=str(tmp_path / "report7.html"),
        config={},
    )
    assert ok is False
    assert "No records to report on" in message


def test_appendices_lettered_in_toc(rendered_html: str) -> None:
    """Selected appendices keep alphabetic labels (A., B., ...)."""
    assert "A. Glossary" in rendered_html


def test_glossary_table_renders_terms(rendered_html: str) -> None:
    """The glossary appendix shows real terms, not a placeholder."""
    assert "OFGEM Price Cap" in rendered_html
    assert "Period Charge (£)" in rendered_html


def test_render_context_still_in_registry_order() -> None:
    """Sanity: the shared RenderContext numbers a fresh selection the
    same way the HTML dispatcher consumes it (registry order)."""
    ctx = RenderContext(
        ["exec_summary", "key_findings", "evidence_index", "timeline", "data_quality"]
    )
    labels = [spec.label for spec in ctx.sections_in_order]
    assert labels == ["1.", "2.", "3.", "4.", "5."]


def test_records_dataframe_roundtrip(sample_records: list[dict[str, Any]]) -> None:
    """The records fixture survives the DataFrame conversion used
    internally by the dispatcher."""
    df = pd.DataFrame(sample_records)
    assert len(df) == 4
    assert set(df["Source"].unique()) == {
        "HTM Account History",
        "PST PDF Attachment",
        "Local PDF Folder",
    }


def test_evidence_file_path_contract() -> None:
    """The evidence file for this task lives at the documented path."""
    evidence = (
        Path(__file__).resolve().parents[1]
        / ".omo/evidence/post-release-feature-wave/task-6f1-2026-08-06-post-release-feature-wave.txt"
    )
    assert evidence.exists(), "task evidence file is missing"
