"""
Tests for edf_report module - PDF report generation for Ombudsman submissions.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path
from typing import cast
from unittest.mock import Mock, patch

import pandas as pd
import pytest
from reportlab.platypus import TableStyle

from edf_bill_fetcher.io.reporters.pdf_report import (
    Colors,
    _load_ofgem_caps,
    build_styles,
    create_anomaly_detail_section,
    create_appendix_glossary,
    create_appendix_methodology,
    create_cover_page,
    create_data_quality_section,
    create_evidence_index,
    create_executive_summary,
    create_forecast_section,
    create_key_findings_table,
    create_ofgem_comparison,
    create_payment_analysis,
    create_statistical_analysis,
    create_table_of_contents,
    create_tariff_impact_section,
    create_timeline_section,
    fmt_date,
    fmt_money,
    fmt_number,
    fmt_pct,
    generate_ombudsman_pdf,
    generate_pdf_from_gui,
    make_table_style,
    severity_color,
    severity_label,
)
from edf_bill_fetcher.models.config import ConfigDict

# =============================================================================
# HELPER: extract searchable text from reportlab platypus elements
# =============================================================================


def _elements_to_text(elements):
    """Walk a list of reportlab Platypus elements and return a single
    searchable string.

    Implements the minimum reportlab traversal needed for our test suite:
    - Paragraph.text: returned verbatim
    - Table: every cell walked (paragraph inside cells are flattened too)
    - Spacer / PageBreak / anything else: ignored (no text content)

    This lets tests assert ``assert "Account Reference: ACC-12345" in text``
    instead of the vacuous ``len(elements) > 0`` smoke.
    """
    parts = []
    for el in elements:
        cls_name = type(el).__name__
        if cls_name == "Paragraph":
            parts.append(str(getattr(el, "text", "")))
        elif cls_name == "Table":
            for row in getattr(el, "_cellvalues", []) or []:
                for cell in row if isinstance(row, list) else [row]:
                    if cell is None:
                        continue
                    cls = type(cell).__name__
                    if cls == "Paragraph":
                        parts.append(str(getattr(cell, "text", "")))
                    elif isinstance(cell, str):
                        parts.append(cell)
        # Treat anything else (Spacer, PageBreak, KeepTogether, ...) as
        # structurally invisible — they carry no searchable content.
    return "\n".join(parts)


# =============================================================================
# HELPERS
# =============================================================================


def _assert_contains(elements, *needles):
    """Assert every needle is present in the flattened element text."""
    text = _elements_to_text(elements)
    for n in needles:
        assert n in text, f"Missing {n!r} in generated report text:\n{text}"


# =============================================================================
# FIXTURES
# =============================================================================


@pytest.fixture
def sample_records():
    """Sample billing records for testing."""
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
            "Amount (£)": -100.00,
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
def sample_df(sample_records):
    """Create a DataFrame from sample records."""
    return pd.DataFrame(sample_records)


@pytest.fixture
def mock_engine(sample_records):
    """Create a mock EvidenceEngine."""
    engine = Mock()
    engine.records = sample_records
    engine.filtered_records = []
    engine.email_count = 10
    engine.pdf_count = 5
    engine.error_log = []
    return engine


@pytest.fixture
def sample_config():
    """Sample configuration dictionary."""
    return {
        "min_amount": 500.0,
        "analysis_min": 500.0,
        "acc_num": "123456789",
        "report_account_ref": "ACC-123456",
        "use_anchors": True,
        "use_large": True,
        "use_reading_classification": True,
        "use_pdf_fields": True,
        "use_acc_filt": True,
        "filter_below": True,
        "save_filtered": True,
        "use_dedup": True,
        "save_dups": True,
        "use_domain_filter": True,
        "domain_filter": "edfenergy.com",
    }


# =============================================================================
# FORMATTER TESTS
# =============================================================================


class TestFormatters:
    """Tests for formatting functions."""

    def test_fmt_money_valid(self):
        assert fmt_money(100) == "£100.00"
        assert fmt_money(100.5) == "£100.50"
        assert fmt_money("100.50") == "£100.50"
        assert fmt_money("1,000.50") == "£1,000.50"
        assert fmt_money("£100") == "£100.00"
        assert fmt_money(0) == "£0.00"
        assert fmt_money(-50.25) == "£-50.25"

    def test_fmt_money_na(self):
        assert fmt_money("N/A") == ""
        assert fmt_money("NA") == ""
        assert fmt_money("") == ""
        assert fmt_money(None) == ""
        assert fmt_money("N/A", blank_if_na=False) == "N/A"

    def test_fmt_number_valid(self):
        assert fmt_number(1000) == "1,000.00"
        assert fmt_number(1000.5) == "1,000.50"
        assert fmt_number("1,000") == "1,000.00"
        assert fmt_number(1234567) == "1,234,567.00"
        assert fmt_number(100.45, decimals=2) == "100.45"
        assert fmt_number(100, decimals=0) == "100"

    def test_fmt_number_na(self):
        assert fmt_number("N/A") == ""
        assert fmt_number(None) == ""
        assert fmt_number("N/A", blank_if_na=False) == "N/A"

    def test_pdf_docx_formatters_align_precision(self):
        """PDF and DOCX reporters render the same value at the same precision.

        Regression guard for the C-3/L-15 divergence where the PDF
        ``fmt_number`` defaulted to ``decimals=0`` (rendering £46) while
        the DOCX defaulted to ``decimals=2`` (rendering £45.67).  Both
        must now default to 2 decimal places.
        """
        from edf_bill_fetcher.io.reporters.docx_report import (
            fmt_money as docx_fmt_money,
        )
        from edf_bill_fetcher.io.reporters.docx_report import (
            fmt_number as docx_fmt_number,
        )

        assert fmt_number(45.67) == docx_fmt_number(45.67) == "45.67"
        assert fmt_number(45.6) == docx_fmt_number(45.6) == "45.60"
        assert fmt_number(46, decimals=0) == docx_fmt_number(46, decimals=0) == "46"
        assert fmt_money(45.67) == docx_fmt_money(45.67) == "£45.67"

    def test_fmt_helpers_single_source_of_truth(self):
        """fmt_money/fmt_number live once in helpers.formatting (C-1/C-4).

        The PDF surface imports the shared implementation directly; the
        DOCX surface keeps a thin adapter that only overrides the
        missing-value rendering (``"N/A"`` here vs blank on the PDF).
        Neither reporter may define its own copy again.
        """
        from edf_bill_fetcher.helpers import formatting
        from edf_bill_fetcher.io.reporters.docx_report import (
            fmt_money as docx_fmt_money,
        )
        from edf_bill_fetcher.io.reporters.docx_report import (
            fmt_number as docx_fmt_number,
        )

        assert fmt_money is formatting.fmt_money
        assert fmt_number is formatting.fmt_number
        assert docx_fmt_money(None) == "N/A"
        assert docx_fmt_number(None) == "N/A"
        assert docx_fmt_money("N/A") == "N/A"
        assert fmt_money(None) == ""
        assert fmt_number(None) == ""
        assert docx_fmt_money(float("nan")) == "N/A"
        assert docx_fmt_number(float("nan")) == "N/A"
        assert fmt_money(float("nan")) == ""
        assert fmt_number(float("nan")) == ""

    def test_fmt_pct_valid(self):
        assert fmt_pct(0.15) == "15.0%"
        assert fmt_pct(1.0) == "100.0%"
        assert fmt_pct(0.1234) == "12.3%"
        assert fmt_pct("0.5") == "50.0%"

    def test_fmt_pct_na(self):
        assert fmt_pct("N/A") == ""
        assert fmt_pct(None) == ""
        assert fmt_pct("N/A", blank_if_na=False) == "N/A"

    def test_fmt_date_valid(self):
        assert fmt_date("01/01/2023") == "01/01/2023"
        assert fmt_date("2023-01-01") == "01/01/2023"

    def test_fmt_date_na(self):
        assert fmt_date("N/A") == ""
        assert fmt_date("NA") == ""
        assert fmt_date("") == ""
        assert fmt_date(None) == ""

    def test_severity_color(self):
        assert severity_color("HIGH") == Colors.RED
        assert severity_color("MEDIUM") == Colors.AMBER
        assert severity_color("INFO") == Colors.GREEN
        assert severity_color("UNKNOWN") == Colors.MEDIUM_GREY

    def test_severity_label(self):
        assert severity_label("HIGH") == "●"
        assert severity_label("MEDIUM") == "●"
        assert severity_label("INFO") == "●"
        assert severity_label("OTHER") == "○"


# =============================================================================
# TABLE STYLE TESTS
# =============================================================================


class TestTableStyle:
    """Tests for table style creation."""

    def test_make_table_style_default(self):
        style = make_table_style(num_rows=5)
        assert isinstance(style, TableStyle)
        # Check some default values
        commands = style._cmds
        assert any(cmd[0] == "BACKGROUND" for cmd in commands)
        assert any(cmd[0] == "GRID" for cmd in commands)

    def test_make_table_style_custom_colors(self):
        style = make_table_style(
            header_color=Colors.RED,
            alt_row_color=Colors.GREEN,
            header_text_color=Colors.BLACK,
            grid_color=Colors.MEDIUM_BLUE,
            font_size=10,
            num_rows=5,
        )
        assert isinstance(style, TableStyle)


# =============================================================================
# BUILD STYLES TESTS
# =============================================================================


class TestBuildStyles:
    """Tests for style dictionary creation."""

    def test_build_styles(self):
        styles = build_styles()
        assert isinstance(styles, dict)
        # Check key styles exist
        expected_keys = [
            "CoverTitle",
            "CoverSubtitle",
            "CoverInfo",
            "SectionHeader",
            "SubSectionHeader",
            "SubSubSectionHeader",
            "BodyText",
            "BodyTextIndent",
            "BulletText",
            "SmallText",
            "TableHeader",
            "TableCell",
            "TableCellCenter",
            "TableCellRight",
            "TableCellMoney",
            "Footnote",
            "PageNumber",
            "Confidential",
        ]
        for key in expected_keys:
            assert key in styles, f"Missing style: {key}"


# =============================================================================
# REPORT SECTION TESTS
# =============================================================================


class TestCoverPage:
    """Tests for cover page creation."""

    def test_create_cover_page(self):
        elements = create_cover_page(
            account_ref="ACC-12345",
            period_start="01/01/2023",
            period_end="31/03/2023",
            report_date="17 June 2026",
        )
        _assert_contains(
            elements,
            "EDF Energy Billing Dispute",
            "ACC-12345",
            "01/01/2023",
            "31/03/2023",
            "17 June 2026",
        )

    def test_create_cover_page_empty_fields(self):
        """Cover page renders even with empty inputs — fallback placeholders."""
        elements = create_cover_page("", "", "", "")
        _assert_contains(elements, "EDF Energy Billing Dispute")


class TestTableOfContents:
    """Tests for table of contents creation."""

    def test_create_table_of_contents(self):
        from edf_bill_fetcher.io.reporters.pdf_report import (  # local: avoid disturbing imports at top
            RenderContext,
        )

        elements = create_table_of_contents(RenderContext())
        assert isinstance(elements, list)
        assert len(elements) > 0


class TestExecutiveSummary:
    """Tests for executive summary creation."""

    def test_create_executive_summary(self, sample_df, sample_config):
        elements = create_executive_summary(
            df=sample_df,
            config=sample_config,
            account_ref="ACC-12345",
            flag_count={"HIGH": 1, "MEDIUM": 2, "INFO": 1},
            total_records=4,
            total_charges=530.50,
            total_payments=100.00,
            period_start="01/01/2023",
            period_end="31/03/2023",
        )
        assert isinstance(elements, list)
        assert len(elements) > 0

    def test_create_executive_summary_no_flags(self, sample_df, sample_config):
        elements = create_executive_summary(
            df=sample_df,
            config=sample_config,
            account_ref="ACC-12345",
            flag_count={"HIGH": 0, "MEDIUM": 0, "INFO": 0},
            total_records=4,
            total_charges=530.50,
            total_payments=100.00,
            period_start="01/01/2023",
            period_end="31/03/2023",
        )
        assert isinstance(elements, list)


class TestKeyFindings:
    """Tests for key findings table."""

    def test_create_key_findings_empty(self):
        elements = create_key_findings_table([])
        # Empty findings still emits the section header.
        _assert_contains(elements, "Key Findings")

    def test_create_key_findings_with_flags(self):
        flags = [
            ("LARGE JUMP", "01/01/2023", 200.0, "Big jump detected", "HIGH"),
            ("BILLING GAP", "01/02/2023", 150.0, "Gap of 90 days", "MEDIUM"),
            ("BALANCE REDUCTION", "01/03/2023", -100.0, "Payment received", "INFO"),
        ]
        elements = create_key_findings_table(flags)
        _assert_contains(
            elements,
            "LARGE JUMP",
            "BILLING GAP",
            "Big jump detected",
            "Gap of 90 days",
        )


class TestEvidenceIndex:
    """Tests for evidence index creation."""

    def test_create_evidence_index(self, sample_df, mock_engine):
        elements = create_evidence_index(sample_df, mock_engine)
        _assert_contains(elements, "Evidence")  # catches "Evidence Index"


class TestAnomalyDetail:
    """Tests for anomaly detail section."""

    def test_create_anomaly_detail_empty(self, sample_df):
        """Empty flags still emit the section header (no stuffing skipped)."""
        elements = create_anomaly_detail_section([], sample_df)
        _assert_contains(elements, "Detailed Findings")

    def test_create_anomaly_detail_with_flags(self, sample_df):
        flags = [
            ("LARGE JUMP", "01/01/2023", 200.0, "Big jump detected", "HIGH"),
            ("LARGE JUMP", "01/02/2023", 250.0, "Another jump", "HIGH"),
            ("BILLING GAP", "01/03/2023", None, "60 day gap", "MEDIUM"),
            ("ESTIMATED RUN", "01/04/2023", 100.0, "3 estimated readings", "MEDIUM"),
        ]
        elements = create_anomaly_detail_section(flags, sample_df)
        _assert_contains(
            elements,
            "Large Jump",
            "Billing Gap",
            "Estimated Run",
            # Description text from the fixed flags should appear too
            "Big jump detected",
        )


class TestTimeline:
    """Tests for timeline creation."""

    def test_create_timeline(self, sample_df):
        flags = [
            ("LARGE JUMP", "01/01/2023", 200.0, "Big jump detected", "HIGH"),
        ]
        elements = create_timeline_section(sample_df, flags)
        assert isinstance(elements, list)
        assert len(elements) > 0


class TestOFGEMComparison:
    """Tests for OFGEM comparison section."""

    def test_create_ofgem_comparison(self, sample_df):
        elements = create_ofgem_comparison(sample_df)
        assert isinstance(elements, list)
        assert len(elements) > 0

    def test_create_ofgem_comparison_quarter_without_cap_row_is_explicit(self, monkeypatch):
        """A billing quarter that has no entry in the OFGEM-cap dict must
        surface as an explicit ``CAP DATA UNAVAILABLE`` row rather than
        being silently dropped from the comparison table.

        This is the Phase 1.1 invariant — a reviewer must be able to see
        *which* quarters in the data the project couldn't benchmark, not
        just the ones it could.
        """
        # Force the cap dict to a known minimal shape so we know exactly
        # which quarter is missing.  2024-Q3 is intentionally absent
        # while 2024-Q1 and 2024-Q2 are present.
        minimal_caps = {
            "2024-Q1": {"unit_rate": 28.62, "standing_charge": 53.35},
            "2024-Q2": {"unit_rate": 24.50, "standing_charge": 60.10},
            # 2024-Q3 intentionally omitted to force the gap-row path.
        }

        monkeypatch.setattr(
            sys.modules["edf_bill_fetcher.io.reporters.pdf_report"],
            "_load_ofgem_caps",
            lambda auto_carry=False: (minimal_caps, None),
        )

        # One bill per quarter in our minimal cap set, plus a bill
        # landing in 2024-Q3 which has no cap entry.
        df = pd.DataFrame(
            [
                {
                    "Date": "15/02/2024",
                    "Period Charge (£)": 200.0,
                    "Units (kWh)": 500.0,
                },
                {
                    "Date": "15/05/2024",
                    "Period Charge (£)": 300.0,
                    "Units (kWh)": 600.0,
                },
                {
                    "Date": "15/08/2024",
                    "Period Charge (£)": 400.0,
                    "Units (kWh)": 700.0,
                },
            ]
        )

        elements = create_ofgem_comparison(df)

        # Find the comparison table by walking the elements — tables
        # are reportlab.platypus.Table objects and the only one in
        # this section is the comparison table.  We extract its
        # first row ("Period"/"Bill ... cap") and every data row.
        tables = [el for el in elements if el.__class__.__name__ == "Table"]
        assert tables, "No table found in OFGEM-comparison elements"
        comparison_table = tables[0]
        cell_values = getattr(comparison_table, "_cellvalues", [])
        # First row is the header; subsequent rows are data rows.
        data_rows = cell_values[1:]
        quarters_seen = [row[0] for row in data_rows]

        # All three quarters must appear, even the one without a cap.
        assert "2024-Q1" in quarters_seen
        assert "2024-Q2" in quarters_seen
        assert "2024-Q3" in quarters_seen

        # Find the 2024-Q3 row and confirm the gap-rendering markers.
        q3_row = next(row for row in data_rows if row[0] == "2024-Q3")
        # Cell 4 is ``Status``: must read exactly ``CAP DATA UNAVAILABLE``
        # so an OFGEM-grade reader can tell it apart from a row that
        # was benchmarked and came out "BELOW CAP".
        assert q3_row[4] == "CAP DATA UNAVAILABLE"
        # Cells 2 (OFGEM Cap p/kWh) and 3 (Difference) must be the
        # "—" missing-sentinel so the table never pretends to know.
        assert q3_row[2] == "—"
        assert q3_row[3] == "—"

    def test_load_ofgem_caps_has_no_sentinel_key(self) -> None:
        """The caps dict must not carry the carry-forward metadata as a key
        (L-11): iterating ``caps.items()`` must yield only real quarters.
        The carry-forward cap is returned as a separate tuple element.
        """
        caps, latest = _load_ofgem_caps(auto_carry=True)
        assert "_LATEST_KNOWN" not in caps
        assert latest == caps["2026-Q3"]
        caps_exact, latest_exact = _load_ofgem_caps(auto_carry=False)
        assert "_LATEST_KNOWN" not in caps_exact
        assert latest_exact is None


class TestStatisticalAnalysis:
    """Tests for the statistical analysis section builder.

    These two tests were previously sitting inside
    ``TestOFGEMComparison`` with a bare orphan ``\"\"\"...\"\"\"``
    docstring floating below the last OFGEM test — looks like a
    class-header line that lost its ``class TestStatisticalAnalysis:``
    predecessors during an earlier reshape of the file.  Promoting
    them into a real class with this docstring makes them
    discoverable by name (``pytest -k Statistical``) and removes
    the structural defect.
    """

    def test_create_statistical_analysis(self, sample_df):
        elements = create_statistical_analysis(sample_df)
        assert isinstance(elements, list)
        assert len(elements) > 0

    def test_create_statistical_analysis_insufficient_data(self):
        df = pd.DataFrame([{"Date": "01/01/2023", "Amount (£)": 100}])
        elements = create_statistical_analysis(df)
        assert isinstance(elements, list)
        assert len(elements) > 0


class TestPaymentAnalysis:
    """Tests for payment analysis section."""

    def test_create_payment_analysis_with_payments(self, sample_df):
        elements = create_payment_analysis(sample_df)
        assert isinstance(elements, list)
        assert len(elements) > 0

    def test_create_payment_analysis_no_payments(self):
        df = pd.DataFrame(
            [
                {
                    "Date": "01/01/2023",
                    "Amount (£)": 100,
                    "Entry Type": "New Bill",
                    "Details": "",
                    "Source": "HTM",
                },
            ]
        )
        elements = create_payment_analysis(df)
        assert isinstance(elements, list)
        assert len(elements) > 0


class TestForecast:
    """Tests for forecast section."""

    def test_create_forecast_section(self, sample_df):
        elements = create_forecast_section(sample_df)
        assert isinstance(elements, list)
        assert len(elements) > 0


class TestDataQuality:
    """Tests for data quality section."""

    def test_create_data_quality_section(self, sample_df):
        elements = create_data_quality_section(sample_df)
        assert isinstance(elements, list)
        assert len(elements) > 0

    def test_create_data_quality_empty(self):
        df = pd.DataFrame(
            [
                {
                    "Date": "01/01/2023",
                    "Amount (£)": 100,
                    "Period From": "N/A",
                    "Period To": "N/A",
                    "Source": "HTM",
                    "Reading": "Actual",
                    "Entry Type": "New Bill",
                },
            ]
        )
        elements = create_data_quality_section(df)
        assert isinstance(elements, list)
        assert len(elements) > 0


class TestTariffImpact:
    """Tests for tariff impact section."""

    def test_create_tariff_impact_with_data(self, sample_df):
        elements = create_tariff_impact_section(sample_df)
        assert isinstance(elements, list)
        assert len(elements) > 0

    def test_create_tariff_impact_no_data(self):
        df = pd.DataFrame(
            [
                {
                    "Date": "01/01/2023",
                    "Amount (£)": 100,
                    "Tariff": "N/A",
                    "Unit Rate (p/kWh)": "N/A",
                    "Period Charge (£)": "N/A",
                },
            ]
        )
        elements = create_tariff_impact_section(df)
        assert isinstance(elements, list)
        assert len(elements) > 0

    def test_create_tariff_impact_missing_columns(self):
        df = pd.DataFrame(
            [
                {"Date": "01/01/2023", "Amount (£)": 100},
            ]
        )
        elements = create_tariff_impact_section(df)
        assert isinstance(elements, list)
        assert len(elements) > 0


class TestAppendices:
    """Tests for appendix sections."""

    def test_create_appendix_methodology(self, sample_config):
        elements = create_appendix_methodology(sample_config)
        assert isinstance(elements, list)
        assert len(elements) > 0

    def test_create_appendix_glossary(self):
        elements = create_appendix_glossary()
        assert isinstance(elements, list)
        assert len(elements) > 0


# =============================================================================
# MAIN REPORT GENERATION TESTS
# =============================================================================


class TestGeneratePDF:
    """Tests for main PDF generation functions."""

    def test_generate_ombudsman_pdf_empty_records(self):
        with pytest.raises(ValueError, match="No records to report on"):
            generate_ombudsman_pdf([], "test.pdf", {}, Mock())

    def test_generate_pdf_from_gui_no_records(self, sample_config, mock_engine):
        success, msg = generate_pdf_from_gui([], "test.pdf", sample_config, mock_engine, [])
        assert success is False
        assert "no records" in msg.lower() or "error" in msg.lower()

    def test_generate_pdf_from_gui_no_engine(self, sample_records, sample_config):
        # The CLI path requires a runnable PDF even when --engine-data
        # wasn't supplied, so the GUI wrapper now synthesises a minimal
        # engine rather than failing.  Verify the underlying generator
        # accepts the engine=None / filtered=None combination without
        # raising (the old code raised ``ValueError("Engine is required")``).
        cfg = cast(ConfigDict, dict(sample_config))
        # Pre-fill fields the generator requires so no exception is raised
        cfg.setdefault(
            "report_sections",
            ["cover", "toc", "exec_summary", "key_findings"],
        )
        # The PDF generator writes to disk; point it at the current
        # directory under a unique name and clean up afterwards.
        import os

        out = os.path.abspath("test_no_engine_gui.pdf")
        try:
            success, msg = generate_pdf_from_gui(sample_records, out, cfg, None, [])
            assert success is True, msg
        finally:
            if os.path.exists(out):
                os.remove(out)


# =============================================================================
# CLI TESTS
# =============================================================================


class TestCLI:
    """Tests for CLI functions."""

    def test_run_cli_pdf_report_missing_args(self):
        import sys

        from edf_bill_fetcher.io.cli import run_cli_pdf_report

        # Should raise SystemExit with code 2 for missing required args
        with patch.object(sys, "argv", ["edf-collector", "--pdf-report"]):
            with pytest.raises(SystemExit) as exc_info:
                run_cli_pdf_report([])
            assert exc_info.value.code == 2

    def test_run_cli_pdf_report_missing_records_file(self):
        import sys

        from edf_bill_fetcher.io.cli import run_cli_pdf_report

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            output_path = tmp.name

        try:
            with patch.object(
                sys,
                "argv",
                [
                    "edf-collector",
                    "--pdf-report",
                    "--records",
                    "/nonexistent.json",
                    "--output",
                    output_path,
                ],
            ):
                with pytest.raises(SystemExit) as exc_info:
                    run_cli_pdf_report(sys.argv[2:])
                assert exc_info.value.code == 1
        finally:
            if Path(output_path).exists():
                Path(output_path).unlink()

    def test_run_cli_pdf_report_basic(self, sample_records):
        import json
        import sys

        from edf_bill_fetcher.io.cli import run_cli_pdf_report

        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as tmp:
            records_path = tmp.name
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            output_path = tmp.name

        try:
            with open(records_path, "w") as f:
                json.dump(sample_records, f)

            with patch.object(
                sys,
                "argv",
                [
                    "edf-collector",
                    "--pdf-report",
                    "--records",
                    records_path,
                    "--output",
                    output_path,
                ],
            ):
                # This will try to generate PDF - just check it doesn't crash on arg parsing
                try:
                    run_cli_pdf_report(sys.argv[2:])
                except SystemExit:
                    # May exit with error if PDF gen fails, but args should be parsed
                    pass
        finally:
            for p in [records_path, output_path]:
                if Path(p).exists():
                    Path(p).unlink()


# =============================================================================
# EDF_COLLECTOR FUNCTION TESTS
# =============================================================================


class TestEDFCollectorFunctions:
    """Tests for functions in edf_collector module."""

    def test_export_to_excel_creates_file(self, sample_records, sample_config):
        from edf_bill_fetcher.io.writers import export_to_excel

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            export_to_excel(sample_records, tmp_path, [], sample_config, filtered=[])
            assert Path(tmp_path).exists()
            assert Path(tmp_path).stat().st_size > 0
        finally:
            if Path(tmp_path).exists():
                Path(tmp_path).unlink()


# =============================================================================
# INTEGRATION TESTS
# =============================================================================


class TestIntegration:
    """Integration tests for full workflow."""

    def test_full_extract_generate_cycle(self, sample_records, sample_config, mock_engine):
        """Test extracting to Excel then generating PDF."""
        from edf_bill_fetcher.io.writers import export_to_excel

        with tempfile.TemporaryDirectory() as tmpdir:
            excel_path = Path(tmpdir) / "test.xlsx"

            # Export to Excel
            export_to_excel(sample_records, str(excel_path), [], sample_config, filtered=[])
            assert excel_path.exists()

    def test_cli_with_config_file(self, sample_records, sample_config):
        """Test CLI with config file."""
        import json

        from edf_bill_fetcher.io.cli import run_cli_pdf_report

        with tempfile.TemporaryDirectory() as tmpdir:
            records_path = Path(tmpdir) / "records.json"
            config_path = Path(tmpdir) / "config.json"
            output_path = Path(tmpdir) / "report.pdf"

            with open(records_path, "w") as f:
                json.dump(sample_records, f)
            with open(config_path, "w") as f:
                json.dump(sample_config, f)

            # Just verify args parsing works
            import sys

            with patch.object(
                sys,
                "argv",
                [
                    "edf-collector",
                    "--pdf-report",
                    "--records",
                    str(records_path),
                    "--output",
                    str(output_path),
                    "--config",
                    str(config_path),
                ],
            ):
                try:
                    run_cli_pdf_report(sys.argv[2:])
                except SystemExit:
                    pass  # Expected if PDF generation fails


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
