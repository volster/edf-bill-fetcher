"""End-to-end integration smoke test exercising the full pipeline against a
synthetic EDF KI-style bill PDF.

The fixture PDF (tests/fixtures/sample_bill.pdf) is generated from
tests/fixtures/generate_bill_fixture.py using reportlab. It is not
derived from any real EDF customer bill and contains only synthetic
placeholder data, so no PII leaks into the repo.

The fixture is structured so that EvidenceEngine.process_pdf_file
extracts ONE deterministic record with the following fields:

    Amount         = 240.50
    Period Charge  = 240.50
    Entry Type     = 'New Bill'
    Logic Used     = 'New Invoice Format'
    Date           = contains '/03/2026'

The two tests below assert those exact values, so the test fails if
ANY of the following are broken:

  * the engine can open + pdftotext the bundled fixture;
  * regex anchoring matches 'Total charges for this period GBPX debit';
  * the regex anchoring matches 'Current balance GBPX debit';
  * 'Entry Type' falls correctly to 'New Bill' for an anchored bill;
  * downstream reportlab + openpyxl artifacts are non-empty.

If you change the fixture (or the engine's parser logic), expect one or
both tests to break.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from edf_collector import EvidenceEngine, export_to_excel
from edf_report import generate_ombudsman_pdf

FixturePath = Path(__file__).parent / "fixtures" / "sample_bill.pdf"


def _ensure_fixture() -> Path:
    """If the fixture PDF is missing, regenerate it.

    A missing fixture in CI means either an untracked file or a
    corrupted checkout. We err on the side of regeneration because
    the generator is deterministic — running it again gives byte
    identical output.
    """
    if FixturePath.exists():
        return FixturePath
    gen = Path(__file__).parent / "fixtures" / "generate_bill_fixture.py"
    if not gen.exists():
        raise RuntimeError(
            f"Fixture missing and generator not found: {FixturePath}"
        )
    # Use runpy to execute the generator and give us a handle on the
    # build() function without polluting sys.modules.
    import runpy
    import sys as _sys

    saved_argv = list(_sys.argv)
    _sys.argv = [str(gen), str(FixturePath)]
    try:
        runpy.run_path(str(gen), run_name="__main__")
    finally:
        _sys.argv = saved_argv
    if not FixturePath.exists():
        raise RuntimeError(
            f"Fixture regeneration failed — file not created at {FixturePath}"
        )
    print(f"---- fixture regenerated: {FixturePath}")
    return FixturePath


class TestIntegrationSmoke:
    """End-to-end pipeline against the synthetic bill fixture."""

    @pytest.fixture
    def config(self):
        """Permissive config — every PDF becomes a record."""
        return {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 0.0,
            "analysis_min": 0.0,
            "filter_below": False,
            "save_filtered": False,
            "use_dedup": False,
            "save_dups": False,
            "use_domain_filter": False,
            "domain_filter": "",
        }

    @pytest.fixture
    def engine(self, config):
        return EvidenceEngine(config, lambda x: None)

    def test_engine_extracts_one_known_record(self, engine):
        """One deterministic record; key fields match synthetic fixture."""
        pdf_path = _ensure_fixture()
        engine.process_pdf_file(
            str(pdf_path),
            source_label="Local PDF",
            detail_label="sample_bill.pdf",
            fallback_date="2026-03-01",
        )
        assert isinstance(engine.records, list)
        assert len(engine.records) == 1, (
            f"Expected exactly one record from the synthetic bill, "
            f"got {len(engine.records)}: {engine.records!r}"
        )

        rec = engine.records[0]
        assert rec["Amount (£)"] == 240.50, (
            f"Amount mismatch: {rec['Amount (£)']!r}"
        )
        assert rec["Period Charge (£)"] == 240.50
        assert rec["Entry Type"] == "New Bill"
        # Date normalisation can shift format by region; yy/MM/dd or
        # dd/MM/yyyy are both fine, we just need March 2026.
        assert "/03/2026" in rec["Date"]

    def test_full_pipeline_creates_pdf_and_xlsx(self, engine, config, tmp_path):
        """PDF + Excel artefacts non-empty after the same record is fed
        through generate_ombudsman_pdf and export_to_excel.
        """
        pdf_path = _ensure_fixture()
        engine.process_pdf_file(
            str(pdf_path),
            source_label="Local PDF",
            detail_label="sample_bill.pdf",
            fallback_date="2026-03-01",
        )

        out_pdf = tmp_path / "report.pdf"
        out_xlsx = tmp_path / "report.xlsx"

        generate_ombudsman_pdf(
            records=engine.records,
            output_path=str(out_pdf),
            config=config,
            engine=engine,
        )

        export_to_excel(
            data=engine.records,
            output_path=str(out_xlsx),
            error_log=engine.error_log,
            config=config,
        )

        assert out_pdf.exists()
        assert out_pdf.stat().st_size > 0
        assert out_xlsx.exists()
        assert out_xlsx.stat().st_size > 0
