"""End-to-end integration smoke test for the PDF -> engine -> PDF pipeline.

This test exercises the real bundled ``test.pdf`` through:

1. ``EvidenceEngine`` — parses the PDF (via ``pdfplumber``) and extracts
   records internally.
2. ``generate_ombudsman_pdf`` — assembles the extracted records into a new
   reportlab PDF.
3. ``export_to_excel`` — serialises the same dataset to XLSX.

It guarantees the full pipeline does not raise an exception and
produces non-empty artifacts.

The bundled ``test.pdf`` is a 17-page EDF Ombudsman Evidence Report
template, not raw bill data — so the pipeline produces empty-but-valid
output. The smoke still validates that the wiring survives a real
PDF round-trip.
"""

from pathlib import Path

import pytest

from edf_collector import EvidenceEngine, export_to_excel
from edf_report import generate_ombudsman_pdf


class TestIntegrationSmoke:
    """Full-pipeline smoke test using the bundled ``test.pdf`` sample."""

    @pytest.fixture
    def config(self):
        """Permissive config that accepts any record."""
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

    def test_pdf_to_records_no_exception(self, engine, config):
        """Process the bundled ``test.pdf`` without raising."""
        pdf_path = Path(__file__).parent.parent / "test.pdf"
        assert pdf_path.exists(), "Bundled test.pdf must exist"
        engine.process_pdf_file(
            str(pdf_path),
            source_label="Local PDF",
            detail_label="test.pdf",
            fallback_date="2024-01-01",
        )
        assert isinstance(engine.records, list)

    def test_full_pipeline_creates_pdf_and_xlsx(self, engine, config, tmp_path):
        """Run PDF extraction → report generation → Excel export."""
        pdf_path = Path(__file__).parent.parent / "test.pdf"
        assert pdf_path.exists(), "Bundled test.pdf must exist"
        engine.process_pdf_file(
            str(pdf_path),
            source_label="Local PDF",
            detail_label="test.pdf",
            fallback_date="2024-01-01",
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
