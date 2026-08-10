"""Tests for EvidenceEngine PDF/PST/HTM processing methods."""

from pathlib import Path
from unittest.mock import MagicMock, Mock, mock_open, patch

import pytest

from edf_bill_fetcher.collectors.engine import EvidenceEngine
from edf_bill_fetcher.models.config import ConfigDict


class TestEvidenceEnginePDF:
    """Tests for PDF processing methods."""

    def _make_engine(self):
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 100.0,
            "analysis_min": 100.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        return EvidenceEngine(config, lambda x: None)

    def _make_pdf_mock(self, text):
        """Create a mock pdfplumber context manager that returns the given text."""
        mock_page = MagicMock()
        mock_page.extract_text.return_value = text

        mock_pdf = MagicMock()
        mock_pdf.pages = [mock_page]
        mock_pdf.__enter__.return_value = mock_pdf
        mock_pdf.__exit__.return_value = None

        return mock_pdf

    @patch("edf_bill_fetcher.collectors.engine.pdfplumber.open")
    def test_process_new_invoice(self, mock_pdf_open):
        invoice_text = """
        Invoice number: KI-12345678
        Account number: A-31105244
        Date issued: 15 January 2024
        Your charges: 01 Jan 2024 - 31 Jan 2024
        Current balance £1,234.56 debit
        Total charges for this period £89.99 debit
        Electricity used 350 kWh
        Standing charge 25.50p/day
        Tariff name Standard Variable
        """
        mock_pdf_open.return_value = self._make_pdf_mock(invoice_text)

        engine = self._make_engine()

        # Mock the file read
        with patch("builtins.open", mock_open(read_data=b"fake pdf content")):
            engine.process_pdf_file("test.pdf", "Test", "test.pdf", "01/01/2024")

        assert len(engine.records) == 1
        assert engine.records[0]["Invoice #"] == "KI-12345678"
        assert engine.records[0]["Amount (£)"] == 1234.56

    @patch("edf_bill_fetcher.collectors.engine.pdfplumber.open")
    def test_process_new_credit(self, mock_pdf_open):
        credit_text = """
        Credit note number: KCR-87654321
        Account number: A-31105244
        Date issued: 15 January 2024
        Total credits for this bill £150.00
        """
        mock_pdf_open.return_value = self._make_pdf_mock(credit_text)

        engine = self._make_engine()

        with patch("builtins.open", mock_open(read_data=b"fake pdf content")):
            engine.process_pdf_file("test.pdf", "Test", "test.pdf", "01/01/2024")

        assert len(engine.records) == 1
        assert engine.records[0]["Entry Type"] == "Credit"
        assert engine.records[0]["Amount (£)"] == 150.00

    @patch("edf_bill_fetcher.collectors.engine.pdfplumber.open")
    def test_process_old_format_pdf(self, mock_pdf_open):
        old_text = "Some text without KI or KCR markers\nYour new account balance £500.00"
        mock_pdf_open.return_value = self._make_pdf_mock(old_text)

        engine = self._make_engine()

        with patch("builtins.open", mock_open(read_data=b"fake pdf content")):
            engine.process_pdf_file("test.pdf", "Test", "test.pdf", "01/01/2024")

        assert len(engine.records) == 1
        # The amount is found via "Your new account balance" pattern (Smart Context)
        assert engine.records[0]["Amount (£)"] == 500.00

    def test_pdf_extract_error_handling(self):
        engine = self._make_engine()

        with patch("builtins.open", mock_open(read_data=b"fake")):
            with patch(
                "edf_bill_fetcher.collectors.engine.pdfplumber.open",
                side_effect=Exception("PDF read error"),
            ):
                engine.process_pdf_file("bad.pdf", "Test", "bad.pdf", "01/01/2024")

        assert len(engine.error_log) == 1
        assert "PDF read error" in engine.error_log[0]


class TestEvidenceEnginePST:
    """Tests for PST crawling."""

    def _make_engine(self):
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 100.0,
            "analysis_min": 100.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        return EvidenceEngine(config, lambda x: None)

    def test_crawl_pst_skips_non_message(self):
        engine = self._make_engine()

        mock_folder = Mock()
        mock_folder.get_number_of_sub_messages.return_value = 0
        mock_folder.get_number_of_sub_folders.return_value = 0
        mock_folder.sub_folders = []

        engine.crawl_pst(mock_folder)
        assert len(engine.records) == 0


class TestEvidenceEngineHTM:
    """Tests for HTM file processing."""

    def _make_engine(self):
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 100.0,
            "analysis_min": 100.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        return EvidenceEngine(config, lambda x: None)

    def test_process_htm_file(self):
        engine = self._make_engine()

        htm_content = """
        15 Jan 2024 We charged your account £89.99 For 350 kWh of electricity used between 01 Jan 2024 and 31 Jan 2024 Balance £1,234.56 in debit
        01 Feb 2024 You paid us £200.00 Balance £1,034.56 in debit
        """

        with patch("builtins.open", mock_open(read_data=htm_content)):
            engine.process_htm_file("test.htm")

        assert len(engine.records) == 2
        assert engine.records[0]["Entry Type"] == "Ongoing Balance"
        assert engine.records[1]["Entry Type"] == "Payment"

    def test_process_htm_file_not_found(self):
        engine = self._make_engine()
        engine.process_htm_file("nonexistent.htm")
        assert len(engine.error_log) == 1


class TestEvidenceEngineLocalPDFs:
    """Tests for local PDF folder crawling."""

    def _make_engine(self):
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 100.0,
            "analysis_min": 100.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        return EvidenceEngine(config, lambda x: None)

    @patch("edf_bill_fetcher.collectors.engine.pdfplumber.open")
    def test_crawl_local_pdfs(self, mock_pdf_open):
        mock_pdf = MagicMock()
        mock_page = MagicMock()
        mock_page.extract_text.return_value = "Current balance £500.00 debit"
        mock_pdf.pages = [mock_page]
        mock_pdf.__enter__.return_value = mock_pdf
        mock_pdf.__exit__.return_value = None
        mock_pdf_open.return_value = mock_pdf

        # Phase 2.2 — crawl_local_pdfs now uses os.walk (recursive
        # walk) instead of os.listdir.  The legacy test stubbed
        # listdir; we now stub walk to return the same two bills.
        with (
            patch("os.path.exists", return_value=True),
            patch(
                "os.walk",
                return_value=[
                    ("/fake/path", [], ["bill1.pdf", "bill2.pdf"]),
                ],
            ),
            patch("os.path.getmtime", return_value=1704067200),
            patch("builtins.open", mock_open(read_data=b"fake pdf content")),
        ):
            engine = self._make_engine()
            engine.crawl_local_pdfs("/fake/path")

        assert engine.pdf_count == 2

    @patch("edf_bill_fetcher.collectors.engine.pdfplumber.open")
    def test_crawl_local_pdfs_recurses_into_subfolders(self, mock_pdf_open: MagicMock) -> None:
        """Phase 2.2 — recursive walk yields bills in nested
        folders.  Real EDF customers commonly organise their
        local PDF tree by year (``pdfs/2023/2023-01.pdf``) and
        the legacy top-level-only scan silently dropped every
        bill below the surface; this pin asserts the recursive
        walk.  We stub ``os.walk`` to return a synthetic tree
        with a top-level bill and a bill two folders deep.
        """
        mock_pdf = MagicMock()
        mock_page = MagicMock()
        mock_page.extract_text.return_value = "Current balance £500.00 debit"
        mock_pdf.pages = [mock_page]
        mock_pdf.__enter__.return_value = mock_pdf
        mock_pdf.__exit__.return_value = None
        mock_pdf_open.return_value = mock_pdf

        # Synthesised directory tree:
        #   /fake/path/
        #     ├── top.pdf
        #     └── 2023/
        #         └── bills/
        #             └── nested.pdf
        with (
            patch("os.path.exists", return_value=True),
            patch(
                "os.walk",
                return_value=[
                    ("/fake/path", ["2023"], ["top.pdf"]),
                    ("/fake/path/2023", ["bills"], []),
                    ("/fake/path/2023/bills", [], ["nested.pdf"]),
                ],
            ),
            patch("os.path.getmtime", return_value=1704067200),
            patch("builtins.open", mock_open(read_data=b"fake pdf content")),
        ):
            engine = self._make_engine()
            engine.crawl_local_pdfs("/fake/path")

        # Both surfaces of the directory tree must be reached.
        assert engine.pdf_count == 2, (
            f"recursive walk discovered only {engine.pdf_count} of 2 expected PDFs"
        )


# ---------------------------------------------------------------------------
# Stream P5 — engine.source_paths population (spec §3.9, issue 8b root cause)
# ---------------------------------------------------------------------------

# Minimal valid 1-page PDF (built via reportlab; reproduces the live
# evidence_files/ backdrop the user's bug report referenced). Kept inline
# so the test is fully self-contained.
_MINIMAL_PDF_B64 = (
    b"JVBERi0xLjMKJZOMi54gUmVwb3J0TGFiIEdlbmVyYXRlZCBQREYgZG9jdW1lbnQgKG9wZW5zb3VyY2Up"
    b"CjEgMCBvYmoKPDwKL0YxIDIgMCBSCj4+CmVuZG9iagoyIDAgb2JqCjw8Ci9CYXNlRm9udCAvSGVsdmV0aWNh"
    b"IC9FbmNvZGluZyAvV2luQW5zaUVuY29kaW5nIC9OYW1lIC9GMSAvU3VidHlwZSAvVHlwZTEgL1R5cGUg"
    b"L0ZvbnQKPj4KZW5kb2JqCjMgMCBvYmoKPDwKL0NvbnRlbnRzIDcgMCBSIC9NZWRpYUJveCBbIDAgMCA1"
    b"OTUuMjc1NiA4NDEuODg5OCBdIC9QYXJlbnQgNiAwIFIgL1Jlc291cmNlcyA8PAovRm9udCAxIDAgUiAv"
    b"UHJvY1NldCBbIC9QREYgL1RleHQgL0ltYWdlQiAvSW1hZ2VDIC9JbWFnZUkgXQo+PiAvUm90YXRlIDAg"
    b"IC9UcmFucyA8PAo+PiAKICAvVHlwZSAvUGFnZQo+PgplbmRvYmoKNCAwIG9iago8PAovUGFnZU1vZGUg"
    b"L1VzZU5vbmUgL1BhZ2VzIDYgMCBSIC9UeXBlIC9DYXRhbG9nCj4+CmVuZG9iago1IDAgb2JqCjw8Ci9B"
    b"dXRob3IgKGFub255bW91cykgL0NyZWF0aW9uRGF0ZSAoRDoyMDI2MDcyNTEzMTIyNSswMScwMCcpIC9D"
    b"cmVhdG9yIChhbm9ueW1vdXMpIC9LZXl3b3JkcyAoKSAvTW9kRGF0ZSAoRDoyMDI2MDcyNTEzMTIyNSsw"
    b"MScwMCcpIC9Qcm9kdWNlciAoUmVwb3J0TGFiIFBERiBMaWJyYXJ5IC0gXChvcGVuc291cmNlXCkpIAog"
    b"IC9TdWJqZWN0ICh1bnNwZWNpZmllZCkgL1RpdGxlICh1bnRpdGxlZCkgL1RyYXBwZWQgL0ZhbHNlCj4+"
    b"CmVuZG9iago2IDAgb2JqCjw8Ci9Db3VudCAxIC9LaWRzIFsgMyAwIFIgXSAvVHlwZSAvUGFnZXMKPj4K"
    b"ZW5kb2JqCjcgMCBvYmoKPDwKL0ZpbHRlciBbIC9BU0NJSTg1RGVjb2RlIC9GbGF0ZURlY29kZSBdIC9M"
    b"ZW5ndGggMTQ3Cj4+CnN0cmVhbQpHYXBARVltUz8lJ0xoYkZgPlIybG8wVC5fQlE4IlRuNCcqPTlx"
    b"W1o0UCFuZzIpYXVDUiUlImpIOVRBczAvVD1gY11jXCpgODosMSM0Jy06YS9xNmxbVHJaXnRoOUNXYiku"
    b"biZfQ0hNSWwuSVo/ZTBzUlQ4XFheXzVrc21yQzdyYyojNG8uSlVoYjBDSS5oM10pfj5lbmRzdHJlYW0K"
    b"ZW5kb2JqCnhyZWYKMCA4CjAwMDAwMDAwMDAgNjU1MzUgZiAKMDAwMDAwMDA2MSAwMDAwMCBuIAowMDAw"
    b"MDAwMDkyIDAwMDAwIG4gCjAwMDAwMDAxOTkgMDAwMDAgbiAKMDAwMDAwMDQwMiAwMDAwMCBuIAowMDAw"
    b"MDAwMDQ3MCAwMDAwMCBuIAowMDAwMDAwMDczMSAwMDAwMCBuIAowMDAwMDAwMDc5MCAwMDAwMCBuIAp0"
    b"cmFpbGVyCjw8Ci9JRCAKWzw2YTM1NTI2MjYxZTZlODlmN2UyZGI0ZmVlZjY3OWYwMj48NmEzNTUyNjI2"
    b"MWU2ZTg5ZjdlMmRiNGZlZWY2NzlmMDI+XQolIFJlcG9ydExhYiBnZW5lcmF0ZWQgUERGIGRvY3VtZW50"
    b"IC0tIGRpZ2VzdCAob3BlbnNvdXJjZSkKCi9JbmZvIDUgMCBSCi9Sb290IDQgMCBSCi9TaXplIDgKPj4K"
    b"c3RhcnR4cmVmCjEwMjcKJSVFT0YK"
)


def _write_minimal_pdf(tmp_path: Path, name: str = "test-invoice.pdf") -> str:
    import base64

    p = tmp_path / name
    p.write_bytes(base64.b64decode(_MINIMAL_PDF_B64))
    return str(p)


def test_evidence_engine_init_has_source_paths_dict() -> None:
    """EvidenceEngine.__init__ must initialise self.source_paths to an
    empty dict so save_evidence_files can find files via
    getattr(engine, "source_paths", {}). Spec §3.9 (issue 8b root cause)."""
    eng = EvidenceEngine(config={}, update_ui_cb=lambda *a, **k: None)
    assert hasattr(eng, "source_paths"), "EvidenceEngine.source_paths missing"
    assert isinstance(eng.source_paths, dict), "source_paths must be a dict"
    assert eng.source_paths == {}, "source_paths must start empty"


def test_process_pdf_file_populates_source_paths(tmp_path: Path) -> None:
    """process_pdf_file must record the path under
    self.source_paths[attachment_name] so save_evidence_files can find it.
    Spec §3.9 (issue 8b root cause)."""
    pdf = _write_minimal_pdf(tmp_path, "test-invoice.pdf")
    eng = EvidenceEngine(config={"acc_num": ""}, update_ui_cb=lambda *a, **k: None)
    eng.process_pdf_file(pdf, "Local PDF Folder", "test-invoice.pdf", "01/01/2024")
    assert "test-invoice.pdf" in eng.source_paths, eng.source_paths
    assert eng.source_paths["test-invoice.pdf"] == pdf


def test_process_pdf_file_explicit_attachment_name_recorded(tmp_path: Path) -> None:
    """When the caller passes attachment_name explicitly (PST emails),
    source_paths must use that name, not detail_label. Spec §3.9."""
    pdf = _write_minimal_pdf(tmp_path, "stub.pdf")
    eng = EvidenceEngine(config={}, update_ui_cb=lambda *a, **k: None)
    eng.process_pdf_file(
        pdf,
        "PST PDF Attachment",
        "stub.pdf",
        "01/01/2024",
        attachment_name="edf-invoice-KI-1234-0001-3.pdf",
    )
    assert "edf-invoice-KI-1234-0001-3.pdf" in eng.source_paths
    assert "stub.pdf" not in eng.source_paths


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
