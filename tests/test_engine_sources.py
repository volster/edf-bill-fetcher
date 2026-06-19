"""Tests for EvidenceEngine PDF/PST/HTM processing methods."""

from unittest.mock import MagicMock, Mock, mock_open, patch

import pytest

from edf_collector import EvidenceEngine


class TestEvidenceEnginePDF:
    """Tests for PDF processing methods."""

    def _make_engine(self):
        config = {
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

    @patch("edf_collector.pdfplumber.open")
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

    @patch("edf_collector.pdfplumber.open")
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

    @patch("edf_collector.pdfplumber.open")
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
            with patch("edf_collector.pdfplumber.open", side_effect=Exception("PDF read error")):
                engine.process_pdf_file("bad.pdf", "Test", "bad.pdf", "01/01/2024")

        assert len(engine.error_log) == 1
        assert "PDF read error" in engine.error_log[0]


class TestEvidenceEnginePST:
    """Tests for PST crawling."""

    def _make_engine(self):
        config = {
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
        config = {
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
        config = {
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

    @patch("edf_collector.pdfplumber.open")
    def test_crawl_local_pdfs(self, mock_pdf_open):
        mock_pdf = MagicMock()
        mock_page = MagicMock()
        mock_page.extract_text.return_value = "Current balance £500.00 debit"
        mock_pdf.pages = [mock_page]
        mock_pdf.__enter__.return_value = mock_pdf
        mock_pdf.__exit__.return_value = None
        mock_pdf_open.return_value = mock_pdf

        with (
            patch("os.path.exists", return_value=True),
            patch("os.listdir", return_value=["bill1.pdf", "bill2.pdf"]),
            patch("os.path.getmtime", return_value=1704067200),
            patch("builtins.open", mock_open(read_data=b"fake pdf content")),
        ):
            engine = self._make_engine()
            engine.crawl_local_pdfs("/fake/path")

        assert engine.pdf_count == 2


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
