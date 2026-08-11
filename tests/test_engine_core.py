"""Tests for EvidenceEngine core methods to improve coverage."""

import pytest

from edf_bill_fetcher.collectors.engine import EvidenceEngine
from edf_bill_fetcher.models.config import ConfigDict


class TestEvidenceEngineCore:
    """Tests for EvidenceEngine core processing methods."""

    def test_find_billing_period(self):
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 500.0,
            "analysis_min": 500.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        engine = EvidenceEngine(config, lambda x: None)

        period_from, period_to = engine.find_billing_period(
            "Your charges: 01 Jan 2024 - 31 Jan 2024"
        )
        assert period_from == "01/01/2024"
        assert period_to == "31/01/2024"

        period_from, period_to = engine.find_billing_period("No period here")
        assert period_from == "N/A"
        assert period_to == "N/A"

    def test_add_record_with_filter(self):
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 500.0,
            "analysis_min": 500.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        engine = EvidenceEngine(config, lambda x: None)

        engine._add_record(
            {
                "Source": "Test",
                "Date": "01/01/2024",
                "Amount (£)": 100.0,  # Below min_amount
                "Details": "Test record",
                "Logic Used": "Test",
            }
        )

        assert len(engine.records) == 0
        assert len(engine.filtered_records) == 1
        assert engine.filtered_records[0]["Reason"] == "Amount magnitude below £500.00 threshold"

    def test_add_record_above_filter(self):
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 500.0,
            "analysis_min": 500.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        engine = EvidenceEngine(config, lambda x: None)

        engine._add_record(
            {
                "Source": "Test",
                "Date": "01/01/2024",
                "Amount (£)": 1000.0,  # Above min_amount
                "Details": "Test record",
                "Logic Used": "Test",
            }
        )

        assert len(engine.records) == 1
        assert len(engine.filtered_records) == 0

    def test_add_record_with_na_amount_does_not_crash(self):
        """M-1 regression: an uncoercible string amount (e.g. ``"N/A"``) must not
        crash ``_add_record``.  The amount becomes ``None``, the filter check is
        skipped, and the record is kept in main records."""
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 500.0,
            "analysis_min": 500.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        engine = EvidenceEngine(config, lambda x: None)

        engine._add_record(
            {
                "Source": "Test",
                "Date": "01/01/2024",
                "Amount (£)": "N/A",
                "Details": "Test record",
                "Logic Used": "Test",
            }
        )

        assert len(engine.records) == 1
        assert len(engine.filtered_records) == 0
        assert engine.records[0]["Amount (£)"] == "N/A"

    def test_add_record_coerces_numeric_string_amount(self):
        """M-1: a numeric string amount is coerced to float for the filter check,
        so a below-threshold numeric string is still filtered."""
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 500.0,
            "analysis_min": 500.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        engine = EvidenceEngine(config, lambda x: None)

        engine._add_record(
            {
                "Source": "Test",
                "Date": "01/01/2024",
                "Amount (£)": "100",  # below min_amount, but string-typed
                "Details": "Test record",
                "Logic Used": "Test",
            }
        )

        assert len(engine.records) == 0
        assert len(engine.filtered_records) == 1
        assert engine.filtered_records[0]["Amount (£)"] == 100.0

    def test_add_record_negative_amount_above_threshold_kept(self):
        """Regression test: a high-magnitude refund (e.g. ``-£1000``) must stay in
        main records when ``min_amount=500`` because ``abs(-1000) >= 500``.

        Pre-fix the comparison was ``amt < min_amount`` so ``-1000 < 500``
        filtered the refund out — losing valuable dispute evidence.
        """
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 500.0,
            "analysis_min": 500.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        engine = EvidenceEngine(config, lambda x: None)
        engine._add_record(
            {
                "Source": "PST",
                "Date": "01/01/2024",
                "Amount (£)": -1000.0,  # High-magnitude refund
                "Details": "Refund",
                "Logic Used": "Test refund",
            }
        )

        # Refund with abs(amt) >= min_amount stays in main records.
        assert len(engine.records) == 1
        assert engine.records[0]["Amount (£)"] == -1000.0
        assert len(engine.filtered_records) == 0

    def test_add_record_negative_amount_below_threshold_filtered(self):
        """A small-magnitude negative amount (e.g. ``-£5``) IS filtered when
        ``min_amount=500`` because ``abs(-5) < 500`` — small refunds of
        incidental credit balances aren't dispute evidence."""
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 500.0,
            "analysis_min": 500.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        engine = EvidenceEngine(config, lambda x: None)
        engine._add_record(
            {
                "Source": "PST",
                "Date": "01/01/2024",
                "Amount (£)": -5.0,  # Small-magnitude negative
                "Details": "Trivial refund",
                "Logic Used": "Test",
            }
        )

        assert len(engine.records) == 0
        assert len(engine.filtered_records) == 1
        assert engine.filtered_records[0]["Amount (£)"] == -5.0
        assert engine.filtered_records[0]["Reason"] == "Amount magnitude below £500.00 threshold"

    def test_is_cancelled(self):
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 500.0,
            "analysis_min": 500.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        import threading

        cancel_event = threading.Event()
        engine = EvidenceEngine(config, lambda x: None, cancel_event=cancel_event)

        assert not engine.is_cancelled()
        cancel_event.set()
        assert engine.is_cancelled()

    def test_log_error(self):
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 500.0,
            "analysis_min": 500.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        engine = EvidenceEngine(config, lambda x: None)

        engine.log_error("Test context", "Test error")
        assert len(engine.error_log) == 1
        assert "Test context" in engine.error_log[0]
        assert "Test error" in engine.error_log[0]


class TestEvidenceEngineProcessing:
    """Tests for EvidenceEngine process_text and related methods."""

    def test_process_text_basic(self):
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
        engine = EvidenceEngine(config, lambda x: None)

        text = "Current balance £750.00 debit"
        engine.process_text(text, "Email", "Test email", "01/01/2024")

        assert len(engine.records) == 1
        assert engine.records[0]["Amount (£)"] == 750.00
        assert engine.records[0]["Logic Used"] == "Smart Context"

    def test_process_text_with_account_filter_match(self):
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": True,
            "acc_num": "A-12345678",
            "min_amount": 100.0,
            "analysis_min": 100.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        engine = EvidenceEngine(config, lambda x: None)

        text = "Account number: A-12345678\nCurrent balance £750.00 debit"
        engine.process_text(text, "Email", "Test email", "01/01/2024")

        assert len(engine.records) == 1
        assert engine.records[0]["Amount (£)"] == 750.00

    def test_process_text_with_account_filter_no_match(self):
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": True,
            "acc_num": "A-99999999",  # Different account
            "min_amount": 100.0,
            "analysis_min": 100.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        engine = EvidenceEngine(config, lambda x: None)

        text = "Account number: A-12345678\nCurrent balance £750.00 debit"
        engine.process_text(text, "Email", "Test email", "01/01/2024")

        assert len(engine.records) == 0

    def test_process_text_large_fallback(self):
        config: ConfigDict = {
            "use_anchors": False,  # Disable anchors
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
        engine = EvidenceEngine(config, lambda x: None)

        text = "Some text with £500.00 and £600.00 amounts"
        engine.process_text(text, "Email", "Test email", "01/01/2024")

        assert len(engine.records) == 1
        assert engine.records[0]["Amount (£)"] == 600.00  # Max above threshold
        assert engine.records[0]["Logic Used"] == "Large Amount Fallback"

    def test_process_text_no_amount_found(self):
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
        engine = EvidenceEngine(config, lambda x: None)

        text = "No amounts in this text at all"
        engine.process_text(text, "Email", "Test email", "01/01/2024")

        assert len(engine.records) == 0

    def test_process_text_reading_classification(self):
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
        engine = EvidenceEngine(config, lambda x: None)

        text = "Current balance £750.00 debit\nEstimated reading"
        engine.process_text(text, "Email", "Test email", "01/01/2024")

        assert len(engine.records) == 1
        assert engine.records[0]["Reading"] == "Estimated"

    def test_process_text_pdf_fields_extraction(self):
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
        engine = EvidenceEngine(config, lambda x: None)

        text = "Current balance £750.00 debit\n500 kWh\n25.50p per day"
        engine.process_text(text, "PDF", "Test PDF", "01/01/2024")

        assert len(engine.records) == 1
        assert engine.records[0]["Units (kWh)"] == "500"
        assert engine.records[0]["Standing Chg (p/day)"] == "25.50"

    def test_process_text_empty_text(self):
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
        engine = EvidenceEngine(config, lambda x: None)

        engine.process_text("", "Email", "Test", "01/01/2024")
        engine.process_text(None, "Email", "Test", "01/01/2024")

        assert len(engine.records) == 0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
