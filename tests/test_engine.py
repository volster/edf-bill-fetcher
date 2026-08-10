"""Tests for EvidenceEngine config handling and record filtering."""

import pytest

from edf_bill_fetcher.collectors.engine import EvidenceEngine
from edf_bill_fetcher.models.config import ConfigDict


class TestEvidenceEngineConfig:
    """Tests for EvidenceEngine configuration handling."""

    def test_config_defaults(self):
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
        assert engine.config["min_amount"] == 500.0
        assert engine.config["use_anchors"] is True
        assert engine.config["filter_below"] is True

    def test_filter_below_min_amount(self):
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 1000.0,
            "analysis_min": 500.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        captured = []

        def capture_ui(msg):
            captured.append(msg)

        engine = EvidenceEngine(config, capture_ui)
        engine._add_record(
            {
                "Source": "Test",
                "Date": "01/01/2024",
                "Amount (£)": 500.0,  # Below min_amount
                "Details": "Test record",
                "Logic Used": "Test",
            }
        )
        # Should be filtered out
        assert len(engine.records) == 0
        assert len(engine.filtered_records) == 1
        assert engine.filtered_records[0]["Reason"] == "Amount magnitude below £1,000.00 threshold"

    def test_no_filter_when_disabled(self):
        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 1000.0,
            "analysis_min": 500.0,
            "filter_below": False,  # Disabled
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
                "Amount (£)": 500.0,  # Below min_amount but filter disabled
                "Details": "Test record",
                "Logic Used": "Test",
            }
        )
        # Should NOT be filtered out
        assert len(engine.records) == 1
        assert len(engine.filtered_records) == 0

    def test_thread_safety_lock(self):
        import threading

        config: ConfigDict = {
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "use_acc_filter": False,
            "acc_num": "",
            "min_amount": 0.0,
            "analysis_min": 500.0,
            "filter_below": True,
            "save_filtered": True,
            "use_dedup": True,
            "save_dups": True,
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
        }
        engine = EvidenceEngine(config, lambda x: None)

        def add_records(count):
            for i in range(count):
                engine._add_record(
                    {
                        "Source": "Test",
                        "Date": "01/01/2024",
                        "Amount (£)": float(i),
                        "Details": f"Record {i}",
                        "Logic Used": "Test",
                    }
                )

        threads = [threading.Thread(target=add_records, args=(100,)) for _ in range(10)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        assert len(engine.records) == 1000


class TestDomainFilter:
    """Tests for sender email domain filtering."""

    def test_domain_filter_exact_match(self):
        from edf_bill_fetcher.collectors.engine import _matches_domain_filter

        assert _matches_domain_filter("billing@edfenergy.com", "edfenergy.com")

    def test_domain_filter_subdomain(self):
        from edf_bill_fetcher.collectors.engine import _matches_domain_filter

        assert _matches_domain_filter("alerts@notifications.edfenergy.com", "edfenergy.com")

    def test_domain_filter_wildcard(self):
        from edf_bill_fetcher.collectors.engine import _matches_domain_filter

        assert _matches_domain_filter("billing@edfenergy.com", "*.edfenergy.com")

    def test_domain_filter_full_email(self):
        from edf_bill_fetcher.collectors.engine import _matches_domain_filter

        assert _matches_domain_filter("billing@edfenergy.com", "billing@edfenergy.com")

    def test_domain_filter_no_match(self):
        from edf_bill_fetcher.collectors.engine import _matches_domain_filter

        assert not _matches_domain_filter("spam@gmail.com", "edfenergy.com")

    def test_domain_filter_empty(self):
        from edf_bill_fetcher.collectors.engine import _matches_domain_filter

        assert not _matches_domain_filter("", "edfenergy.com")
        assert not _matches_domain_filter("billing@edfenergy.com", "")


class TestSenderEmailExtraction:
    """Tests for extracting sender email from PST messages."""

    def test_extract_from_headers(self):
        from edf_bill_fetcher.collectors.engine import _extract_sender_email

        class MockMsg:
            def get_transport_headers(self):
                return b'From: "EDF Energy" <billing@edfenergy.com>\nTo: user@example.com\nSubject: Your Bill'

            def get_sender_name(self):
                return "EDF Energy"

        msg = MockMsg()
        assert _extract_sender_email(msg) == "billing@edfenergy.com"

    def test_extract_from_sender_name(self):
        from edf_bill_fetcher.collectors.engine import _extract_sender_email

        class MockMsg:
            def get_transport_headers(self):
                return None

            def get_sender_name(self):
                return "EDF Energy <billing@edfenergy.com>"

        msg = MockMsg()
        assert _extract_sender_email(msg) == "billing@edfenergy.com"

    def test_no_email_found(self):
        from edf_bill_fetcher.collectors.engine import _extract_sender_email

        class MockMsg:
            def get_transport_headers(self):
                return b'From: "EDF Energy"\nSubject: Your Bill'

            def get_sender_name(self):
                return "EDF Energy"

        msg = MockMsg()
        assert _extract_sender_email(msg) == ""


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
