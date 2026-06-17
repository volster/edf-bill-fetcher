"""Tests for HTM account history parser."""

import sys

import pytest

sys.path.insert(0, "C:/Users/matthew/edf-bill-fetcher")

from edf_collector import parse_htm_account_history


class TestHTMParser:
    """Tests for parsing EDF MyAccount HTM exports."""

    def test_parse_charge_entries(self):
        text = """
        28 Feb 2026 We charged your account £1,070.48 For 2354 kWh of electricity used between 01 Feb 2026 and 28 Feb 2026 Balance £46,182.13 in debit
        """
        records = parse_htm_account_history(text)
        assert len(records) == 1
        r = records[0]
        assert r["Source"] == "HTM Account History"
        assert r["Date"] == "28/02/2026"
        assert r["Period From"] == "01/02/2026"
        assert r["Period To"] == "28/02/2026"
        assert r["Amount (£)"] == 46182.13
        assert r["Period Charge (£)"] == 1070.48
        assert r["Entry Type"] == "Ongoing Balance"
        assert r["Units (kWh)"] == "2354"
        assert r["Logic Used"] == "HTM Charge"

    def test_parse_payment_entries(self):
        text = """
        27 Feb 2026 You paid us £850.00 Bank Transfer Balance £45,111.65 in debit
        """
        records = parse_htm_account_history(text)
        assert len(records) == 1
        r = records[0]
        assert r["Source"] == "HTM Account History"
        assert r["Date"] == "27/02/2026"
        assert r["Amount (£)"] == 45111.65
        assert r["Entry Type"] == "Payment"
        assert r["Logic Used"] == "HTM Payment"

    def test_parse_reversal_entries(self):
        text = """
        26 Feb 2026 Reversed account charge £100.00 Some reason Balance £44,011.65 in debit
        """
        records = parse_htm_account_history(text)
        assert len(records) == 1
        r = records[0]
        assert r["Source"] == "HTM Account History"
        assert r["Date"] == "26/02/2026"
        assert r["Amount (£)"] == 44011.65
        assert r["Entry Type"] == "Credit"
        assert r["Logic Used"] == "HTM Reversal"

    def test_parse_multiple_entries(self):
        text = """
        28 Feb 2026 We charged your account £1,070.48 For 2354 kWh of electricity used between 01 Feb 2026 and 28 Feb 2026 Balance £46,182.13 in debit
        27 Feb 2026 You paid us £850.00 Bank Transfer Balance £45,111.65 in debit
        26 Feb 2026 Reversed account charge £100.00 Refund Balance £44,011.65 in debit
        """
        records = parse_htm_account_history(text)
        assert len(records) == 3
        assert records[0]["Entry Type"] == "Ongoing Balance"
        assert records[1]["Entry Type"] == "Payment"
        assert records[2]["Entry Type"] == "Credit"

    def test_parse_charge_without_kwh(self):
        text = """
        28 Feb 2026 We charged your account £1,070.48 Balance £46,182.13 in debit
        """
        records = parse_htm_account_history(text)
        assert len(records) == 1
        r = records[0]
        assert r["Units (kWh)"] == "N/A"
        assert r["Period From"] == "N/A"
        assert r["Period To"] == "N/A"

    def test_parse_empty_text(self):
        records = parse_htm_account_history("")
        assert records == []

    def test_parse_no_matches(self):
        text = "Some random text without EDF patterns"
        records = parse_htm_account_history(text)
        assert records == []


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
