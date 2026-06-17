"""Tests for EDF bill extraction patterns and helpers."""

import re

# Import the functions from the main module
import sys
from datetime import datetime

import pytest

sys.path.insert(0, "C:/Users/matthew/edf-bill-fetcher")

from edf_collector import (
    _ISO_DATE_RE,
    AMOUNT_PATTERNS,
    PERIOD_RE,
    READING_PATTERNS,
    detect_pdf_format,
    extract_new_credit_fields,
    extract_new_invoice_fields,
    parse_to_display_date,
    parse_to_sort_date,
    to_excel_date,
)


class TestDateHelpers:
    """Tests for date parsing and conversion functions."""

    def test_parse_to_sort_date_iso_format(self):
        assert parse_to_sort_date("2024-01-15") == datetime(2024, 1, 15)

    def test_parse_to_sort_date_uk_format(self):
        assert parse_to_sort_date("15 Jan 2024") == datetime(2024, 1, 15)

    def test_parse_to_sort_date_uk_format_with_slashes(self):
        assert parse_to_sort_date("15/01/2024") == datetime(2024, 1, 15)

    def test_parse_to_sort_date_empty_string(self):
        assert parse_to_sort_date("") is not None  # returns NaT

    def test_parse_to_sort_date_none(self):
        assert parse_to_sort_date(None) is not None  # returns NaT

    def test_parse_to_sort_date_unknown(self):
        assert parse_to_sort_date("Unknown") is not None  # returns NaT

    def test_parse_to_display_date_iso(self):
        assert parse_to_display_date("2024-01-15") == "15/01/2024"

    def test_parse_to_display_date_uk(self):
        assert parse_to_display_date("15 Jan 2024") == "15/01/2024"

    def test_to_excel_date_valid(self):
        result = to_excel_date("15 Jan 2024")
        assert isinstance(result, datetime)
        assert result == datetime(2024, 1, 15)

    def test_to_excel_date_invalid(self):
        assert to_excel_date("invalid") is None
        assert to_excel_date("") is None


class TestPDFFormatDetection:
    """Tests for detecting EDF PDF bill formats."""

    def test_detect_new_invoice(self):
        text = "Invoice number: KI-12345678\nAccount number: A-12345678"
        assert detect_pdf_format(text) == "new_invoice"

    def test_detect_new_credit(self):
        text = "Credit note number: KCR-12345678\nAccount number: A-12345678"
        assert detect_pdf_format(text) == "new_credit"

    def test_detect_old_format(self):
        text = "Your new account balance £1,234.56"
        assert detect_pdf_format(text) == "old"

    def test_detect_case_insensitive(self):
        text = "invoice number: ki-12345678"
        assert detect_pdf_format(text) == "new_invoice"

        text = "credit note number: kcr-12345678"
        assert detect_pdf_format(text) == "new_credit"


class TestNewInvoiceExtraction:
    """Tests for extracting fields from new-style KI invoices."""

    def test_extract_new_invoice_basic(self):
        text = """
        Invoice number: KI-12345678
        Account number: A-12345678
        Date issued: 15 Jan 2024
        Your charges: 01 Jan 2024 - 31 Jan 2024
        Current balance £1,234.56 debit
        Total charges for this period £500.00 debit
        Electricity used 1,234 kWh
        Standing charge 31 days @ 45.5p/day
        Tariff name Standard Variable Payment type
        """
        fields = extract_new_invoice_fields(text)
        assert fields["inv_num"] == "KI-12345678"
        assert fields["acc_num"] == "A-12345678"
        assert fields["date"] == "15/01/2024"
        assert fields["period_from"] == "01/01/2024"
        assert fields["period_to"] == "31/01/2024"
        assert fields["amount"] == 1234.56
        assert fields["period_charge"] == 500.00
        assert fields["units_used"] == "1,234"
        assert fields["standing_charge"] == "45.5"
        assert fields["tariff"] == "Standard Variable"

    def test_extract_new_invoice_missing_optional(self):
        text = """
        Invoice number: KI-12345678
        Current balance £1,234.56 debit
        """
        fields = extract_new_invoice_fields(text)
        assert fields["inv_num"] == "KI-12345678"
        assert fields["amount"] == 1234.56
        assert "period_charge" not in fields
        assert "units_used" not in fields

    def test_extract_new_invoice_no_amount(self):
        text = "Invoice number: KI-12345678\nDate issued: 15 Jan 2024"
        fields = extract_new_invoice_fields(text)
        assert "amount" not in fields


class TestNewCreditExtraction:
    """Tests for extracting fields from new-style KCR credit notes."""

    def test_extract_new_credit_basic(self):
        text = """
        Credit note number: KCR-12345678
        Account number: A-12345678
        Date issued: 15 Jan 2024
        Total credits for this bill £500.00
        """
        fields = extract_new_credit_fields(text)
        assert fields["inv_num"] == "KCR-12345678"
        assert fields["acc_num"] == "A-12345678"
        assert fields["date"] == "15/01/2024"
        assert fields["amount"] == 500.00

    def test_extract_new_credit_missing_amount(self):
        text = "Credit note number: KCR-12345678"
        fields = extract_new_credit_fields(text)
        assert "amount" not in fields


class TestAmountPatterns:
    """Tests for the AMOUNT_PATTERNS regex list."""

    def test_pattern_current_balance_debit(self):
        text = "Current balance £1,234.56 debit"
        for p in AMOUNT_PATTERNS:
            m = re.search(p, text, re.IGNORECASE)
            if m:
                assert float(m.group(1).replace(",", "")) == 1234.56
                break
        else:
            pytest.fail("No pattern matched 'Current balance £1,234.56 debit'")

    def test_pattern_total_charges_period_debit(self):
        text = "Total charges for this period £500.00 debit"
        for p in AMOUNT_PATTERNS:
            m = re.search(p, text, re.IGNORECASE)
            if m:
                assert float(m.group(1).replace(",", "")) == 500.00
                break
        else:
            pytest.fail("No pattern matched 'Total charges for this period £500.00 debit'")

    def test_pattern_total_credits_bill(self):
        text = "Total credits for this bill £250.00"
        for p in AMOUNT_PATTERNS:
            m = re.search(p, text, re.IGNORECASE)
            if m:
                assert float(m.group(1).replace(",", "")) == 250.00
                break
        else:
            pytest.fail("No pattern matched 'Total credits for this bill £250.00'")

    def test_pattern_your_new_account_balance(self):
        text = "Your new account balance £800.00"
        for p in AMOUNT_PATTERNS:
            m = re.search(p, text, re.IGNORECASE)
            if m:
                assert float(m.group(1).replace(",", "")) == 800.00
                break
        else:
            pytest.fail("No pattern matched 'Your new account balance £800.00'")

    def test_pattern_balance_with_context(self):
        text = "Balance brought forward £1,000.00"
        for p in AMOUNT_PATTERNS:
            m = re.search(p, text, re.IGNORECASE)
            if m:
                assert float(m.group(1).replace(",", "")) == 1000.00
                break
        else:
            pytest.fail("No pattern matched 'Balance brought forward £1,000.00'")

    def test_pattern_pound_amount_debit(self):
        text = "£1,500.00 in debit"
        for p in AMOUNT_PATTERNS:
            m = re.search(p, text, re.IGNORECASE)
            if m:
                assert float(m.group(1).replace(",", "")) == 1500.00
                break
        else:
            pytest.fail("No pattern matched '£1,500.00 in debit'")


class TestReadingPatterns:
    """Tests for READING_PATTERNS classification."""

    def test_estimated_reading(self):
        for variant in ["estimated", "est.", "ESTIMATE", "Estimated reading"]:
            assert READING_PATTERNS["Estimated"].search(variant)

    def test_actual_reading(self):
        for variant in ["actual", "customer reading", "your reading", "ACTUAL"]:
            assert READING_PATTERNS["Actual"].search(variant)

    def test_smart_reading(self):
        for variant in ["smart meter", "automated reading", "smart reading", "SMART METER"]:
            assert READING_PATTERNS["Smart"].search(variant)


class TestPeriodRegex:
    """Tests for PERIOD_RE billing period extraction."""

    def test_period_dash_format(self):
        text = "01 Jan 2024 to 31 Jan 2024"
        m = PERIOD_RE.search(text)
        assert m is not None
        assert m.group(1).strip() == "01 Jan 2024"
        assert m.group(2).strip() == "31 Jan 2024"

    def test_period_en_dash(self):
        text = "01 Jan 2024 – 31 Jan 2024"
        m = PERIOD_RE.search(text)
        assert m is not None

    def test_period_slash_format(self):
        text = "01/01/2024 to 31/01/2024"
        m = PERIOD_RE.search(text)
        assert m is not None

    def test_period_hyphen_format(self):
        text = "01-01-2024 to 31-01-2024"
        m = PERIOD_RE.search(text)
        assert m is not None


class TestISODateRegex:
    """Tests for _ISO_DATE_RE."""

    def test_valid_iso_date(self):
        assert _ISO_DATE_RE.match("2024-01-15")
        assert _ISO_DATE_RE.match("2024-12-31")

    def test_invalid_formats(self):
        assert _ISO_DATE_RE.match("15/01/2024") is None
        assert _ISO_DATE_RE.match("15 Jan 2024") is None
        assert _ISO_DATE_RE.match("2024/01/15") is None


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
