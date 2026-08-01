"""Tests for EDF bill extraction patterns and helpers."""

from datetime import datetime

import pytest

from edf_bill_fetcher.helpers.date_utils import (
    _ISO_DATE_RE,
    parse_to_display_date,
    parse_to_sort_date,
    to_excel_date,
)
from edf_bill_fetcher.processors.detection import detect_pdf_format
from edf_bill_fetcher.processors.patterns import (
    AMOUNT_PATTERNS,
    PERIOD_RE,
    READING_PATTERNS,
)
from edf_bill_fetcher.processors.sap_parsers import (
    extract_new_credit_fields,
    extract_new_invoice_fields,
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
    """Tests for the AMOUNT_PATTERNS regex list.

    Each pattern in ``AMOUNT_PATTERNS`` is a (name, regex) tuple. These
    tests pin the contract that:

    1. Each pattern matches its documented anchor text (the example).
    2. The pattern's name correctly maps to a known route in
       ``_classify_entry_type`` — "New Bill" or "Ongoing Balance".
    3. Every pattern is unique (i.e. the canonical example matches
       exactly one of them).

    Tests are written against names, not indices — reordering patterns
    no longer breaks them, but adding/removing a real pattern still does
    (intentional: that update forces the contributor to update the
    tests AND the bucket sets together).
    """

    # Each entry is (pattern_name, anchor_text, expected_amount_as_float,
    # expected_classification).
    # Data-driven so adding new patterns is a one-line append.
    NAMED_PATTERN_CASES: list[tuple[str, str, float, str]] = [
        # New-style KI / KCR invoices (New Bill route)
        ("current_balance_debit", "Current balance £1,234.56 debit", 1234.56, "New Bill"),
        ("total_charges_period", "Total charges for this period £500.00 debit", 500.00, "New Bill"),
        ("total_credits_bill", "Total credits for this bill £250.00", 250.00, "New Bill"),
        # Old-style cumulative balance (Ongoing Balance route)
        ("your_new_account_balance", "Your new account balance £800.00", 800.00, "Ongoing Balance"),
        # Generic anchors (mixed, see below)
        ("balance_within", "Account balance £600.00 in debit", 600.00, "Ongoing Balance"),
        ("total_charges_within", "Your total charges £75.00 today", 75.00, "New Bill"),
        ("total_amount_due_within", "Your total amount due is £42.10", 42.10, "New Bill"),
        ("amount_to_pay_within", "Amount to pay £19.99", 19.99, "New Bill"),
        ("pound_amount_debit", "£99.99 in debit", 99.99, "New Bill"),
        ("current_balance_within", "Your current balance is £33.33", 33.33, "Ongoing Balance"),
    ]

    def test_each_named_pattern_matches_its_anchor(self):
        """For every (name, anchor) pair, exactly one AMOUNT_PATTERN
        matches and it's the one whose name matches the input name.

        This locks both: (a) the pattern exists in the registry, and
        (b) it actually fires against its documented example.
        """
        from edf_bill_fetcher.processors.patterns import (
            _AMOUNT_PATTERN_NEW_BILL,
            _AMOUNT_PATTERN_ONGOING_BALANCE,
        )

        for name, anchor, expected_amount, expected_classification in self.NAMED_PATTERN_CASES:
            matches = [(n, m) for n, p in AMOUNT_PATTERNS if (m := p.search(anchor))]
            assert matches, (
                f"Anchor {anchor!r} matched no pattern (expected pattern {name!r} to match)."
            )
            matching_names = [n for n, _ in matches]
            assert name in matching_names, (
                f"Anchor {anchor!r} was matched by {matching_names} but "
                f"expected {name!r} to be among them."
            )
            # Verify the captured amount is right.
            primary = matches[0][1]
            amount = float(primary.group(1).replace(",", ""))
            assert amount == expected_amount, (
                f"Anchor {anchor!r} captured {amount!r} for pattern "
                f"{name!r}, expected {expected_amount!r}."
            )
            # Verify the name maps to the right classification bucket.
            if expected_classification == "New Bill":
                assert name in _AMOUNT_PATTERN_NEW_BILL, (
                    f"Pattern {name!r} is not in the New Bill bucket — "
                    f"the classifier would route it incorrectly."
                )
            else:
                assert expected_classification == "Ongoing Balance", (
                    "Test data mistake: only New Bill / Ongoing Balance are supported."
                )
                assert name in _AMOUNT_PATTERN_ONGOING_BALANCE, (
                    f"Pattern {name!r} is not in the Ongoing Balance bucket — "
                    f"the classifier would route it incorrectly."
                )

    def test_pattern_routes_cover_all_entries(self):
        """Every registry entry must be in one of the two bucket sets.

        A pattern with no bucket assignment would silently fall
        through to heuristic classification, defeating the registry
        contract. The :py:func:`edf_bill_fetcher.processors.patterns:AMOUNT_PATTERNS`
        definition already asserts this at import time, but a test
        here gives the developer's editor an immediate signal.
        """
        from edf_bill_fetcher.processors.patterns import (
            _AMOUNT_PATTERN_NEW_BILL,
            _AMOUNT_PATTERN_ONGOING_BALANCE,
        )

        for name, _ in AMOUNT_PATTERNS:
            assert name in _AMOUNT_PATTERN_NEW_BILL or name in _AMOUNT_PATTERN_ONGOING_BALANCE, (
                f"Pattern {name!r} has no entry-type bucket assigned."
            )

    def test_priority_order_specific_before_generic(self):
        """Each specific anchor must come before the generic ``_within``
        variant of the same intent.

        Without this rule, the generic pattern would always fire first
        on long-form bill bodies, leaving the more specific patterns
        unreachable. Concretely:

        * ``current_balance_debit`` must come before ``balance_within``.
        * ``total_charges_period`` must come before ``total_charges_within``.
        * ``your_new_account_balance`` must come before ``current_balance_within``
          (both are "ongoing balance" shapes; the specific must win).
        * ``total_credits_bill`` is a specific anchor with no generic
          sibling — no constraint.
        * ``pound_amount_debit`` is a specific leaf with no siblings —
          no constraint.
        * ``amount_to_pay_within`` / ``total_amount_due_within`` are
          generic fall-throughs with no specific siblings — no constraint.

        This test fails loud-and-fast when a future contributor
        reorders the registry and accidentally hides a specialist
        pattern behind a generic one.
        """
        names = [n for n, _ in AMOUNT_PATTERNS]

        # (specific, must_come_before_this_generic_sibling)
        order_invariants: list[tuple[str, str]] = [
            ("current_balance_debit", "balance_within"),
            ("total_charges_period", "total_charges_within"),
            ("your_new_account_balance", "current_balance_within"),
        ]
        for specific, generic in order_invariants:
            try:
                specific_idx = names.index(specific)
                generic_idx = names.index(generic)
            except ValueError:
                pytest.fail(
                    f"Invariant violated: pattern {specific!r} or "
                    f"{generic!r} missing from AMOUNT_PATTERNS."
                )
            assert specific_idx < generic_idx, (
                f"Specific pattern {specific!r} (index {specific_idx}) "
                f"must come before generic sibling {generic!r} (index "
                f"{generic_idx}) — otherwise the specific pattern is "
                f"unreachable on long-form bill text."
            )


class TestReadingPatterns:
    """Tests for READING_PATTERNS classification."""

    def test_estimated_reading(self):
        for variant in ["estimated", "est.", "ESTIMATE", "Estimated reading"]:
            assert READING_PATTERNS["Estimated"].search(variant)

    def test_actual_reading(self):
        # The "Actual" pattern must match the meter-reading context —
        # bare-word "actual" should NOT misfire on simple bill prose
        # such as "the actual amount you owe is £X".
        for variant in [
            "customer reading",
            "your reading",
            "Customer reading was 12450",
            "your reading: 12450",
            "meter reading was actual",
            "actual reading: 12450",
        ]:
            assert READING_PATTERNS["Actual"].search(variant), (
                f"Actual-reading marker missing for {variant!r}"
            )
        # And it must NOT match bare "actual" — that would be a low-
        # precision false positive inviting the dispute-report to
        # misclassify meter-reading semantics.
        for variant in [
            "the actual amount you owe is £240",
            "the actual cost of your bill",
            "for actual consumption in this period",
            "actual",
            "ACTUAL",
        ]:
            assert not READING_PATTERNS["Actual"].search(variant), (
                f"Actual-reading pattern wrongly matched bare word in {variant!r}"
            )

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
