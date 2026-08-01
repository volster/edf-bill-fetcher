"""Regression tests for two financial-output defects surfaced during
real-data review.

The fixtures here do NOT use any real EDF account numbers, real
amounts, or real customer identifiers. All numbers and strings below
are deliberately synthetic.

DEFECTS PINNED
==============

1. ``fmt_money`` rendering signed-zero as ``"-0.00"``.

   Pre-fix ``edf_report.fmt_money(-0.001)`` returned ``"£-0.00"``.
   On the Financial Summary page that produced the visually wrong
   "Total Payments/Credits £-0.00" line. The fix coerces any
   value whose magnitude rounds to 0.00 at 2-dp to plain zero
   before formatting.

2. ``extract_new_invoice_fields`` / ``extract_new_credit_fields``
   only matching the compact ``A-NNNNNNNN`` account-number form.

   Pre-fix the regex ``Account number:`` followed by a literal ``A-``
   prefix and digit-only body failed on real EDF invoices that render
   the account number as spaced digits (``"Account number: 671 078
   701 920"``). The fix matches both renderings. The synthetic strings
   below were shaped after the public EDF format but use placeholder
   digits so no real account numbers leak into the test suite.
"""

from __future__ import annotations

from edf_bill_fetcher.processors.sap_parsers import (
    extract_new_credit_fields,
    extract_new_invoice_fields,
)
from edf_report import fmt_money


class TestFmtMoneySignedZero:
    """Signed-zero elimination for the Financial Summary page."""

    def test_zero_zero(self):
        assert fmt_money(0.0) == "£0.00"

    def test_zero_int(self):
        assert fmt_money(0) == "£0.00"

    def test_minus_zero_float(self):
        # -0.0 is signed-zero: pre-fix returned "£-0.00".
        assert fmt_money(-0.0) == "£0.00"

    def test_tiny_negative_rounds_zero(self):
        # -0.001 rounded to 2-dp becomes -0.00; pre-fix returned that.
        assert fmt_money(-0.001) == "£0.00"

    def test_tiny_positive_rounds_zero(self):
        assert fmt_money(0.00499) == "£0.00"

    def test_negative_one_penny_rounds_to_one_penny(self):
        # -0.005 rounds exactly to 0.01, NOT 0.00 — guard boundary.
        assert fmt_money(-0.005) == "£-0.01"

    def test_normal_negative_still_negative(self):
        assert fmt_money(-240.50) == "£-240.50"

    def test_normal_positive_still_positive(self):
        assert fmt_money(240.50) == "£240.50"

    def test_string_zero(self):
        assert fmt_money("0") == "£0.00"

    def test_string_neg_tiny(self):
        assert fmt_money("-0.001") == "£0.00"


class TestAccountNumberSpacedForm:
    """Account-number regex must accept both EDF renderings."""

    KI_HEADER = (
        "Your VAT invoice\n"
        "Invoice number: KI-0000000-0000\n"
        "Date issued: 1 March 2026\n"
        "Your charges: 1 February 2026 - 28 February 2026\n"
        "Total charges for this period £240.50 credit\n"
        "Current balance £240.50 credit\n"
    )

    KCR_HEADER = (
        "Credit note\n"
        "Credit note number: KCR-0000000-0000\n"
        "Date issued: 1 March 2026\n"
        "Total credits for this bill £10.00\n"
    )

    def test_invoice_compact_account_number(self):
        text = self.KI_HEADER + "Account number: A-0000000\n"
        fields = extract_new_invoice_fields(text)
        assert fields.get("acc_num") == "A-0000000"

    def test_invoice_spaced_account_number(self):
        # The pre-fix regex dropped this entire body — which is the
        # whole reason the report lost the account reference field.
        text = self.KI_HEADER + "Account number: 601 234 567 890\n"
        fields = extract_new_invoice_fields(text)
        assert fields.get("acc_num") == "601 234 567 890"

    def test_invoice_spaced_account_number_tight_spaces(self):
        # Some EDF invoices come with single-space separation; either
        # form must match.
        text = self.KI_HEADER + "Account number: 671 078 701 920\n"
        fields = extract_new_invoice_fields(text)
        assert fields.get("acc_num") == "671 078 701 920"

    def test_invoice_missing_account_number(self):
        text = self.KI_HEADER  # no account line at all
        fields = extract_new_invoice_fields(text)
        assert "acc_num" not in fields

    def test_credit_compact_account_number(self):
        text = self.KCR_HEADER + "Account number: A-0000000\n"
        fields = extract_new_credit_fields(text)
        assert fields.get("acc_num") == "A-0000000"

    def test_credit_spaced_account_number(self):
        text = self.KCR_HEADER + "Account number: 601 234 567 890\n"
        fields = extract_new_credit_fields(text)
        assert fields.get("acc_num") == "601 234 567 890"


class TestAccountNumberFilter:
    """Phase 1.3 — ``--acc-filter`` must reject substring-of-longer-string.

    Pre-fix the engine's account-filter was a plain
    ``acc_numeric in text_stripped`` substring check on digits-only
    forms.  An account-number config of ``"31"`` would false-match a
    meter-serial ``A-31000001`` because "31" is contained inside
    "31000001".  The post-fix helper
    ``account_number_matches(acc, text)`` rejects those substring
    matches while still letting a real standalone account number
    through (the same shape EDF renders as ``A-NNNNNNNN`` or as a
    spaced digit run).
    """

    def test_substring_of_longer_number_must_not_match(self):
        """The exact regression: short-account-vs-meter-serial."""
        from edf_bill_fetcher.collectors.engine import account_number_matches

        # Config: filtering for the account "31".
        # Suspicious text: meter serial A-31000001 — old logic
        # accepted because "31" was a substring of "31000001".
        text = "Your bill\nMeter serial: A-31000001\nCurrent balance £240.50 debit\n"
        assert account_number_matches("31", text) is False, (
            "AccountNumberFilter let '31' match 'A-31000001' — "
            "substring false-positive; pre-fix bug not actually fixed."
        )

    def test_legitimate_standalone_account_still_matches(self):
        """The same shorter config NEVER drops a real standalone hit."""
        from edf_bill_fetcher.collectors.engine import account_number_matches

        text = "Your bill\nAccount number: A-12345678\nCurrent balance £240.50 debit\n"
        assert account_number_matches("12345678", text) is True

    def test_legitimate_account_long_config_runs(self):
        """Filter configured as the full EDF ``A-NNNNNNNN`` form."""
        from edf_bill_fetcher.collectors.engine import account_number_matches

        text = "Your bill\nAccount number: A-12345678\nCurrent balance £240.50 debit\n"
        assert account_number_matches("A-12345678", text) is True

    def test_substring_match_with_digit_on_one_side_does_not_match(self):
        """``"12345678"`` must NOT match inside ``"A12345678B123``
        because — even though the first run ends at ``B`` and matches
        — the second digit context isn't the issue, we are pinning
        the *digit-bounded* invariant: the helper must NOT match
        when the digit-run immediately adjacent to the account
        number continues into another digit.

        Concretely: text ``"A1234567890"`` with filter ``"12345678"``
        must NOT match — "12345678" is followed by "90" (more digits).
        """
        from edf_bill_fetcher.collectors.engine import account_number_matches

        text = "Meter serial: A1234567890"
        assert account_number_matches("12345678", text) is False, (
            "AccountNumberFilter let '12345678' match 'A1234567890' — "
            "the right side of the candidate run continues with digits."
        )

    def test_substring_embedded_in_longer_serial_does_not_match(self):
        """Standard EDF-style false-positive: account config of all but
        the last digit of a phone/meter serial.  The longer run must
        not yield a match for the shorter prefix.
        """
        from edf_bill_fetcher.collectors.engine import account_number_matches

        text = "Phone: 07700 900123"
        assert account_number_matches("07700900", text) is False

    def test_genuine_standalone_run_matches_with_letter_bookend(self):
        """Bookend letters (not digits) DO NOT block a match."""
        from edf_bill_fetcher.collectors.engine import account_number_matches

        text = "Billed to: A12345678 "  # trailing space, no digit after
        assert account_number_matches("12345678", text) is True

    def test_empty_filter_matches_everything(self):
        """Empty-filter is the documented bypass — never silently reject."""
        from edf_bill_fetcher.collectors.engine import account_number_matches

        assert account_number_matches("", "anything goes here") is True
        assert account_number_matches("", "") is True

    def test_unusable_filter_passes_through(self):
        """A filter of just whitespace/letters shouldn't reject everything."""
        from edf_bill_fetcher.collectors.engine import account_number_matches

        # Defensive: if a user types a non-digit-only filter
        # by accident, we should not silently drop every record.
        assert account_number_matches("ab-cd", "any text") is True

    def test_engine_filter_through_helper_end_to_end(self):
        """Drive the helper through EvidenceEngine.process_text so a
        future refactor of the inner loop can't accidentally bypass it.

        We drive three end-to-end scenarios to pin the same contract
        the helper pins:
          * meter-serial substring-false-positive  → filter rejects
          * standalone account "31 555 4444"        → filter accepts,
            record added
          * a text with the *filter* embedded in a larger run
            like "311 555 4444"                       → filter rejects
            (the meter-serial case generalised to the realistic
            EDF format)
        """
        from edf_bill_fetcher.collectors.engine import EvidenceEngine

        def _build_engine(legacy_anchor=True):
            cfg = {
                "use_anchors": legacy_anchor,
                # Use the large fallback so even a "weak" anchor still
                # produces a record — the assertion target is whether
                # the filter passed/rejected, not the amount parsing.
                "use_large": True,
                "use_reading_classification": False,
                "use_pdf_fields": False,
                "min_amount": 1.0,
                "analysis_min": 1.0,
                "filter_below": False,
                "save_filtered": False,
                "use_dedup": False,
                "save_dups": False,
                "use_domain_filter": False,
                "use_acc_filter": True,
                "acc_num": "31",
            }
            return EvidenceEngine(cfg, lambda *_: None)

        # === Scenario 1: false-positive (meter-serial substring) ===
        engine = _build_engine()
        engine.process_text(
            "Meter serial: A-31000001\nCurrent balance £240.50 in debit",
            "PDF",
            "bill.pdf",
            "01/01/2024",
        )
        assert engine.records == [], (
            f"Meter-serial substring still caused an end-to-end match "
            f"({len(engine.records)} records); the engine-side filter isn't "
            f"funneling through the Phase 1.3 helper."
        )

        # === Scenario 2: genuine standalone account ===
        engine = _build_engine()
        engine.process_text(
            "Account number: 31 555 4444\nCurrent balance £240.50 in debit",
            "PDF",
            "bill2.pdf",
            "01/01/2024",
        )
        assert len(engine.records) >= 1, (
            "Genuine standalone account was dropped by the new "
            "boundary-imposed filter; the helper is too tight."
        )

        # === Scenario 3: account "31" embedded in a longer digit run ===
        engine = _build_engine()
        engine.process_text(
            "Account number: 311 555 4444\nCurrent balance £240.50 in debit",
            "PDF",
            "bill3.pdf",
            "01/01/2024",
        )
        assert engine.records == [], (
            "Filter accepted account '31' inside a longer digit run "
            "'311 555 4444' (no digit boundary on the left side); "
            "the helper dropped the digit-boundary invariant."
        )
