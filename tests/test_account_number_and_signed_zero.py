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

from edf_collector import extract_new_credit_fields, extract_new_invoice_fields
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
