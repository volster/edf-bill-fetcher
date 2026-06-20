"""Regression tests for the credit-balance handling in the new-style
KI/KCR invoice parser — closes the same gap as the #15 HTM parser fix
but for processed-by-`_process_new_invoice` PDFs.

Pre-fix:
    ``extract_new_invoice_fields`` hard-required the literal "debit"
    after both "Current balance" and "Total charges for this period".
    A credit-labelled statement silently dropped the amount.

Post-fix:
    Both regexes accept ``debit | credit | <omitted>`` and the parser
    populates the Amount and Period Charge columns in all three cases.

These tests pin the post-fix behaviour; if you remove the alternation
from the regex without thinking, two of these will fail.
"""

from __future__ import annotations

from edf_collector import extract_new_invoice_fields

KI_HEADER = (
    "Your VAT invoice\n"
    "Invoice number: KI-0000000-0000\n"
    "Account number: A-0000000\n"
    "Date issued: 1 March 2026\n"
    "Your charges: 1 February 2026 - 28 February 2026\n"
)


def _with_balance_clause(clause):
    """Compose a minimal KI body with the given balance+ period-charge
    clauses (no extra context).
    """
    return KI_HEADER + clause + "\nPlease pay this invoice by 10 March 2026.\n"


def test_balance_in_debit_still_parses():
    """Regression: the legacy ``debit`` path must still match."""
    fields = extract_new_invoice_fields(
        _with_balance_clause(
            "Total charges for this period £240.50 debit\nCurrent balance £240.50 debit\n"
        )
    )
    assert fields["amount"] == 240.50
    assert fields["period_charge"] == 240.50
    assert fields.get("amount_side") == "debit"


def test_balance_in_credit_is_now_parsed():
    """The #15-sibling fix: a credit-labelled statement should populate
    Amount and Period Charge the same as a debit-labelled one."""
    fields = extract_new_invoice_fields(
        _with_balance_clause(
            "Total charges for this period £240.50 credit\nCurrent balance £240.50 credit\n"
        )
    )
    assert fields["amount"] == 240.50, (
        f"credit-labelled Current balance should still populate amount, "
        f"got: {fields.get('amount')!r}"
    )
    assert fields["period_charge"] == 240.50, (
        f"credit-labelled period charge should still populate, got: {fields.get('period_charge')!r}"
    )
    assert fields.get("amount_side") == "credit"


def test_balance_with_no_side_label_parses():
    """A line with just ``Current balance GBPX`` (no debit|credit) is
    not common but legal — some older statements left it unqualified.
    Must still match.
    """
    fields = extract_new_invoice_fields(
        _with_balance_clause("Total charges for this period £240.50\nCurrent balance £240.50\n")
    )
    assert fields["amount"] == 240.50
    assert fields["period_charge"] == 240.50
    assert fields.get("amount_side") == ""


def test_amount_side_field_absent_when_regex_did_not_match():
    """If neither Current balance nor Total charges for this period
    appear in the body, the parser should still return a dict and just
    not include "amount" / "period_charge" / "amount_side".
    """
    fields = extract_new_invoice_fields("Your VAT invoice\nInvoice number: KI-1\n")
    # Other fields ARE expected (invoice number populated) but the
    # currency cells are not.
    assert "inv_num" in fields
    assert "amount" not in fields
    assert "period_charge" not in fields
    assert "amount_side" not in fields
