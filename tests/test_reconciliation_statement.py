"""Tests for the multi-row reconciliation-statement detector + extractor."""

from edf_collector import (
    detect_reconciliation_statement,
    extract_reconciliation_statement_rows,
)

_PLAIN_INVOICE = """Solland Farm
EDF_0010
Bill reference: 28261421 (7 April 2026)
Account number: A-31105244
Your estimated energy bill
Balance on your last bill £37,301.48 debit
"""

_RECON_DOC = """Smell Gas - Immediately call 0800 111 999 (24hrs)
Solland Farm
Bill reference: 28261421 (7 April 2026)
Account number: A-31105244
EDF_0010
A
Your estimated energy bill
Balance on your last bill £37,301.48 debit
Charges
Electricity 14 May 2024 - 30 June 2024 £1,347.96
Electricity 1 July 2024 - 31 July 2024 £841.36 switching your tariff or supplier.
Electricity 14 May 2024 - 29 Aug. 2024 £980.28
Late Payment Charge £30.00
Reversed electricity charge 5 Sept. 2024 £841.36
Reversed electricity charge 7 July 2025 £823.47
(14 May 2024 - 30 Sept. 2024)
Payments
 1 April 2026 £850.00
 27 Feb. 2026 £850.00
 2 Feb. 2026 £850.00
Your new balance £45,332.13 debit
"""

_RECON_DOC_2 = """Some text
Bill reference: 28261421 (7 April 2026)
Account number: A-31105244
Balance on your last bill £37,301.48 debit
Electricity 1 Dec. 2025 - 31 Dec. 2025 £1,093.26
Electricity 1 Jan. 2026 - 31 Jan. 2026 £1,230.38
Reversed electricity charge 7 July 2025 £949.01
(1 Dec. 2024 - 31 Dec. 2024)
Late Payment Charge £30.00
Payments
 2 Feb. 2026 £850.00
 24 Dec. 2025 £850.00
Your new balance £45,332.13 debit
"""


def test_detector_true_on_recon_marker() -> None:
    assert detect_reconciliation_statement(_RECON_DOC) is True


def test_detector_still_true_without_payments_section() -> None:
    txt = (
        "Bill reference: 28261421 (7 April 2026)\n"
        "Account number: A-31105244\n"
        "Balance on your last bill £100.00 credit"
    )
    assert detect_reconciliation_statement(txt) is True


def test_detector_false_on_plain_invoice_without_bill_ref_match() -> None:
    """Plain local-PDF invoice lacks the Bill reference + Account number proximity marker."""
    txt = (
        "Solland Farm\n"
        "Invoice number: T12345\n"
        "Balance on your last bill £100.00 debit\n"
        "Account number: A-31105244"
    )
    assert detect_reconciliation_statement(txt) is False


def test_extractor_emits_charge_reversal_late_payment_rows() -> None:
    rows = extract_reconciliation_statement_rows(_RECON_DOC, "A-31105244-28261421-1-2.pdf")
    # 3 charges + 2 reversals + 1 late-payment + 3 payments + 1 statement-meta = 10
    assert len(rows) == 10
    # All rows carry the source label
    for row in rows:
        assert row["Source"] == "Statement Reconciliation"
        assert row["Attachment Name"] == "A-31105244-28261421-1-2.pdf"

    charge_rows = [r for r in rows if r["Entry Type"] == "Charge"]
    assert len(charge_rows) == 3
    # First charge: 14 May 2024 - 30 June 2024 £1,347.96
    assert charge_rows[0]["Period From"] == "14/05/2024"
    assert charge_rows[0]["Period To"] == "30/06/2024"
    assert charge_rows[0]["Amount (£)"] == 1347.96

    reversal_rows = [r for r in rows if r["Entry Type"] == "Credit"]
    assert len(reversal_rows) == 2
    # Reversal amounts are emitted as negative credits.
    assert reversal_rows[0]["Amount (£)"] == -841.36
    # First reversal is attached to a 5 Sept 2024 date.
    assert reversal_rows[0]["Date"] == "05/09/2024"
    # Second reversal has the parenthetical period captured into Details.
    assert reversal_rows[1]["Amount (£)"] == -823.47
    assert "14 May 2024 - 30 Sept. 2024" in reversal_rows[1]["Details"]

    late_rows = [r for r in rows if r["Entry Type"] == "Late Payment"]
    assert len(late_rows) == 1
    assert late_rows[0]["Amount (£)"] == 30.00

    pay_rows = [r for r in rows if r["Entry Type"] == "Payment"]
    # The first doc excerpt shows 3 payments.
    assert len(pay_rows) == 3
    assert pay_rows[0]["Date"] == "01/04/2026"
    assert pay_rows[0]["Amount (£)"] == 850.00

    meta_rows = [r for r in rows if r["Entry Type"] == "Statement Reconciliation"]
    assert len(meta_rows) == 1
    assert meta_rows[0]["Invoice #"] == "28261421"
    assert meta_rows[0]["Date"] == "07/04/2026"
    # Opening + closing balance captured.
    assert meta_rows[0]["Amount (£)"] == 45332.13
    assert meta_rows[0]["Balance Last Bill (£)"] == 37301.48


def test_extractor_meta_only_when_no_charges_or_reversals_or_payments() -> None:
    txt = (
        "Bill reference: 28261421 (7 April 2026)\n"
        "Account number: A-31105244\n"
        "Balance on your last bill £37,301.48 debit\n"
        "Your new balance £45,332.13 debit\n"
    )
    rows = extract_reconciliation_statement_rows(txt, "x.pdf")
    assert len(rows) == 1
    assert rows[0]["Entry Type"] == "Statement Reconciliation"


def test_extractor_second_doc_excerpt() -> None:
    rows = extract_reconciliation_statement_rows(_RECON_DOC_2, "x.pdf")
    # 2 charges + 1 reversal + 1 late-payment + 2 payments + 1 meta = 7
    assert len(rows) == 7
    reversals = [r for r in rows if r["Entry Type"] == "Credit"]
    assert len(reversals) == 1
    assert reversals[0]["Date"] == "07/07/2025"
    # Reversal amounts are emitted as negative credits.
    assert reversals[0]["Amount (£)"] == -949.01
    assert "1 Dec. 2024 - 31 Dec. 2024" in reversals[0]["Details"]


def test_extractor_handles_abbreviated_month_names() -> None:
    """Months like 'Sept.' or 'Aug.' appear in the recon PDF."""
    txt = """Bill reference: 12345678 (7 April 2026)
Account number: A-31105244
Electricity 1 Aug. 2024 - 30 Sept. 2024 £100.50
Your new balance £0.00 debit
"""
    rows = extract_reconciliation_statement_rows(txt, "x.pdf")
    charges = [r for r in rows if r["Entry Type"] == "Charge"]
    assert len(charges) == 1
    assert charges[0]["Period From"] == "01/08/2024"
    assert charges[0]["Period To"] == "30/09/2024"
    assert charges[0]["Amount (£)"] == 100.50


def test_extractor_no_balance_lines_handled_gracefully() -> None:
    """No balance lines → meta still emits with N/A balances."""
    txt = """Bill reference: 12345678 (7 April 2026)
Account number: A-31105244
Electricity 1 Aug. 2024 - 30 Sept. 2024 £100.50
"""
    rows = extract_reconciliation_statement_rows(txt, "x.pdf")
    metas = [r for r in rows if r["Entry Type"] == "Statement Reconciliation"]
    assert len(metas) == 1
    assert metas[0]["Balance Last Bill (£)"] == "N/A"
    assert metas[0]["Amount (£)"] == "N/A"
