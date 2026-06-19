"""Tests for the EDF MyAccount HTM account-history parser.

LAYMAN'S GUIDE
==============

EDF lets customers download their payment history as a chunk of HTML
(``Payments and Invoices`` → save as... → ``.htm``). The page has one
line per event. A typical line looks like:

    "28 Feb 2026 We charged your account £1,070.48 For 2354 kWh of
     electricity used between 01 Feb 2026 and 28 Feb 2026
     Balance £46,182.13 in debit"

There are three flavours: a charge (EDF charged us), a payment
(we paid them back), and a reversed account charge (EDF issued us
a credit note for overpayment). All three end with a running balance.

The parser is ``edf_collector.parse_htm_account_history``. Its job
is to read the textual lines and emit structured ``dict`` records we
can feed into the Excel/PDF/DOCX reports.

These tests pin down what is currently parsed. They also flag
(**failure**) the credit-balance gap that is still open — see
``TestHTMCreditBalance`` below — so the next person to fix it has
a crisp RED→GREEN target.
"""

# We import from the project module via the absolute import
# (the test runner adds the project root to sys.path;
# see ``pyproject.toml`` ``[tool.pytest.ini_options] pythonpath``).

from edf_collector import parse_htm_account_history


class TestHTMParserCharge:
    """Charge entries: ``DD Mon YYYY We charged your account £X.XX …``.

    EDF charges us for electricity used in a billing period. The HTM
    export records each charge as a line with the period start, period
    end, units (kWh) and the running balance. Our parser turns each
    one into an ``Entry Type == "Ongoing Balance"`` record storing the
    charge amount under ``"Period Charge (£)"`` and the running
    balance under ``"Amount (£)"`` (because ``Amount`` is the column
    that downstream sorters pick on).
    """

    def test_charge_record_shape(self):
        # Happy path: full charge line with kWh and a balance.
        text = """
        28 Feb 2026 We charged your account £1,070.48 For 2354 kWh of electricity used between 01 Feb 2026 and 28 Feb 2026 Balance £46,182.13 in debit
        """

        records = parse_htm_account_history(text)

        assert len(records) == 1, "should produce exactly one charge record"
        r = records[0]
        assert r["Source"] == "HTM Account History"
        assert r["Date"] == "28/02/2026"
        # Charge amount lives in Period Charge (£); running balance in Amount (£).
        # (Layman: Amount = running balance, Period Charge = the bill itself.)
        assert r["Amount (£)"] == 46182.13
        assert r["Period Charge (£)"] == 1070.48
        # kWh consumption comes from the middle of the line.
        assert r["Units (kWh)"] == "2354"
        # Period dates.
        assert r["Period From"] == "01/02/2026"
        assert r["Period To"] == "28/02/2026"
        # Classifier for this row type.
        assert r["Entry Type"] == "Ongoing Balance"
        assert r["Logic Used"] == "HTM Charge"

    def test_charge_without_kwh_uses_na(self):
        # A rare variant: the line has no ``For N kWh ... between``. The
        # parser must still produce a record — fields fall back to "N/A".
        text = """
        28 Feb 2026 We charged your account £1,070.48 Balance £46,182.13 in debit
        """

        records = parse_htm_account_history(text)

        assert len(records) == 1
        r = records[0]
        assert r["Units (kWh)"] == "N/A"
        assert r["Period From"] == "N/A"
        assert r["Period To"] == "N/A"
        # The "Period Charge (£)" should still be the charge amount.
        assert r["Period Charge (£)"] == 1070.48


class TestHTMParserPayment:
    """Payment entries: ``DD Mon YYYY You paid us £X.XX … Balance £Y.YY in debit``.

    We paid EDF. The line carries no period or kWh detail — only the
    payment amount and the post-payment running balance. Our parser
    emits an ``Entry Type == "Payment"`` record with the running
    balance sitting under ``"Amount (£)"`` and the payment amount
    captured under ``"Period Charge (£)"`` (which is a slight
    misnomer in this context — there's no period — but the column is
    where downstream code looks for "this row's money").
    """

    def test_payment_record_shape(self):
        text = """
        27 Feb 2026 You paid us £850.00 Bank Transfer Balance £45,111.65 in debit
        """

        records = parse_htm_account_history(text)

        assert len(records) == 1
        r = records[0]
        assert r["Source"] == "HTM Account History"
        assert r["Date"] == "27/02/2026"
        # The running balance is captured as Amount (£).
        assert r["Amount (£)"] == 45111.65
        # No period info for a payment.
        assert r["Period Charge (£)"] == "N/A"
        assert r["Period From"] == "N/A"
        assert r["Period To"] == "N/A"
        # Classifier for this row type.
        assert r["Entry Type"] == "Payment"
        assert r["Logic Used"] == "HTM Payment"


class TestHTMParserReversal:
    """Reversed-account-charge entries: ``DD Mon YYYY Reversed account charge £X.XX …``.

    EDF issued a credit note. The line item is what flagged the credit;
    the running balance is therefore smaller than the prior row.
    The parser emits ``Entry Type == "Credit"``.
    """

    def test_reversal_record_shape(self):
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


class TestHTMParserMixed:
    """Mixed input: multiple lines, one each of charge / payment / reversal.

    Confirms the parser preserves order and doesn't lose any line.
    """

    def test_three_rows_in_order(self):
        text = """
        28 Feb 2026 We charged your account £1,070.48 For 2354 kWh of electricity used between 01 Feb 2026 and 28 Feb 2026 Balance £46,182.13 in debit
        27 Feb 2026 You paid us £850.00 Bank Transfer Balance £45,111.65 in debit
        26 Feb 2026 Reversed account charge £100.00 Refund Balance £44,011.65 in debit
        """

        records = parse_htm_account_history(text)

        assert len(records) == 3
        types = [r["Entry Type"] for r in records]
        assert types == ["Ongoing Balance", "Payment", "Credit"], (
            "expect charge first, then payment, then credit (matches input order)"
        )


class TestHTMParserEdgeCases:
    """Inputs that should produce no records at all."""

    def test_empty_text_yields_no_records(self):
        records = parse_htm_account_history("")
        assert records == []

    def test_unrelated_text_yields_no_records(self):
        text = "Some random text without EDF patterns whatsoever."
        records = parse_htm_account_history(text)
        assert records == []


class TestHTMCreditBalance:
    """Credit-balance handling — the known gap.

    BACKGROUND
    ----------
    All three current regexes (charge, payment, reversal) end with
    ``Balance £X in debit``. EDF's HTM export can also contain credit
    balances. A typical credit-balance line is:

        "DD Mon YYYY Balance £X.XX in credit"

    The current parser drops these silently — it returns no record,
    even though credit balances describe real-world money owed by EDF
    to the customer and may matter for an Ombudsman submission.

    These tests describe the desired behaviour. They currently FAIL.
    Removing the ``xfail`` marker (and adding the underlying parser
    logic) will mark them passing once the gap is closed.
    """

    import pytest  # local import so the rest of the file doesn't need pytest at module load

    @pytest.mark.xfail(
        reason="TODO(credit-balance): parser currently drops lines ending in 'in credit'",
        strict=True,
    )
    def test_charge_in_credit_is_parsed(self):
        # Charge line where the running balance is in credit (EDF
        # over-accounted in the past or this is an opening credit balance).
        text = """
        28 Feb 2026 We charged your account £50.00 For 100 kWh between 01 Feb 2026 and 28 Feb 2026 Balance £250.00 in credit
        """
        records = parse_htm_account_history(text)
        assert len(records) == 1
        r = records[0]
        assert r["Entry Type"] == "Ongoing Balance"
        # Balance carried into the Amount column regardless of debit/credit.
        assert r["Amount (£)"] == 250.0
        assert r["Period Charge (£)"] == 50.0

    @pytest.mark.xfail(
        reason="TODO(credit-balance): parser currently drops lines ending in 'in credit'",
        strict=True,
    )
    def test_payment_in_credit_is_parsed(self):
        # Standard payment line but balance in credit (after a refund).
        text = """
        05 Mar 2026 You paid us £200.00 Bank Transfer Balance £150.00 in credit
        """
        records = parse_htm_account_history(text)
        assert len(records) == 1
        r = records[0]
        assert r["Entry Type"] == "Payment"
        assert r["Amount (£)"] == 150.0

    @pytest.mark.xfail(
        reason="TODO(credit-balance): parser currently drops standalone 'Balance £X in credit' lines",
        strict=True,
    )
    def test_standalone_balance_in_credit_is_parsed(self):
        # A credit-balance line with no preceding charge/payment/reversal
        # verb. These appear at the top of an HTM export when the
        # customer's overall balance is in credit and there is no
        # transaction recorded for the period.
        text = """
        01 Jan 2026 Balance £123.45 in credit
        """
        records = parse_htm_account_history(text)
        assert len(records) == 1
        r = records[0]
        assert r["Amount (£)"] == 123.45


if __name__ == "__main__":
    import pytest

    pytest.main([__file__, "-v"])
