"""Tests for the multi-regex fallback chain (Task 5)."""

from edf_bill_fetcher.collectors.engine import (
    _fallback_amount,
    _fallback_inv_num,
    _fallback_period_from,
    _fallback_period_to,
)
from edf_bill_fetcher.processors.patterns import (
    _COVER_BLOCK_INV_RE,
    _COVER_BLOCK_PERIOD_RE,
    _FALLBACK_AMOUNT_RE,
    _FALLBACK_INV_RE,
)

_TEXT_NEW_INVOICE = """Solland Farm
Invoice number: KI-31105244-0001-3
Account number: A-31105244
Date issued: 14 May 2024
Your charges: 14 May 2024 - 30 June 2024
Total charges for this period £1,347.96 debit
"""

_TEXT_COVER_BLOCK = """Cover page
Some prefix text
Invoice number T78701920068 attached
for the period 14 May 2024 - 30 June 2024
Amount: £1,347.96
"""

_TEXT_KCR = """Statement
Credit note number: KCR-31105244-0001-3
bill period 14 May 2024 to 30 June 2024
Amount £450.75
"""

_TEXT_NO_HINTS = """This document has nothing useful.
Some random text without an invoice marker.
Just plain words.
"""


def test_fallback_inv_num_inv_number_re() -> None:
    val, label = _fallback_inv_num(_TEXT_NEW_INVOICE)
    assert val == "KI-31105244-0001-3"
    assert label == "_INV_NUMBER_RE"


def test_fallback_inv_num_cover_block_re() -> None:
    val, label = _fallback_inv_num(_TEXT_COVER_BLOCK)
    assert val == "T78701920068"
    assert label == "_COVER_BLOCK_INV_RE"


def test_fallback_inv_num_fallthrough_net_re() -> None:
    val, label = _fallback_inv_num(_TEXT_KCR)
    # The ``Credit note number: KCR-`` regex matches first, ahead of the
    # loose bare-token fallback ``_FALLBACK_INV_RE``.
    assert val == "KCR-31105244-0001-3"
    assert label == "_CREDIT_NUMBER_RE"


def test_fallback_inv_num_returns_none_when_no_match() -> None:
    val, label = _fallback_inv_num(_TEXT_NO_HINTS)
    assert val is None
    assert label == ""


def test_fallback_period_from_picks_billing_period_re() -> None:
    val, label = _fallback_period_from(_TEXT_NEW_INVOICE)
    assert val == "14 May 2024"
    assert label == "_BILLING_PERIOD_RE"


def test_fallback_period_from_picks_cover_block_re() -> None:
    val, label = _fallback_period_from(_TEXT_COVER_BLOCK)
    assert val == "14 May 2024"
    assert label == "_COVER_BLOCK_PERIOD_RE"


def test_fallback_period_from_picks_kcr_bill_period() -> None:
    val, label = _fallback_period_from(_TEXT_KCR)
    assert val == "14 May 2024"
    assert label == "_COVER_BLOCK_PERIOD_RE"


def test_fallback_period_to_uses_same_match_as_period_from() -> None:
    val, label = _fallback_period_to(_TEXT_NEW_INVOICE)
    assert val == "30 June 2024"
    assert label == "_BILLING_PERIOD_RE"

    val, label = _fallback_period_to(_TEXT_COVER_BLOCK)
    assert val == "30 June 2024"
    assert label == "_COVER_BLOCK_PERIOD_RE"


def test_fallback_period_returns_none_when_no_match() -> None:
    val, label = _fallback_period_from(_TEXT_NO_HINTS)
    assert val is None
    assert label == ""


def test_fallback_amount_picks_period_charge_re_first() -> None:
    val, label = _fallback_amount(_TEXT_NEW_INVOICE)
    assert val == 1347.96
    # Period charge is the strongest source for £ in a new-format invoice.
    assert label == "_PERIOD_CHARGE_RE"


def test_fallback_amount_picks_pound_amount_for_other_text() -> None:
    # Cover-block text has no Period Charge pattern — falls through.
    val, label = _fallback_amount(_TEXT_COVER_BLOCK)
    assert val == 1347.96
    assert label == "_POUND_AMOUNT_FALLBACK_RE"


def test_fallback_amount_picks_first_amount_for_kcr_credit() -> None:
    val, label = _fallback_amount(_TEXT_KCR)
    assert val == 450.75
    assert label == "_POUND_AMOUNT_FALLBACK_RE"


def test_fallback_amount_does_not_truncate_large_uncommaed_amount() -> None:
    # Regression: £1234.56 (4 digits, no comma) was previously truncated
    # to 123.0 because the fallback regex only matched 1-3 digits before
    # the optional comma group. It must parse as the full amount.
    val, label = _fallback_amount("Amount: £1234.56")
    assert val == 1234.56
    assert label == "_POUND_AMOUNT_FALLBACK_RE"


def test_fallback_amount_still_handles_commaed_amounts() -> None:
    # Comma-grouped amounts must keep working after the \d{1,3} -> \d+
    # widening: "£12,345.67" -> 12345.67.
    val, label = _fallback_amount("Amount: £12,345.67")
    assert val == 12345.67
    assert label == "_POUND_AMOUNT_FALLBACK_RE"


def test_fallback_amount_returns_none_when_no_money() -> None:
    val, label = _fallback_amount(_TEXT_NO_HINTS)
    assert val is None
    assert label == ""


def test_regex_constants_are_compiled() -> None:
    import re

    assert isinstance(_COVER_BLOCK_INV_RE, re.Pattern)
    assert isinstance(_COVER_BLOCK_PERIOD_RE, re.Pattern)
    assert isinstance(_FALLBACK_INV_RE, re.Pattern)
    assert isinstance(_FALLBACK_AMOUNT_RE, re.Pattern)
