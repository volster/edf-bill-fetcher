"""Tests that regex pattern constants are importable from the processors.patterns submodule.

All tests are RED at Phase 0 because ``edf_bill_fetcher.processors.patterns``
does not yet exist.
"""

from __future__ import annotations


def test_amount_patterns_importable() -> None:
    from edf_bill_fetcher.processors.patterns import AMOUNT_PATTERNS

    assert AMOUNT_PATTERNS is not None


def test_reading_patterns_importable() -> None:
    from edf_bill_fetcher.processors.patterns import READING_PATTERNS

    assert READING_PATTERNS is not None


def test_amount_pattern_new_bill_importable() -> None:
    from edf_bill_fetcher.processors.patterns import _AMOUNT_PATTERN_NEW_BILL

    assert _AMOUNT_PATTERN_NEW_BILL is not None


def test_amount_pattern_ongoing_balance_importable() -> None:
    from edf_bill_fetcher.processors.patterns import _AMOUNT_PATTERN_ONGOING_BALANCE

    assert _AMOUNT_PATTERN_ONGOING_BALANCE is not None


def test_cover_block_inv_re_importable() -> None:
    from edf_bill_fetcher.processors.patterns import _COVER_BLOCK_INV_RE

    assert _COVER_BLOCK_INV_RE is not None


def test_cover_block_period_re_importable() -> None:
    from edf_bill_fetcher.processors.patterns import _COVER_BLOCK_PERIOD_RE

    assert _COVER_BLOCK_PERIOD_RE is not None


def test_fallback_amount_re_importable() -> None:
    from edf_bill_fetcher.processors.patterns import _FALLBACK_AMOUNT_RE

    assert _FALLBACK_AMOUNT_RE is not None


def test_fallback_inv_re_importable() -> None:
    from edf_bill_fetcher.processors.patterns import _FALLBACK_INV_RE

    assert _FALLBACK_INV_RE is not None


def test_period_re_importable() -> None:
    from edf_bill_fetcher.processors.patterns import PERIOD_RE

    assert PERIOD_RE is not None
