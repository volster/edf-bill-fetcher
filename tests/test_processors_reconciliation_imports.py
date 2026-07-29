"""Tests that reconciliation functions are importable from the processors.reconciliation submodule.

All tests are RED at Phase 0 because ``edf_bill_fetcher.processors.reconciliation``
does not yet exist.
"""

from __future__ import annotations


def test_write_reconciliation_sheet_importable() -> None:
    from edf_bill_fetcher.processors.reconciliation import write_reconciliation_sheet

    assert write_reconciliation_sheet is not None


def test_detect_reconciliation_statement_importable() -> None:
    from edf_bill_fetcher.processors.reconciliation import detect_reconciliation_statement

    assert detect_reconciliation_statement is not None


def test_extract_reconciliation_statement_rows_importable() -> None:
    from edf_bill_fetcher.processors.reconciliation import extract_reconciliation_statement_rows

    assert extract_reconciliation_statement_rows is not None
