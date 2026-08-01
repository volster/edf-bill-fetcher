"""Tests that reconciliation functions are importable from their canonical homes.

- ``write_reconciliation_sheet`` lives in ``edf_bill_fetcher.io.writers.reconciliation``
- ``detect_reconciliation_statement`` and ``extract_reconciliation_statement_rows``
  live in ``edf_bill_fetcher.processors.detection``
"""

from __future__ import annotations


def test_write_reconciliation_sheet_importable() -> None:
    from edf_bill_fetcher.io.writers.reconciliation import write_reconciliation_sheet

    assert write_reconciliation_sheet is not None


def test_detect_reconciliation_statement_importable() -> None:
    from edf_bill_fetcher.processors.detection import detect_reconciliation_statement

    assert detect_reconciliation_statement is not None


def test_extract_reconciliation_statement_rows_importable() -> None:
    from edf_bill_fetcher.processors.detection import extract_reconciliation_statement_rows

    assert extract_reconciliation_statement_rows is not None
