"""Tests that detector functions are importable from the processors.detection submodule.

All tests are RED at Phase 0 because ``edf_bill_fetcher.processors.detection``
does not yet exist.
"""

from __future__ import annotations


def test_detect_back_billing_importable() -> None:
    from edf_bill_fetcher.processors.detection import detect_back_billing

    assert detect_back_billing is not None


def test_detect_rebilling_importable() -> None:
    from edf_bill_fetcher.processors.detection import detect_rebilling

    assert detect_rebilling is not None


def test_detect_meter_rollover_importable() -> None:
    from edf_bill_fetcher.processors.detection import detect_meter_rollover

    assert detect_meter_rollover is not None


def test_detect_pdf_format_importable() -> None:
    from edf_bill_fetcher.processors.detection import detect_pdf_format

    assert detect_pdf_format is not None


def test_detect_reconciliation_statement_importable() -> None:
    from edf_bill_fetcher.processors.detection import detect_reconciliation_statement

    assert detect_reconciliation_statement is not None
