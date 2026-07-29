"""Tests that SAP parser functions are importable from the processors.sap_parsers submodule.

All tests are RED at Phase 0 because ``edf_bill_fetcher.processors.sap_parsers``
does not yet exist.
"""

from __future__ import annotations


def test_detect_sap_dump_importable() -> None:
    from edf_bill_fetcher.processors.sap_parsers import detect_sap_dump

    assert detect_sap_dump is not None


def test_parse_sap_contract_history_importable() -> None:
    from edf_bill_fetcher.processors.sap_parsers import parse_sap_contract_history

    assert parse_sap_contract_history is not None


def test_parse_sap_financial_transactions_importable() -> None:
    from edf_bill_fetcher.processors.sap_parsers import parse_sap_financial_transactions

    assert parse_sap_financial_transactions is not None


def test_parse_sap_meter_read_history_importable() -> None:
    from edf_bill_fetcher.processors.sap_parsers import parse_sap_meter_read_history

    assert parse_sap_meter_read_history is not None


def test_extract_new_credit_fields_importable() -> None:
    from edf_bill_fetcher.processors.sap_parsers import extract_new_credit_fields

    assert extract_new_credit_fields is not None


def test_extract_new_invoice_fields_importable() -> None:
    from edf_bill_fetcher.processors.sap_parsers import extract_new_invoice_fields

    assert extract_new_invoice_fields is not None


def test_extract_reconciliation_statement_rows_importable() -> None:
    from edf_bill_fetcher.processors.sap_parsers import extract_reconciliation_statement_rows

    assert extract_reconciliation_statement_rows is not None


def test_detect_reconciliation_statement_importable() -> None:
    from edf_bill_fetcher.processors.sap_parsers import detect_reconciliation_statement

    assert detect_reconciliation_statement is not None
