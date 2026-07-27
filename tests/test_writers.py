"""Tests for edf_bill_fetcher.writers submodule.

Verifies that the writer functions are importable from the writers
package and behave correctly.
"""

from __future__ import annotations


def test_writers_submodule_importable():
    from edf_bill_fetcher import writers

    assert writers is not None


def test_write_evidence_sheet_importable():
    from edf_bill_fetcher.writers import write_evidence_sheet

    assert write_evidence_sheet is not None


def test_write_summary_sheet_importable():
    from edf_bill_fetcher.writers import write_summary_sheet

    assert write_summary_sheet is not None


def test_write_reconciliation_sheet_importable():
    from edf_bill_fetcher.writers import write_reconciliation_sheet

    assert write_reconciliation_sheet is not None


def test_write_back_billing_sheet_importable():
    from edf_bill_fetcher.writers import write_back_billing_sheet

    assert write_back_billing_sheet is not None


def test_write_rebilling_sheet_importable():
    from edf_bill_fetcher.writers import write_rebilling_sheet

    assert write_rebilling_sheet is not None


def test_write_meter_readings_sheet_importable():
    from edf_bill_fetcher.writers import write_meter_readings_sheet

    assert write_meter_readings_sheet is not None


def test_write_contract_history_sheet_importable():
    from edf_bill_fetcher.writers import write_contract_history_sheet

    assert write_contract_history_sheet is not None


def test_write_sap_contract_history_sheet_importable():
    from edf_bill_fetcher.writers import write_sap_contract_history_sheet

    assert write_sap_contract_history_sheet is not None


def test_write_sap_bb_matches_sheet_importable():
    from edf_bill_fetcher.writers import _write_sap_bb_matches_sheet

    assert _write_sap_bb_matches_sheet is not None


def test_export_to_excel_importable():
    from edf_bill_fetcher.writers import export_to_excel

    assert export_to_excel is not None


def test_writers_re_exported_from_edf_collector():
    from edf_bill_fetcher.writers import write_reconciliation_sheet as WRS_writers
    from edf_collector import write_reconciliation_sheet as WRS_collector

    assert WRS_collector is WRS_writers
