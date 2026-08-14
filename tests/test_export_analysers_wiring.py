from __future__ import annotations

import os

import pytest
from openpyxl import load_workbook

from edf_bill_fetcher.io.writers import export_to_excel


def _sample_data() -> list[dict]:
    return [
        {
            "Source": "Local PDF Folder",
            "Sender": "edf.co.uk",
            "Date": "01 Sep 2023",
            "Period From": "01 Jan 2022",
            "Period To": "31 Aug 2023",
            "Invoice #": "T-X1",
            "Amount (£)": 1000.0,
            "Period Charge (£)": 800.0,
            "Unit Rate (p/kWh)": 25.0,
            "% Change": None,
            "Entry Type": "New Bill",
            "Reading": "Actual",
            "Units (kWh)": 300.0,
            "Standing Chg (p/day)": 50.0,
            "Tariff": "Standard",
            "Attachment Name": "T-X1.pdf",
            "Details": "Reading was actual",
            "Logic Used": "PDF new-format",
            "Anomaly Flag": "",
            "Cancel/Rebill Admitted": True,
        },
        {
            "Source": "Local PDF Folder",
            "Sender": "edf.co.uk",
            "Date": "01 Oct 2023",
            "Period From": "01 Feb 2022",
            "Period To": "30 Sep 2023",
            "Invoice #": "T-X2",
            "Amount (£)": 1500.0,
            "Period Charge (£)": 1200.0,
            "Unit Rate (p/kWh)": 25.0,
            "% Change": None,
            "Entry Type": "New Bill",
            "Reading": "Actual",
            "Units (kWh)": 400.0,
            "Standing Chg (p/day)": 50.0,
            "Tariff": "Standard",
            "Attachment Name": "T-X2.pdf",
            "Details": "Reading was actual",
            "Logic Used": "PDF new-format",
            "Anomaly Flag": "",
            "Cancel/Rebill Admitted": False,
        },
    ]


@pytest.fixture
def tmp_xlsx(tmp_path):
    return str(tmp_path / "test_run.xlsx")


def test_export_to_excel_emits_four_new_analysis_tabs(tmp_xlsx: str) -> None:
    export_to_excel(
        _sample_data(),
        tmp_xlsx,
        error_log=[],
        config={"use_dedup": False, "acc_num": "0123456789"},
    )
    assert os.path.exists(tmp_xlsx)
    wb = load_workbook(tmp_xlsx, read_only=True)
    names = set(wb.sheetnames)
    # The four new tabs must exist alongside the existing writers.
    assert "Back-billing Analysis" in names
    assert "Rebilling & Corrections" in names
    assert "Meter Readings" in names
    assert "Contract History" in names
    wb.close()


def test_export_writes_superseded_reconciliation_sheet(tmp_xlsx: str) -> None:
    export_to_excel(
        _sample_data(),
        tmp_xlsx,
        error_log=[],
        config={"use_dedup": False, "acc_num": "0123456789"},
    )
    assert os.path.exists(tmp_xlsx)
    wb = load_workbook(tmp_xlsx, read_only=True)
    names = wb.sheetnames
    assert "Superseded Reconciliation" in names
    wb.close()


def test_evidence_index_uses_full_df_rows_after_middle_row_filtered(tmp_xlsx: str) -> None:
    """Analyser hotlinks must target rows on the FULL evidence sheet, not the
    filtered analysis frame.

    INV-001 (amount 100) is below ``analysis_min`` so ``_prepare_analysis_frame``
    drops it from ``dfc``.  The EDF Evidence Report sheet still carries it at
    row 2, so the back-billing hotlink for INV-002 (next on the full frame,
    row 3) must point at ``!A3`` — building the index on ``dfc`` would wrongly
    map INV-002 to row 2 (INV-001's slot).
    """
    rows = [
        {
            "Source": "Local PDF Folder",
            "Sender": "edf.co.uk",
            "Date": "01 Jan 2023",
            "Period From": "01 Jan 2022",
            "Period To": "31 Dec 2022",
            "Invoice #": "INV-001",
            "Amount (£)": 100.0,
            "Period Charge (£)": 100.0,
            "Unit Rate (p/kWh)": 25.0,
            "% Change": None,
            "Entry Type": "New Bill",
            "Reading": "Actual",
            "Units (kWh)": 300.0,
            "Standing Chg (p/day)": 50.0,
            "Tariff": "Standard",
            "Attachment Name": "INV-001.pdf",
            "Details": "below analysis_min",
            "Logic Used": "PDF new-format",
            "Anomaly Flag": "",
        },
        {
            "Source": "Local PDF Folder",
            "Sender": "edf.co.uk",
            "Date": "01 Feb 2023",
            "Period From": "01 Jan 2022",
            "Period To": "31 Jan 2023",
            "Invoice #": "INV-002",
            "Amount (£)": 1000.0,
            "Period Charge (£)": 800.0,
            "Unit Rate (p/kWh)": 25.0,
            "% Change": None,
            "Entry Type": "New Bill",
            "Reading": "Actual",
            "Units (kWh)": 400.0,
            "Standing Chg (p/day)": 50.0,
            "Tariff": "Standard",
            "Attachment Name": "INV-002.pdf",
            "Details": "back-billed",
            "Logic Used": "PDF new-format",
            "Anomaly Flag": "",
        },
        {
            "Source": "Local PDF Folder",
            "Sender": "edf.co.uk",
            "Date": "01 Apr 2023",
            "Period From": "01 Mar 2023",
            "Period To": "31 Mar 2023",
            "Invoice #": "INV-003",
            "Amount (£)": 900.0,
            "Period Charge (£)": 700.0,
            "Unit Rate (p/kWh)": 25.0,
            "% Change": None,
            "Entry Type": "New Bill",
            "Reading": "Actual",
            "Units (kWh)": 300.0,
            "Standing Chg (p/day)": 50.0,
            "Tariff": "Standard",
            "Attachment Name": "INV-003.pdf",
            "Details": "short period",
            "Logic Used": "PDF new-format",
            "Anomaly Flag": "",
        },
    ]
    export_to_excel(
        rows,
        tmp_xlsx,
        error_log=[],
        config={"use_dedup": False, "acc_num": "0123456789", "analysis_min": 500.0},
    )
    wb = load_workbook(tmp_xlsx)
    # Evidence sheet row layout: header row 1, then full-frame rows 2-4.
    ws_ev = wb["EDF Evidence Report"]
    assert ws_ev.cell(row=2, column=6).value == "INV-001"
    assert ws_ev.cell(row=3, column=6).value == "INV-002"
    # Back-billing Analysis hotlink for INV-002 (col 14) must point at its
    # FULL-frame row 3, not the dfc-relative row 2.
    ws_bb = wb["Back-billing Analysis"]
    target = None
    for r in range(8, ws_bb.max_row + 1):
        if ws_bb.cell(row=r, column=1).value == "INV-002":
            target = r
            break
    assert target is not None, "INV-002 not found on Back-billing Analysis sheet"
    cell = ws_bb.cell(row=target, column=14)
    assert cell.hyperlink is not None, "INV-002 hotlink missing"
    assert cell.hyperlink.location.endswith("!A3"), cell.hyperlink.location
    wb.close()


def test_evidence_report_does_not_leak_diagnostic_columns(tmp_xlsx: str) -> None:
    """The ``Source PDF Text`` / ``_regex_trace`` / ``Balance Last Bill (£)``
    columns are diagnostic-only and must NEVER appear on the saved
    ``EDF Evidence Report`` tab.  Pre-fix the col_order reindex never
    dropped them so the visible workbook carried the 4 KB-per-row
    PDF body blocks for every record, polluting the main sheet with
    "process internals" noise that adds ~50% to the saved workbook
    size for nothing visible to the consumer.

    This test feeds the same data through ``export_to_excel`` and
    asserts the saved Evidence Report sheet has exactly the
    canonical 19 header columns (Source … Duplicate Of).
    """
    rows = _sample_data()
    # Plant every diagnostic col on the records so the regression
    # would surface if any of them leaked.
    for row in rows:
        row["Source PDF Text"] = "very long body text " * 200
        row["_regex_trace"] = "trace path"
        row["Balance Last Bill (£)"] = 123.45
    export_to_excel(
        rows,
        tmp_xlsx,
        error_log=[],
        config={"use_dedup": False, "acc_num": "0123456789"},
    )
    wb = load_workbook(tmp_xlsx, read_only=True)
    ws = wb["EDF Evidence Report"]
    headers = []
    for cell in ws[1]:
        if cell.value is not None:
            headers.append(cell.value)
    wb.close()
    assert "Source PDF Text" not in headers, (
        f"diagnostic column 'Source PDF Text' leaked into Evidence Report: {headers}"
    )
    assert "_regex_trace" not in headers, (
        f"diagnostic column '_regex_trace' leaked into Evidence Report: {headers}"
    )
    assert "Balance Last Bill (£)" not in headers, (
        f"diagnostic column 'Balance Last Bill (£)' leaked into Evidence Report: {headers}"
    )
