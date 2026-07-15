from __future__ import annotations

import os

import pytest
from openpyxl import load_workbook

from edf_collector import export_to_excel


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
