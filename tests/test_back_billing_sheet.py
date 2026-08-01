from __future__ import annotations

import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from edf_bill_fetcher.io.adapters.pdf import legal_context
from edf_bill_fetcher.io.writers.back_billing import write_back_billing_sheet
from edf_bill_fetcher.processors.detection import detect_back_billing


def _sample_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Invoice #": "T-6715690",
                "Date": "09 Aug 2023",
                "Period From": "04 Apr 2022",
                "Period To": "26 Jul 2023",
                "Amount (£)": 4401.07,
                "Cancel/Rebill Admitted": True,
                "Attachment Name": "671078701920_060264189544_20230809.pdf",
            },
            {
                "Invoice #": "REG-0001",
                "Date": "01 Jan 2024",
                "Period From": "01 Dec 2023",
                "Period To": "31 Dec 2023",
                "Amount (£)": 100.00,
                "Cancel/Rebill Admitted": False,
                "Attachment Name": "reg.pdf",
            },
        ]
    )


def _open_ws(title: str = "Back-billing Analysis") -> Worksheet:
    wb = Workbook()
    ws = wb.active
    ws.title = title
    return ws


def test_write_back_billing_sheet_renders_legal_context_banner() -> None:
    ws = _open_ws()
    bb = detect_back_billing(_sample_df())
    write_back_billing_sheet(ws, bb, account="1234567890")
    # Row 1: title banner with account
    a1 = ws.cell(row=1, column=1).value
    assert isinstance(a1, str)
    assert "BACK-BILLING" in a1.upper()
    assert "1234567890" in a1
    # Row 2: 'LEGAL CONTEXT' label
    a2 = ws.cell(row=2, column=1).value
    assert isinstance(a2, str)
    assert "LEGAL CONTEXT" in a2.upper()
    # Row 3 contains the legal_context() body text
    a3 = ws.cell(row=3, column=1).value
    assert isinstance(a3, str)
    assert legal_context().splitlines()[0] in a3


def test_write_back_billing_sheet_writes_table_headers() -> None:
    ws = _open_ws()
    bb = detect_back_billing(_sample_df())
    write_back_billing_sheet(ws, bb, account="A1")
    # Per spec, row 7 = table header row.
    headers = [ws.cell(row=7, column=c).value for c in range(1, 11)]
    expected = [
        "Invoice #",
        "Bill Date",
        "Period From",
        "Period To",
        "Days Billed",
        "Net Charge (£)",
        "12-Month Limit (days)",
        "Excess Days",
        "Cancel/Rebill Disclosed",
        "Reason Assessment",
    ]
    assert headers == expected


def test_write_back_billing_sheet_one_row_per_backbilled_invoice() -> None:
    ws = _open_ws()
    bb = detect_back_billing(_sample_df())
    write_back_billing_sheet(ws, bb, account="A1")
    # Spec: rows 8+ are data rows. Sample has exactly 1 back-billed invoice.
    a8 = ws.cell(row=8, column=1).value
    assert a8 == "T-6715690"
    # Row 9 carries the TOTAL RETRO... footer (since sample has 1 back-bill).
    a9 = ws.cell(row=9, column=1).value
    assert isinstance(a9, str)
    assert "TOTAL" in a9.upper()
    # No data rows beyond the totals row.
    assert ws.cell(row=10, column=1).value in (None, "")


def test_write_back_billing_sheet_total_charges_footer() -> None:
    ws = _open_ws()
    bb = detect_back_billing(_sample_df())
    write_back_billing_sheet(ws, bb, account="A1")
    # Trailing row somewhere below row 8 carries the totals label and value.
    found = False
    for r in range(9, 15):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and "TOTAL" in v.upper() and "RETRO" in v.upper():
            # The same row's col 5 (or thereabouts) carries the sum.
            sum_cell = ws.cell(row=r, column=6).value
            assert sum_cell == 4401.07
            found = True
            break
    assert found, "TOTAL RETROSPECTIVE CHARGES footer row missing"


def test_write_back_billing_sheet_empty_df_still_renders_header_and_legal_context() -> None:
    ws = _open_ws()
    empty = pd.DataFrame(
        columns=[
            "Invoice #",
            "Bill Date",
            "Period From",
            "Period To",
            "Days Billed",
            "Net Charge (£)",
            "12-Month Limit (days)",
            "Excess Days",
            "Cancel/Rebill Admitted",
            "Reason Assessment",
        ]
    )
    write_back_billing_sheet(ws, empty, account="A1")
    # Legal context still rendered.
    a3 = ws.cell(row=3, column=1).value
    assert isinstance(a3, str)
    assert "back-billing" in a3.lower()
    # Table headers still rendered.
    headers = [ws.cell(row=7, column=c).value for c in range(1, 11)]
    assert headers[0] == "Invoice #"
    # No data rows.
    assert ws.cell(row=8, column=1).value in (None, "")


def test_write_back_billing_sheet_admitted_cell_value_uses_phrase_label() -> None:
    ws = _open_ws()
    bb = detect_back_billing(_sample_df())
    write_back_billing_sheet(ws, bb, account="A1")
    # Admit column (col 9) on row 8 must say 'Admitted phrase' for our
    # sample (the cover-page admit fired).
    v = ws.cell(row=8, column=9).value
    assert v == "Admitted phrase"
