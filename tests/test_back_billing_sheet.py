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
                "Period To": "26 Jul 2022",
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
    # Per spec, row 7 = table header row. 16 columns after Task 4.
    headers = [ws.cell(row=7, column=c).value for c in range(1, 17)]
    expected = [
        "Invoice #",
        "Bill Date",
        "Period From",
        "Period To",
        "Days Billed",
        "Period Charge (£)",
        "Value Source",
        "12-Month Limit (days)",
        "Excess Days",
        "Cancel/Rebill Disclosed",
        "Reason Assessment",
        "Open PDF",
        "View on Evidence Report",
        "Status",
        "Superseded By",
        "Partial Overlap",
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
            "Period Charge (£)",
            "Value Source",
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
    # Table headers still rendered (16 columns after Task 4).
    headers = [ws.cell(row=7, column=c).value for c in range(1, 17)]
    assert headers[0] == "Invoice #"
    assert "Status" in headers
    # No data rows.
    assert ws.cell(row=8, column=1).value in (None, "")


def test_write_back_billing_sheet_admitted_cell_value_uses_phrase_label() -> None:
    ws = _open_ws()
    bb = detect_back_billing(_sample_df())
    write_back_billing_sheet(ws, bb, account="A1")
    # Admit column (col 10) on row 8 must say 'Admitted phrase' for our
    # sample (the cover-page admit fired).
    v = ws.cell(row=8, column=10).value
    assert v == "Admitted phrase"


def _two_row_bb() -> pd.DataFrame:
    """Two back-billing rows with synthetic invoice IDs A and B."""
    return pd.DataFrame(
        [
            {
                "Invoice #": "A",
                "Bill Date": "2021-06-01",
                "Period From": "2020-01-01",
                "Period To": "2021-06-01",
                "Days Billed": 517,
                "Period Charge (£)": 500.0,
                "Value Source": "Period Charge",
                "12-Month Limit (days)": 365,
                "Excess Days": 152,
                "Cancel/Rebill Admitted": False,
                "Reason Assessment": "test",
            },
            {
                "Invoice #": "B",
                "Bill Date": "2021-12-01",
                "Period From": "2020-06-01",
                "Period To": "2021-12-01",
                "Days Billed": 549,
                "Period Charge (£)": 300.0,
                "Value Source": "Period Charge",
                "12-Month Limit (days)": 365,
                "Excess Days": 184,
                "Cancel/Rebill Admitted": False,
                "Reason Assessment": "test",
            },
        ]
    )


def test_write_back_billing_sheet_status_columns() -> None:
    ws = _open_ws()
    bb = _two_row_bb()
    domination_map: dict[str, tuple[str, bool]] = {"B": ("A", False)}
    write_back_billing_sheet(ws, bb, domination_map=domination_map)

    # Row 7 is the header row.
    header_row = [cell.value for cell in ws[7]]
    assert "Status" in header_row
    assert "Superseded By" in header_row
    assert "Partial Overlap" in header_row
    assert "Value Source" in header_row

    status_col = header_row.index("Status") + 1
    superseded_by_col = header_row.index("Superseded By") + 1
    partial_overlap_col = header_row.index("Partial Overlap") + 1
    inv_col = header_row.index("Invoice #") + 1

    for row_idx in range(8, ws.max_row + 1):
        inv_num = ws.cell(row=row_idx, column=inv_col).value
        if inv_num == "A":
            assert ws.cell(row=row_idx, column=status_col).value == "Live"
            assert ws.cell(row=row_idx, column=superseded_by_col).value in (None, "")
            assert ws.cell(row=row_idx, column=partial_overlap_col).value in (None, "")
        elif inv_num == "B":
            assert ws.cell(row=row_idx, column=status_col).value == "Superseded"
            assert ws.cell(row=row_idx, column=superseded_by_col).value == "A"
            assert ws.cell(row=row_idx, column=partial_overlap_col).value in ("", None)
            # Superseded rows are outline-collapsed sub-rows.
            assert ws.row_dimensions[row_idx].outline_level == 1
            assert ws.row_dimensions[row_idx].hidden is True


def test_write_back_billing_sheet_status_partial_overlap_yes() -> None:
    """When domination_map carries partial_overlap=True, the cell reads 'Yes'."""
    ws = _open_ws()
    bb = _two_row_bb()
    domination_map: dict[str, tuple[str, bool]] = {"B": ("A", True)}
    write_back_billing_sheet(ws, bb, domination_map=domination_map)

    header_row = [cell.value for cell in ws[7]]
    partial_overlap_col = header_row.index("Partial Overlap") + 1
    inv_col = header_row.index("Invoice #") + 1
    for row_idx in range(8, ws.max_row + 1):
        if ws.cell(row=row_idx, column=inv_col).value == "B":
            assert ws.cell(row=row_idx, column=partial_overlap_col).value == "Yes"
            return
    raise AssertionError("Row for invoice B not found")


def test_write_back_billing_sheet_total_excludes_superseded() -> None:
    ws = _open_ws()
    bb = _two_row_bb()
    domination_map: dict[str, tuple[str, bool]] = {"B": ("A", False)}
    write_back_billing_sheet(ws, bb, domination_map=domination_map)

    header_row = [cell.value for cell in ws[7]]
    period_charge_col = header_row.index("Period Charge (£)") + 1

    total_row_idx = None
    for row_idx in range(8, ws.max_row + 1):
        v = ws.cell(row=row_idx, column=1).value
        if v and "TOTAL RETROSPECTIVE" in str(v):
            total_row_idx = row_idx
            break
    assert total_row_idx is not None
    total_value = ws.cell(row=total_row_idx, column=period_charge_col).value
    assert total_value == 500.0  # only A (Live), not B (Superseded)


def test_write_back_billing_sheet_no_domination_map_all_live() -> None:
    """Without a domination_map, every row is Live and the total sums all rows."""
    ws = _open_ws()
    bb = _two_row_bb()
    write_back_billing_sheet(ws, bb)

    header_row = [cell.value for cell in ws[7]]
    status_col = header_row.index("Status") + 1
    inv_col = header_row.index("Invoice #") + 1
    for row_idx in range(8, ws.max_row + 1):
        inv_num = ws.cell(row=row_idx, column=inv_col).value
        if inv_num in ("A", "B"):
            assert ws.cell(row=row_idx, column=status_col).value == "Live"
            assert ws.row_dimensions[row_idx].outline_level == 0
