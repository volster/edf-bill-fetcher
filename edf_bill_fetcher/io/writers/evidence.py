"""Evidence sheet + summary sheet writers for the EDF workbook.

Extracted from ``edf_bill_fetcher/writers/__init__.py`` (Phase 5A) during the
modularization refactor.  These two functions are the heart of the Evidence tab
and Summary tab.
"""

from __future__ import annotations

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Re-export shared helpers / theme constants used by these two functions.
from edf_bill_fetcher.helpers.date_utils import to_excel_date
from edf_bill_fetcher.helpers.excel_utils import (
    hcell as _hcell,
)
from edf_bill_fetcher.helpers.excel_utils import (
    literal_cell,
)
from edf_bill_fetcher.helpers.theme import CELL_BORDER
from edf_bill_fetcher.io.writers.sheet_layout import (
    freeze_at,
    write_header_row,
)
from edf_bill_fetcher.processors.preparation import normalize_frame, number, scope

# Column letter map for the evidence sheet (matches ``EVIDENCE_HEADERS`` below):
# A=Source B=Sender C=Date D=PeriodFrom E=PeriodTo F=Invoice#
# G=Amount H=PeriodCharge I=UnitRate J=%Change K=EntryType
# L=Reading M=Units N=StandingChg O=Tariff P=AttachmentName
# Q=Details R=LogicUsed S=AnomalyFlag
# (Duplicate-of-link cells are rendered in a post-loop pass for
# the ``is_duplicate=True`` branch and don't appear in this
# header list.)
#
# F1 (SEV-1):  every COL_* is derived from the headers list, not
# hard-coded.  Inserting a new column at any position requires
# updating exactly one place (the headers list) — the conditional
# formatting range, formula references, column widths and the
# dedup hyperlink pass all read the same index.  Verified by
# ``tests/test_evidence_sheet_columns.py``.
EVIDENCE_HEADERS: list[str] = [
    "Source",
    "Sender",
    "Date",
    "Period From",
    "Period To",
    "Invoice #",
    "Amount (£)",
    "Period Charge (£)",
    "Unit Rate (p/kWh)",
    "% Change",
    "Entry Type",
    "Reading",
    "Units (kWh)",
    "Standing Chg (p/day)",
    # Tariff price-plan name (e.g. "Freedom", "Standard");
    # extracted by ``extract_new_invoice_fields`` on KI-style
    # bills. See ``_process_new_invoice``.
    "Tariff",
    "Attachment Name",
    "Details",
    "Logic Used",
    "Anomaly Flag",
    "Sub Periods",
    "Balance (£)",
    "Transaction Amount (£)",
    "Amount Kind",
    "Account",
    "Meter",
    "Register",
    "Fuel",
    "Register Reading (kWh)",
    "Cancel/Rebill Admitted",
    "Source ID",
    "Source IDs",
    "Content Hash",
    "Original Filename",
    "Page From",
    "Page To",
    "Text Extraction",
    "OCR Pages",
    "OCR Minimum Confidence (%)",
    "OCR Review Required",
    "Document Role",
    "Document Date",
    "First Billed Date",
    "Billing Date Status",
    "Original Invoice #",
    "Source Row",
    "Gross Electricity Charges (£)",
    "Reading Date",
    "Closing Reading Type",
    "Statement Reconciliation Difference (£)",
    "Statement Reconciliation Status",
    "Reconciliation Review",
    "Source Representation",
]
COL_AMOUNT = EVIDENCE_HEADERS.index("Amount (£)") + 1

# --- function body 1: write_evidence_sheet (was writers/__init__.py L182-422) ---


def write_evidence_sheet(ws, df, is_duplicate=False, saved_paths=None):
    """Render canonical data and link only to files actually saved in this run."""
    headers = EVIDENCE_HEADERS + (
        ["Duplicate Of"] if is_duplicate else ["Reason"] if "Reason" in df else []
    )
    write_header_row(ws, 1, headers, bg="888888" if is_duplicate else "FE5716", height=28)
    saved_paths = saved_paths or {}
    previous: dict[tuple[str, ...], tuple[int, float]] = {}
    for row_number, record in enumerate(normalize_frame(df).to_dict("records"), 2):
        record = dict(record)
        key = scope(record)
        balance = number(record.get("Balance (£)"))
        dated = to_excel_date(record.get("Date")) is not None
        prior = previous.get(key)
        record["% Change"] = None
        record["Anomaly Flag"] = ""
        if balance is not None and dated:
            previous[key] = (row_number, balance)
        for col, header in enumerate(headers, 1):
            value = record.get(header)
            if header in {
                "Date",
                "Period From",
                "Period To",
                "Document Date",
                "First Billed Date",
                "Reading Date",
            }:
                value = to_excel_date(value)
            cell = literal_cell(ws, row_number, col, value)
            cell.font = Font(name="Calibri", size=10)
            cell.border = CELL_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if is_duplicate:
                cell.fill = PatternFill("solid", start_color="E0E0E0")
            elif row_number % 2 == 0:
                cell.fill = PatternFill("solid", start_color="FFF3EE")
            if "(£)" in header:
                cell.number_format = "£#,##0.00"
            elif header == "% Change":
                cell.number_format = "0.0%"
            elif header in {"Date", "Period From", "Period To"}:
                cell.number_format = "dd/mm/yyyy"
            if not is_duplicate and dated and balance is not None and prior is not None:
                balance_column = get_column_letter(headers.index("Balance (£)") + 1)
                prior_row, prior_balance = prior
                if header == "% Change" and prior_balance:
                    cell.value = f'=IFERROR(({balance_column}{row_number}-{balance_column}{prior_row})/ABS({balance_column}{prior_row}),"")'
                    cell.data_type = "f"
                elif header == "Anomaly Flag" and prior_balance > 0:
                    cell.value = f'=IF(AND({balance_column}{prior_row}>0,{balance_column}{row_number}>{balance_column}{prior_row}*2),"⚠ >100% INCREASE","")'
                    cell.data_type = "f"
            if header == "Attachment Name":
                source_id = str(record.get("Source ID", "") or "")
                path = saved_paths.get(source_id) if source_id else saved_paths.get(str(value))
                if path:
                    cell.hyperlink = str(path)
                    cell.font = Font(name="Calibri", size=10, color="0000FF", underline="single")
            if header == "Duplicate Of":
                pos = number(record.get("_matches_kept_idx"))
                if pos is not None and pos >= 0:
                    cell.hyperlink = openpyxl.worksheet.hyperlink.Hyperlink(
                        ref=cell.coordinate, location=f"'EDF Evidence Report'!A{int(pos) + 2}"
                    )
    for col, header in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(col)].width = (
            36 if header in {"Details", "Source IDs", "Source ID", "Content Hash"} else 20
        )
    if not is_duplicate and len(df):
        anomaly_column = get_column_letter(headers.index("Anomaly Flag") + 1)
        ws.conditional_formatting.add(
            f"{anomaly_column}2:{anomaly_column}{len(df) + 1}",
            FormulaRule(
                formula=[f'${anomaly_column}2<>""'], fill=PatternFill("solid", start_color="FFCCCC")
            ),
        )
    ws.auto_filter.ref = ws.dimensions
    freeze_at(ws, "A2")


# ---------------------------------------------------------------------------
# Write summary sheet — uses _xlfn.MAXIFS/_xlfn.MINIFS so Excel evaluates
# on load without the dynamic-array compatibility dialog
# ---------------------------------------------------------------------------

# --- function body 2: write_summary_sheet (was writers/__init__.py L431-519) ---


def write_summary_sheet(ws, years, evidence_sheet_name, last_data_row=5000):
    """Render the Annual Summary worksheet with per-year balance aggregates."""
    ws.title = "Annual Summary"

    headers = [
        "Year",
        "Balance Range (£)",
        "Records",
        "Avg Balance (£)",
        "Peak Balance (£)",
        "Lowest Balance (£)",
        "Drill down",
    ]
    for col, h in enumerate(headers, 1):
        _hcell(ws, 1, col, h, bg="10367A")
    ws.row_dimensions[1].height = 28

    alt_fill = PatternFill("solid", start_color="EEF2FF")
    esn = evidence_sheet_name

    date_col = f"'{esn}'!$C$2:$C${last_data_row}"
    # Amount lives in col G of the evidence sheet (not col E, which is
    # Period To); derive it from the shared COL_AMOUNT constant.
    amt_col_letter = get_column_letter(EVIDENCE_HEADERS.index("Balance (£)") + 1)
    amt_col = f"'{esn}'!${amt_col_letter}$2:${amt_col_letter}${last_data_row}"

    for r_idx, year_val in enumerate(years, 2):
        row_fill = alt_fill if r_idx % 2 == 0 else PatternFill()
        yr_cell = f"A{r_idx}"

        # _xlfn. prefix tells Excel to evaluate MAXIFS/MINIFS on load without
        # the dynamic-array compatibility dialog.
        peak_f = f'=IFERROR(_xlfn.MAXIFS({amt_col},{amt_col},"<>",{date_col},">="&DATE({yr_cell},1,1),{date_col},"<"&DATE({yr_cell}+1,1,1)),"")'
        low_f = f'=IFERROR(_xlfn.MINIFS({amt_col},{amt_col},"<>",{date_col},">="&DATE({yr_cell},1,1),{date_col},"<"&DATE({yr_cell}+1,1,1)),"")'
        range_f = f'=IFERROR(_xlfn.MAXIFS({amt_col},{amt_col},"<>",{date_col},">="&DATE({yr_cell},1,1),{date_col},"<"&DATE({yr_cell}+1,1,1))-_xlfn.MINIFS({amt_col},{amt_col},"<>",{date_col},">="&DATE({yr_cell},1,1),{date_col},"<"&DATE({yr_cell}+1,1,1)),"")'

        row_values = [
            int(year_val),
            f'=IF(C{r_idx}=0,"",{range_f[1:]})',
            f'=COUNTIFS({amt_col},"<>",{date_col},">="&DATE({yr_cell},1,1),{date_col},"<"&DATE({yr_cell}+1,1,1))',
            f'=IFERROR(AVERAGEIFS({amt_col},{date_col},">="&DATE({yr_cell},1,1),{date_col},"<"&DATE({yr_cell}+1,1,1)),"")',
            f'=IF(C{r_idx}=0,"",{peak_f[1:]})',
            f'=IF(C{r_idx}=0,"",{low_f[1:]})',
            "→ Drill down",
        ]
        for c_idx, val in enumerate(row_values, 1):
            c = ws.cell(row=r_idx, column=c_idx, value=val)
            c.font = Font(name="Calibri", size=10)
            c.fill = row_fill
            c.border = CELL_BORDER
            c.alignment = Alignment(
                horizontal="center" if c_idx == 1 else "right",
                vertical="top",
            )
            if c_idx == 2:
                c.number_format = "£#,##0.00"
            elif c_idx == 3:
                c.number_format = "#,##0"
            elif c_idx > 3:
                c.number_format = "£#,##0.00"
            if c_idx == 7:
                c.hyperlink = openpyxl.worksheet.hyperlink.Hyperlink(
                    ref=c.coordinate,
                    location="'EDF Evidence Report'!A2",
                    display="→ Drill down",
                    tooltip="Jump to Reconciliation Drill-down sheet",
                )

    # Grand total row — SUM/MAX/MIN over the year rows only, no dynamic-array functions
    n = len(years) + 2
    first_r = 2
    last_r = n - 1
    tot_fill = PatternFill("solid", start_color="10367A")
    tot_specs = [
        ("OVERALL", None, "center"),
        (
            f'=IF(SUM(C{first_r}:C{last_r})=0,"",IFERROR(MAX(E{first_r}:E{last_r})-MIN(F{first_r}:F{last_r}),""))',
            "£#,##0.00",
            "right",
        ),
        (f"=SUM(C{first_r}:C{last_r})", "#,##0", "right"),
        (
            f'=IFERROR(SUMPRODUCT(D{first_r}:D{last_r},C{first_r}:C{last_r})/SUM(C{first_r}:C{last_r}),"")',
            "£#,##0.00",
            "right",
        ),
        (
            f'=IF(SUM(C{first_r}:C{last_r})=0,"",IFERROR(MAX(E{first_r}:E{last_r}),""))',
            "£#,##0.00",
            "right",
        ),
        (
            f'=IF(SUM(C{first_r}:C{last_r})=0,"",IFERROR(MIN(F{first_r}:F{last_r}),""))',
            "£#,##0.00",
            "right",
        ),
    ]
    for c_idx, (val, num_fmt, align) in enumerate(tot_specs, 1):
        c = ws.cell(row=n, column=c_idx, value=val)
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = tot_fill
        c.border = CELL_BORDER
        c.alignment = Alignment(horizontal=align)
        if num_fmt:
            c.number_format = num_fmt

    for col_letter in ["A", "B", "C", "D", "E", "F"]:
        ws.column_dimensions[col_letter].width = 22
    ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# Main export function
# ---------------------------------------------------------------------------
__all__ = ["write_evidence_sheet", "write_summary_sheet"]
