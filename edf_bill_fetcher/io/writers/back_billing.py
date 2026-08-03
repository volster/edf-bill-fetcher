"""Back-billing analysis writer — extracted from writers/__init__.py.

Contains: detect_back_billing (pure-pandas detector), write_back_billing_sheet
(renders the "Back-billing Analysis" worksheet), and the private
_assess_reason helper that builds the deterministic Reason Assessment
narrative.
"""
from __future__ import annotations

from datetime import datetime

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from edf_bill_fetcher.helpers.date_utils import _safe_to_datetime
from edf_bill_fetcher.helpers.excel_utils import (
    hcell as _hcell,
)
from edf_bill_fetcher.helpers.excel_utils import (
    money as _money,
)
from edf_bill_fetcher.helpers.excel_utils import (
    num as _num,
)
from edf_bill_fetcher.helpers.excel_utils import (
    open_pdf_hyperlink_cell as _open_pdf_hyperlink_cell,
)
from edf_bill_fetcher.helpers.excel_utils import (
    text as _text,
)
from edf_bill_fetcher.helpers.theme import CELL_BORDER
from edf_bill_fetcher.io.adapters.pdf import legal_context
from edf_bill_fetcher.writers._helpers import _disclosed_label

# --- _assess_reason (was writers/__init__.py L2663-2689) ---


def _assess_reason(
    invoice: str,
    days: int,
    admitted: bool,
    period_from: pd.Timestamp,
    period_to: pd.Timestamp,
) -> str:
    """Return a short, deterministic narrative for the Reason Assessment.

    column of the Back-billing sheet. Template-driven (no LLM).
    """
    pf = period_from.strftime("%d %b %Y")
    pt = period_to.strftime("%d %b %Y")
    excess = days - 365
    if admitted:
        head = (
            f"Invoice {invoice} billed {days} days ({pf} to {pt}), "
            f"{excess} days past the 12-month back-billing limit. "
            "EDF's cover page admits a cancellation/reversal, which is "
            "direct evidence the bill is a back-billing remedy."
        )
    else:
        head = (
            f"Invoice {invoice} billed {days} days ({pf} to {pt}), "
            f"{excess} days past the 12-month back-billing limit. No "
            "admit-phrase was found on the cover page."
        )
    return head


# --- detect_back_billing (was writers/__init__.py L2704-2806) ---


def detect_back_billing(df: pd.DataFrame) -> pd.DataFrame:
    """Return invoices whose billing period exceeds 12 months.

    Back-billing (Ofgem / Electricity Act 1989 s.84B) bars suppliers
    from charging a domestic customer for energy supplied more than
    12 months before the bill that first raised the charge. This
    detector surfaces any single invoice whose ``Period From`` ->
    ``Period To`` window exceeds 365 days, alongside whether the
    cover page admits a cancellation/reversal (the
    ``Cancel/Rebill Admitted`` column populated earlier in the
    pipeline by :func:`extract_admit_phrase`).

    The function tolerates a missing ``Cancel/Rebill Admitted``
    column (treated as ``False``).

    Output columns:
        Invoice #, Bill Date, Period From, Period To, Days Billed,
        Net Charge (£), 12-Month Limit (days), Excess Days,
        Cancel/Rebill Admitted, Reason Assessment.

    Rows with unparseable ``Period From``/``Period To`` are skipped
    silently. Output is sorted by ``Bill Date`` and re-indexed.

    Architectural note (SAP cross-feeding):
    This detector takes only the inferred-evidence dataframe. SAP
    data-dump rows (Contract-and-Product-Change-History,
    Meter-Read-History, Financial-Transactions) are surfaced in
    their own tabs (SAP Contract History / SAP Meter Readings /
    SAP Financial Transactions) plus the cross-source
    Reconciliation tab; they are NOT joined back into
    ``detect_back_billing`` because:

      * SAP financial transactions carry a Document No. (e.g.
        ``531000424090``) not an Invoice #, and their Transaction
        Text is the generic ledger description
        (``Dr- Consum Billing Receivable`` etc.) -- they cannot
        be unambiguously matched to an inferred invoice.
      * SAP records have no ``Period From`` / ``Period To`` span
        (only Posting Date / Document Date) so they cannot
        independently drive a back-billing judgement.
      * The Reconciliation sheet is the proper place to surface
        agreements and disagreements between the inferred and
        SAP samples; naively joining SAP amounts into the
        backbilling tab would mislead the reviewer.
    If a future resource joins the two sources by a higher-fidelity
    key (e.g. PDF receipt number + SAP Document No. mapping
    table), wire the intersection through ``run_analysers`` here.
    """
    columns = [
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
    if df is None or df.empty:
        return pd.DataFrame(columns=columns)
    has_admit = "Cancel/Rebill Admitted" in df.columns
    rows = []
    for _, r in df.iterrows():
        pf = _safe_to_datetime(r.get("Period From"))
        pt = _safe_to_datetime(r.get("Period To"))
        if pd.isna(pf) or pd.isna(pt):
            continue
        days = int((pt - pf).days)
        if days <= 365:
            continue
        net_raw = r.get("Amount (£)", 0)
        try:
            net = float(net_raw)
        except (TypeError, ValueError):
            net = 0.0
        admitted = bool(r.get("Cancel/Rebill Admitted")) if has_admit else False
        bill_date_raw = r.get("Date", "")
        bill_date_dt = _safe_to_datetime(bill_date_raw)
        rows.append(
            {
                "Invoice #": r.get("Invoice #", ""),
                "Bill Date": bill_date_raw,
                "_bill_date_sort": bill_date_dt if not pd.isna(bill_date_dt) else pd.Timestamp.max,
                "Period From": pf,
                "Period To": pt,
                "Days Billed": days,
                "Net Charge (£)": net,
                "12-Month Limit (days)": 365,
                "Excess Days": days - 365,
                "Cancel/Rebill Admitted": admitted,
                "Reason Assessment": _assess_reason(r.get("Invoice #", ""), days, admitted, pf, pt),
            }
        )
    out = pd.DataFrame(rows)
    if out.empty:
        return pd.DataFrame(columns=columns)
    sort_key = out["_bill_date_sort"]
    out = out.drop(columns=["_bill_date_sort"])
    # Reorder rows by the sort key (parsed Bill Date, ascending).
    out = out.loc[sort_key.sort_values().index].reset_index(drop=True)
    return out[columns]


# --- write_back_billing_sheet (was writers/__init__.py L2809-3018) ---


def write_back_billing_sheet(
    ws: Worksheet,
    bb: pd.DataFrame,
    account: str = "",
    overlapping_invoices: set[str] | None = None,
    evidence_df: pd.DataFrame | None = None,
    evidence_index: dict[str, int] | None = None,
) -> None:
    """Render the Back-billing Analysis tab.

    Layout follows the design spec (§4.1):
      row 1: title banner with SAP account
      row 2: 'LEGAL CONTEXT' section label
      row 3: legal_context() body (one merged paragraph)
      row 4: empty
      row 5: short instruction
      row 6: empty
      row 7: column headers (11 cols incl. Open PDF)
      rows 8+: data rows (sorted by Bill Date as produced by
              :func:`detect_back_billing`)
      trailing: 'TOTAL RETROSPECTIVE CHARGES IN BACK-BILLED INVOICES'

    The ``Cancel/Rebill Disclosed`` cell (col 9) is the
    :func:`_disclosed_label` value taking the row's
    ``Cancel/Rebill Admitted`` bool AND whether this invoice also
    appears in ``overlapping_invoices`` (a set populated by the
    rebilling detector; defaults to empty).

    Open PDF column (col 11) carries hyperlink
    the first ~400 chars of the source PDF text so a reviewer can
    see why N/A entries were N/A and which regex produced which value.
    """
    ws.title = "Back-billing Analysis"
    NAVY = "10367A"
    ORANGE = "FE5716"
    overlaps = overlapping_invoices or set()

    # Row 1: banner with account
    title = "BACK-BILLING EVENTS ANALYSIS"
    if account:
        title = f"{title}  |  Account {account}"
    t1 = ws.cell(row=1, column=1, value=title)
    t1.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    t1.fill = PatternFill("solid", start_color=ORANGE)
    t1.border = CELL_BORDER
    t1.alignment = Alignment(horizontal="left", vertical="center")
    for c in range(2, 12):
        x = ws.cell(row=1, column=c)
        x.fill = PatternFill("solid", start_color=ORANGE)
        x.border = CELL_BORDER
    ws.row_dimensions[1].height = 22

    # Row 2: 'LEGAL CONTEXT' label
    lc_hdr = ws.cell(row=2, column=1, value="LEGAL CONTEXT")
    lc_hdr.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    lc_hdr.fill = PatternFill("solid", start_color=NAVY)
    lc_hdr.border = CELL_BORDER
    for c in range(2, 12):
        x = ws.cell(row=2, column=c)
        x.fill = PatternFill("solid", start_color=NAVY)
        x.border = CELL_BORDER

    # Row 3: legal_context body (merged across the whole width so the
    # paragraph is readable in one cell).
    lc_text = legal_context()
    lc_cell = ws.cell(row=3, column=1, value=lc_text)
    lc_cell.font = Font(name="Calibri", size=10)
    lc_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    lc_cell.border = CELL_BORDER
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=11)
    ws.row_dimensions[3].height = 90

    # Row 5: instruction
    inst = (
        "Each row identifies an invoice where EDF billed more than 12 "
        "months retrospectively. The Excess Days column shows by how "
        "many days beyond the Standard Licence Condition 7A (SLC 7A) "
        "12-month limit the invoice went."
    )
    inst_cell = ws.cell(row=5, column=1, value=inst)
    inst_cell.font = Font(name="Calibri", size=10, italic=True)
    inst_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=11)
    ws.row_dimensions[5].height = 45

    # Row 7: headers
    headers = [
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
        "Open PDF",
        "View on Evidence Report",
    ]
    for col, h in enumerate(headers, 1):
        _hcell(ws, 7, col, h, bg=NAVY)
    ws.row_dimensions[7].height = 28

    # Data rows + running total
    r = 8
    total = 0.0
    alt_fill = PatternFill("solid", start_color="EEF2FF")
    for _, row in bb.iterrows():
        row_fill = alt_fill if r % 2 == 0 else PatternFill()
        bg = None if row_fill.fill_type is None else "EEF2FF"
        inv = str(row.get("Invoice #", ""))
        overlap_flag = inv in overlaps
        disclosed = _disclosed_label(bool(row.get("Cancel/Rebill Admitted")), overlap_flag)
        net = float(row.get("Net Charge (£)", 0.0) or 0.0)
        total += net
        bill_date_val = row.get("Bill Date", "")
        if isinstance(bill_date_val, pd.Timestamp | datetime):
            bill_date_val = bill_date_val.strftime("%d %b %Y")
        pf = row.get("Period From")
        if isinstance(pf, pd.Timestamp | datetime):
            pf = pf.strftime("%d %b %Y")
        pt = row.get("Period To")
        if isinstance(pt, pd.Timestamp | datetime):
            pt = pt.strftime("%d %b %Y")
        _text(ws, r, 1, inv, fill_hex=bg)
        _text(ws, r, 2, bill_date_val, fill_hex=bg)
        _text(ws, r, 3, pf, fill_hex=bg)
        _text(ws, r, 4, pt, fill_hex=bg)
        _num(ws, r, 5, int(row.get("Days Billed", 0)), fmt="#,##0", fill_hex=bg)
        _money(ws, r, 6, net, fill_hex=bg)
        _num(ws, r, 7, int(row.get("12-Month Limit (days)", 365)), fmt="#,##0", fill_hex=bg)
        _num(ws, r, 8, int(row.get("Excess Days", 0)), fmt="#,##0", fill_hex=bg)
        # Highlight excess-days when >30 (i.e. back-billing is materially over)
        if int(row.get("Excess Days", 0)) > 30:
            ws.cell(row=r, column=8).font = Font(name="Calibri", size=10, bold=True, color="C00000")
        _text(ws, r, 9, disclosed, fill_hex=bg)
        _text(ws, r, 10, row.get("Reason Assessment", ""), wrap=True, fill_hex=bg)
        _open_pdf_hyperlink_cell(ws, r, 11, evidence_df, inv)
        # View on Evidence Report (col 12): bidirectional hotlink back to the
        # row on the EDF Evidence Report sheet. Match by Invoice # first,
        # falling back to the amt|days signature.
        target_row = None
        if evidence_index is not None:
            target_row = evidence_index.get(f"inv:{inv}")
            if target_row is None:
                try:
                    amt = float(row.get("Net Charge (£)", 0.0) or 0.0)
                    days = int(row.get("Days Billed", 0) or 0)
                    key = f"amt_days:{amt:.2f}|{days}"
                    target_row = evidence_index.get(key)
                except (TypeError, ValueError):
                    pass
        if target_row is not None:
            cell = ws.cell(row=r, column=12, value="→")
            cell.hyperlink = openpyxl.worksheet.hyperlink.Hyperlink(
                ref=cell.coordinate,
                location=f"'EDF Evidence Report'!A{target_row}",
                display="→",
                tooltip=f"Jump to EDF Evidence Report!A{target_row}",
            )
            cell.font = Font(name="Calibri", size=10, color="0563C1", underline="single")
        else:
            cell = ws.cell(row=r, column=12, value="No match")
            cell.font = Font(name="Calibri", size=10, italic=True, color="A6A6A6")
        r += 1

    # Trailing totals row
    if not bb.empty:
        total_label = "TOTAL RETROSPECTIVE CHARGES IN BACK-BILLED INVOICES"
        label_cell = ws.cell(row=r, column=1, value=total_label)
        label_cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        label_cell.fill = PatternFill("solid", start_color=NAVY)
        label_cell.border = CELL_BORDER
        label_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        for c in range(2, 6):
            x = ws.cell(row=r, column=c)
            x.fill = PatternFill("solid", start_color=NAVY)
            x.border = CELL_BORDER
        total_cell = ws.cell(row=r, column=6, value=total)
        total_cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        total_cell.fill = PatternFill("solid", start_color=NAVY)
        total_cell.border = CELL_BORDER
        total_cell.number_format = "#,##0.00"
        for c in range(7, 13):
            x = ws.cell(row=r, column=c)
            x.fill = PatternFill("solid", start_color=NAVY)
            x.border = CELL_BORDER
        ws.row_dimensions[r].height = 22
        r += 1

    # Column widths
    widths = {
        "A": 18,
        "B": 14,
        "C": 14,
        "D": 14,
        "E": 12,
        "F": 16,
        "G": 18,
        "H": 12,
        "I": 22,
        "J": 60,
        "K": 60,  # Open PDF
        "L": 22,  # View on Evidence Report
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width
    ws.freeze_panes = "A8"
