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
    bill_date: pd.Timestamp,
    excess: int,
    admitted: bool,
    period_from: pd.Timestamp,
    period_to: pd.Timestamp,
) -> str:
    """Return a short, deterministic narrative for the Reason Assessment column of the Back-billing sheet.

    Template-driven (no LLM).  The narrative is keyed to the legally
    correct back-billing rule (SLC 7A / Electricity Act 1989 s.84B):
    a bill is back-billing when it charges for consumption supplied
    more than 12 months before the bill Date.  ``excess`` is the count
    of consumption days in the period that fall more than 365 days
    before ``bill_date``.
    """
    pf = period_from.strftime("%d %b %Y")
    pt = period_to.strftime("%d %b %Y")
    bd = bill_date.strftime("%d %b %Y")
    if admitted:
        head = (
            f"Invoice {invoice} billed on {bd} for consumption from {pf} to {pt}; "
            f"{excess} days of consumption were supplied more than 12 months before the bill, "
            "exceeding the SLC 7A back-billing limit. "
            "EDF's cover page admits a cancellation/reversal, which is "
            "direct evidence the bill is a back-billing remedy."
        )
    else:
        head = (
            f"Invoice {invoice} billed on {bd} for consumption from {pf} to {pt}; "
            f"{excess} days of consumption were supplied more than 12 months before the bill, "
            "exceeding the SLC 7A back-billing limit. No "
            "admit-phrase was found on the cover page."
        )
    return head


# --- detect_back_billing (was writers/__init__.py L2704-2806) ---


def _pull_period_charge(r: pd.Series) -> tuple[float, str]:
    """Pull ``Period Charge (£)`` from the source row; fall back to ``Amount (£)``.

    Returns ``(charge, value_source)`` where ``value_source`` is
    ``"Period Charge"`` when the Period Charge column was used, or
    ``"Amount (fallback)"`` when Period Charge was absent, N/A, or
    unparseable and the Amount column was used instead.
    """
    pc_raw = r.get("Period Charge (£)")
    if pc_raw is not None:
        try:
            return float(pc_raw), "Period Charge"
        except (TypeError, ValueError):
            pass
    amt_raw = r.get("Amount (£)", 0)
    try:
        return float(amt_raw), "Amount (fallback)"
    except (TypeError, ValueError):
        return 0.0, "Amount (fallback)"


def detect_back_billing(df: pd.DataFrame) -> pd.DataFrame:
    """Return invoices that are back-billing under SLC 7A / Electricity Act 1989 s.84B.

    A bill is back-billing when it charges for consumption supplied
    more than 12 months before the bill Date.  The eligibility gate is
    ``Date - Period To > 365 days`` — i.e. the bill was issued more
    than 12 months after the LATEST consumption it charges for.  If
    even the latest consumption (Period To) is within 365 days of the
    bill Date, the invoice is NOT back-billing (regardless of how long
    the period span is).

    ``Excess Days = max(0, (Date - 365 days - Period From).days)`` —
    the count of consumption days in the period that fall more than
    365 days before the bill Date.

    The detector also pulls ``Period Charge (£)`` from the source
    record; if that column is absent, N/A, or unparseable, it falls
    back to ``Amount (£)`` and records the provenance in the
    ``Value Source`` column.

    The function tolerates a missing ``Cancel/Rebill Admitted``
    column (treated as ``False``).

    Output columns:
        Invoice #, Bill Date, Period From, Period To, Days Billed,
        Period Charge (£), Value Source, 12-Month Limit (days),
        Excess Days, Cancel/Rebill Admitted, Reason Assessment.

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
        "Period Charge (£)",
        "Value Source",
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
        bill_date_dt = _safe_to_datetime(r.get("Date"))
        if pd.isna(bill_date_dt):
            continue
        # Legal gate: bill Date must be more than 365 days after Period To.
        gap_to = int((bill_date_dt - pt).days)
        if gap_to <= 365:
            continue
        days = int((pt - pf).days)
        # Excess Days: consumption days supplied more than 365 days before bill Date.
        excess = max(0, int((bill_date_dt - pd.Timedelta(days=365) - pf).days))
        # Period Charge (£) with Amount (£) fallback.
        charge, value_source = _pull_period_charge(r)
        admitted = bool(r.get("Cancel/Rebill Admitted")) if has_admit else False
        bill_date_raw = r.get("Date", "")
        rows.append(
            {
                "Invoice #": r.get("Invoice #", ""),
                "Bill Date": bill_date_raw,
                "_bill_date_sort": bill_date_dt,
                "Period From": pf,
                "Period To": pt,
                "Days Billed": days,
                "Period Charge (£)": charge,
                "Value Source": value_source,
                "12-Month Limit (days)": 365,
                "Excess Days": excess,
                "Cancel/Rebill Admitted": admitted,
                "Reason Assessment": _assess_reason(
                    r.get("Invoice #", ""),
                    bill_date_dt,
                    excess,
                    admitted,
                    pf,
                    pt,
                ),
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
    domination_map: dict[str, tuple[str, bool]] | None = None,
) -> None:
    """Render the Back-billing Analysis tab.

    Layout follows the design spec (§4.1):
      row 1: title banner with SAP account
      row 2: 'LEGAL CONTEXT' section label
      row 3: legal_context() body (one merged paragraph)
      row 4: empty
      row 5: short instruction
      row 6: empty
      row 7: column headers (16 cols incl. Open PDF, View on Evidence
              Report, Status, Superseded By, Partial Overlap)
      rows 8+: data rows (sorted by Bill Date as produced by
              :func:`detect_back_billing`)
      trailing: 'TOTAL RETROSPECTIVE CHARGES IN BACK-BILLED INVOICES'

    The ``Cancel/Rebill Disclosed`` cell (col 10) is the
    :func:`_disclosed_label` value taking the row's
    ``Cancel/Rebill Admitted`` bool AND whether this invoice also
    appears in ``overlapping_invoices`` (a set populated by the
    rebilling detector; defaults to empty).

    Open PDF column (col 12) carries hyperlink
    the first ~400 chars of the source PDF text so a reviewer can
    see why N/A entries were N/A and which regex produced which value.

    ``domination_map`` (from :func:`compute_transitive_domination`)
    maps ``superseded_invoice_id -> (survivor_invoice_id, partial_overlap)``.
    Rows whose ``Invoice #`` is a key in this map are rendered as
    outline-collapsed sub-rows (``outline_level=1``, ``hidden=True``,
    mirroring ``io/writers/sap.py:440``) with ``Status="Superseded"``,
    ``Superseded By=survivor``, ``Partial Overlap="Yes"`` when the flag
    is set. They are preserved for the audit trail but EXCLUDED from
    the trailing total. Rows not in the map are ``Status="Live"`` and
    their ``Period Charge (£)`` is added to the total.
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
    for c in range(2, 17):
        x = ws.cell(row=1, column=c)
        x.fill = PatternFill("solid", start_color=ORANGE)
        x.border = CELL_BORDER
    ws.row_dimensions[1].height = 22

    # Row 2: 'LEGAL CONTEXT' label
    lc_hdr = ws.cell(row=2, column=1, value="LEGAL CONTEXT")
    lc_hdr.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    lc_hdr.fill = PatternFill("solid", start_color=NAVY)
    lc_hdr.border = CELL_BORDER
    for c in range(2, 17):
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
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=16)
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
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=16)
    ws.row_dimensions[5].height = 45

    # Row 7: headers
    headers = [
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
        charge = float(row.get("Period Charge (£)", 0.0) or 0.0)
        value_src = str(row.get("Value Source", ""))
        bill_date_val = row.get("Bill Date", "")
        if isinstance(bill_date_val, pd.Timestamp | datetime):
            bill_date_val = bill_date_val.strftime("%d %b %Y")
        pf = row.get("Period From")
        if isinstance(pf, pd.Timestamp | datetime):
            pf = pf.strftime("%d %b %Y")
        pt = row.get("Period To")
        if isinstance(pt, pd.Timestamp | datetime):
            pt = pt.strftime("%d %b %Y")

        # Domination: is this invoice superseded by another?
        superseded_by = ""
        partial_overlap = ""
        if domination_map is not None and inv in domination_map:
            survivor, partial = domination_map[inv]
            superseded_by = survivor
            partial_overlap = "Yes" if partial else ""
            status = "Superseded"
        else:
            status = "Live"
            total += charge

        _text(ws, r, 1, inv, fill_hex=bg)
        _text(ws, r, 2, bill_date_val, fill_hex=bg)
        _text(ws, r, 3, pf, fill_hex=bg)
        _text(ws, r, 4, pt, fill_hex=bg)
        _num(ws, r, 5, int(row.get("Days Billed", 0)), fmt="#,##0", fill_hex=bg)
        _money(ws, r, 6, charge, fill_hex=bg)
        _text(ws, r, 7, value_src, fill_hex=bg)
        _num(ws, r, 8, int(row.get("12-Month Limit (days)", 365)), fmt="#,##0", fill_hex=bg)
        _num(ws, r, 9, int(row.get("Excess Days", 0)), fmt="#,##0", fill_hex=bg)
        # Highlight excess-days when >30 (i.e. back-billing is materially over)
        if int(row.get("Excess Days", 0)) > 30:
            ws.cell(row=r, column=9).font = Font(name="Calibri", size=10, bold=True, color="C00000")
        _text(ws, r, 10, disclosed, fill_hex=bg)
        _text(ws, r, 11, row.get("Reason Assessment", ""), wrap=True, fill_hex=bg)
        _open_pdf_hyperlink_cell(ws, r, 12, evidence_df, inv)
        # View on Evidence Report (col 13): bidirectional hotlink back to the
        # row on the EDF Evidence Report sheet. Match by Invoice # first,
        # falling back to the amt|days signature.
        target_row = None
        if evidence_index is not None:
            target_row = evidence_index.get(f"inv:{inv}")
            if target_row is None:
                try:
                    amt = float(row.get("Period Charge (£)", 0.0) or 0.0)
                    days = int(row.get("Days Billed", 0) or 0)
                    key = f"amt_days:{amt:.2f}|{days}"
                    target_row = evidence_index.get(key)
                except (TypeError, ValueError):
                    pass
        if target_row is not None:
            cell = ws.cell(row=r, column=13, value="→")
            cell.hyperlink = openpyxl.worksheet.hyperlink.Hyperlink(
                ref=cell.coordinate,
                location=f"'EDF Evidence Report'!A{target_row}",
                display="→",
                tooltip=f"Jump to EDF Evidence Report!A{target_row}",
            )
            cell.font = Font(name="Calibri", size=10, color="0563C1", underline="single")
        else:
            cell = ws.cell(row=r, column=13, value="No match")
            cell.font = Font(name="Calibri", size=10, italic=True, color="A6A6A6")
        # Status / Superseded By / Partial Overlap (cols 14-16)
        _text(ws, r, 14, status, fill_hex=bg)
        _text(ws, r, 15, superseded_by, fill_hex=bg)
        _text(ws, r, 16, partial_overlap, fill_hex=bg)
        # Superseded rows are outline-collapsed sub-rows (mirrors
        # io/writers/sap.py:440) — preserved for the audit trail but
        # visually grouped under their surviving invoice.
        if status == "Superseded":
            ws.row_dimensions[r].outline_level = 1
            ws.row_dimensions[r].hidden = True
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
        for c in range(7, 17):
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
        "H": 14,
        "I": 12,
        "J": 22,
        "K": 60,
        "L": 60,  # Open PDF
        "M": 22,  # View on Evidence Report
        "N": 14,  # Status
        "O": 16,  # Superseded By
        "P": 16,  # Partial Overlap
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width
    ws.freeze_panes = "A8"
