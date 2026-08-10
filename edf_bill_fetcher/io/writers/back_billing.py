"""Back-billing analysis writer — extracted from writers/__init__.py.

Contains: write_back_billing_sheet (renders the "Back-billing Analysis"
worksheet).  The pure-pandas detector ``detect_back_billing`` and its
private helpers ``_assess_reason`` / ``_pull_period_charge`` are
re-exported from :mod:`edf_bill_fetcher.processors.detection` (the
canonical home) so there is exactly one definition in the codebase.
"""

from __future__ import annotations

from datetime import datetime

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

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

# Re-exported from processors.detection (canonical home) for backwards
# compatibility — existing imports of these names from this module keep
# working, but the pipeline only ever uses the processors.detection copy.
from edf_bill_fetcher.processors.detection import (
    _assess_reason,  # noqa: F401 — re-exported for backwards compatibility
    _pull_period_charge,  # noqa: F401
    detect_back_billing,  # noqa: F401
)
from edf_bill_fetcher.writers._helpers import _disclosed_label

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
      row 7: column headers (17 cols incl. Unlawful Charge, Open PDF,
              View on Evidence Report, Status, Superseded By,
              Partial Overlap)
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
    their ``Period Charge (£)`` is added to the total. The
    ``Unlawful Charge (£)`` column is rendered for each row but is
    NOT summed into the trailing total — only ``Period Charge (£)``
    is totaled, consistent with the existing total-row pattern.
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
    for c in range(2, 18):
        x = ws.cell(row=1, column=c)
        x.fill = PatternFill("solid", start_color=ORANGE)
        x.border = CELL_BORDER
    ws.row_dimensions[1].height = 22

    # Row 2: 'LEGAL CONTEXT' label
    lc_hdr = ws.cell(row=2, column=1, value="LEGAL CONTEXT")
    lc_hdr.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    lc_hdr.fill = PatternFill("solid", start_color=NAVY)
    lc_hdr.border = CELL_BORDER
    for c in range(2, 18):
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
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=17)
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
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=17)
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
        "Unlawful Charge (£)",
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
        # Unlawful Charge (£): prorated share of Period Charge for the
        # Excess Days. Rendered as money; NOT summed into the trailing total.
        unlawful = float(row.get("Unlawful Charge (£)", 0.0) or 0.0)
        _money(ws, r, 10, unlawful, fill_hex=bg)
        _text(ws, r, 11, disclosed, fill_hex=bg)
        _text(ws, r, 12, row.get("Reason Assessment", ""), wrap=True, fill_hex=bg)
        _open_pdf_hyperlink_cell(ws, r, 13, evidence_df, inv)
        # View on Evidence Report (col 14): bidirectional hotlink back to the
        # row on the EDF Evidence Report sheet. Match by Invoice # first,
        # falling back to the amt|days signature.
        target_row = None
        if evidence_index is not None:
            target_row = evidence_index.get(f"inv:{inv}")
            if target_row is None:
                try:
                    amt = float(row.get("Amount (£)", 0.0) or 0.0)
                    days = int(row.get("Days Billed", 0) or 0)
                    key = f"amt_days:{amt:.2f}|{days}"
                    target_row = evidence_index.get(key)
                except (TypeError, ValueError):
                    pass
        if target_row is not None:
            cell = ws.cell(row=r, column=14, value="→")
            cell.hyperlink = openpyxl.worksheet.hyperlink.Hyperlink(
                ref=cell.coordinate,
                location=f"'EDF Evidence Report'!A{target_row}",
                display="→",
                tooltip=f"Jump to EDF Evidence Report!A{target_row}",
            )
            cell.font = Font(name="Calibri", size=10, color="0563C1", underline="single")
        else:
            cell = ws.cell(row=r, column=14, value="No match")
            cell.font = Font(name="Calibri", size=10, italic=True, color="A6A6A6")
        # Status / Superseded By / Partial Overlap (cols 15-17)
        _text(ws, r, 15, status, fill_hex=bg)
        _text(ws, r, 16, superseded_by, fill_hex=bg)
        _text(ws, r, 17, partial_overlap, fill_hex=bg)
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
        # Unlawful Charge total (col 10): sum of Unlawful Charge (£) across
        # Live rows only (Superseded rows are excluded, mirroring the Period
        # Charge total). Separate from the Period Charge total so a reviewer
        # can see the prorated unlawful exposure at a glance.
        unlawful_total = 0.0
        for _, _row in bb.iterrows():
            _inv = str(_row.get("Invoice #", ""))
            if domination_map is not None and _inv in domination_map:
                continue
            unlawful_total += float(_row.get("Unlawful Charge (£)", 0.0) or 0.0)
        unlawful_total_cell = ws.cell(row=r, column=10, value=round(unlawful_total, 2))
        unlawful_total_cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        unlawful_total_cell.fill = PatternFill("solid", start_color=NAVY)
        unlawful_total_cell.border = CELL_BORDER
        unlawful_total_cell.number_format = "#,##0.00"
        for c in range(7, 18):
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
        "J": 16,  # Unlawful Charge (£)
        "K": 22,  # Cancel/Rebill Disclosed
        "L": 60,  # Reason Assessment
        "M": 60,  # Open PDF
        "N": 22,  # View on Evidence Report
        "O": 14,  # Status
        "P": 16,  # Superseded By
        "Q": 16,  # Partial Overlap
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width
    ws.freeze_panes = "A8"
