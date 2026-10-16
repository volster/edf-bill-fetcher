"""Stored bill revision decisions, with complete references and no ledger adjustment."""

from openpyxl.workbook import Workbook

from edf_bill_fetcher.helpers.excel_utils import money, text
from edf_bill_fetcher.io.writers.sheet_layout import (
    write_banner,
    write_header_row,
    write_merged_text,
)
from edf_bill_fetcher.models.output_tables import ledger_label
from edf_bill_fetcher.models.report_view import ReportView


def write_revision_sheet(wb: Workbook, view: ReportView) -> None:
    """Render the accepted relationship model without recalculating suggestions."""
    model = view.revisions
    if model is None:
        return
    ws = wb.create_sheet("Bill revisions")
    headers = (
        "Decision",
        "Relationship",
        "Account / ledger",
        "Earlier invoice",
        "Earlier date",
        "Later invoice / credit",
        "Later date",
        "Earlier charge (£)",
        "Later charge (£)",
        "Document charge difference (£)",
        "Earlier included",
        "Later included",
        "Changed fields",
        "Reason / suggestion basis",
        "Earlier source",
        "Earlier pages",
        "Later source",
        "Later pages",
        "Earlier row ID",
        "Later row ID",
    )
    write_banner(ws, "BILL REVISION REVIEW", len(headers))
    write_merged_text(ws, 2, model["notice"], len(headers), height=50)
    write_merged_text(
        ws,
        3,
        "Compare the original and later source PDFs before confirming. This reference sheet is read-only; use Bill revisions in the app to record decisions. Document charge differences are not refunds or changes to account liability."
        + (
            " Suggestions capped at 500; all explicit decisions are retained."
            if model["truncated"]
            else ""
        ),
        len(headers),
        height=45,
    )
    write_header_row(ws, 4, list(headers), height=34)
    for ordinal, link in enumerate(model["relationships"], 5):
        left, right = link["earlier"], link["later"]
        values = [
            link["status"].title(),
            link["kind"].replace("_", " ").title(),
            ledger_label(left),
            left.get("Invoice #"),
            left.get("Date"),
            right.get("Invoice #"),
            right.get("Date"),
            left.get("Period Charge (£)"),
            right.get("Period Charge (£)"),
            link["charge_difference"],
            "Yes" if link["earlier_included"] else "No",
            "Yes" if link["later_included"] else "No",
            ", ".join(link["changed_fields"]),
            link["reason"] or link["basis"],
            left.get("Attachment Name"),
            _pages(left),
            right.get("Attachment Name"),
            _pages(right),
            link["older"],
            link["newer"],
        ]
        for column, value in enumerate(values, 1):
            if column in {8, 9, 10}:
                money(ws, ordinal, column, value)
            else:
                text(ws, ordinal, column, value, wrap=True)
        ws.row_dimensions[ordinal].height = 60
    if not model["relationships"]:
        text(
            ws,
            5,
            1,
            "No revision relationships found. Use the app selectors to add a missed relationship.",
        )
    ws.freeze_panes = "D5"
    ws.auto_filter.ref = f"A4:T{max(4, ws.max_row)}"
    for letter in "ABCDEFGHIJKLMNOPQRST":
        ws.column_dimensions[letter].width = 24 if letter not in "CNOQ" else 42
    ws.column_dimensions.group("S", "T", hidden=True)
    ws.print_title_rows = "1:4"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0


def _pages(row: dict) -> str:
    return "–".join(str(row[key]) for key in ("Page From", "Page To") if row.get(key) is not None)
