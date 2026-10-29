"""Workbook presentation of accepted, evidence-scoped data-quality checks."""

from __future__ import annotations

from collections import Counter

from openpyxl.styles import Alignment, Font

from edf_bill_fetcher.helpers.excel_utils import hcell, num, section_hdr, text
from edf_bill_fetcher.models.quality_tables import quality_notice, quality_table
from edf_bill_fetcher.models.report_models import compute_data_quality_report


def write_data_quality_sheet(ws, df, *, payload=None):
    """Use shared check scopes/statuses and separate distributions from quality checks."""
    ws.title = "Data Quality Report"
    report = payload if payload is not None else compute_data_quality_report(df)
    if not report.total_records:
        hcell(ws, 1, 1, "No data to analyze", bg="10367A")
        ws.column_dimensions["A"].width = 40
        return

    ws.merge_cells("A1:E1")
    hcell(ws, 1, 1, "EDF ENERGY DISPUTE — DATA QUALITY REPORT", bg="10367A")
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A2:E4")
    cell = text(ws, 2, 1, quality_notice(report), wrap=True)
    cell.alignment = Alignment(vertical="top", wrap_text=True)
    for row in range(2, 5):
        ws.row_dimensions[row].height = 26

    table = quality_table(report)
    for column, label in enumerate(table[0], 1):
        hcell(ws, 5, column, label, bg="007F82")
    for row, values in enumerate(table[1:], 6):
        fill = "EEF5F8" if row % 2 == 0 else None
        for column, value in enumerate(values, 1):
            if column in {2, 3}:
                num(ws, row, column, int(value), fill_hex=fill)
            else:
                text(ws, row, column, value, fill_hex=fill, wrap=True)
        if values[-1] == "REVIEW":
            ws.cell(row, 5).font = Font(name="Calibri", bold=True, color="9C5700")
        ws.row_dimensions[row].height = 30
    last_check = 5 + len(table) - 1
    ws.auto_filter.ref = f"A5:E{last_check}"

    row = last_check + 2
    statuses = Counter(values[-1] for values in table[1:])
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    summary = f"{len(table) - 1} checks | " + " | ".join(
        f"{label}: {statuses[label]}" for label in ("PASS", "REVIEW", "INFO", "N/A")
    )
    hcell(ws, row, 1, summary, bg="10367A")
    ws.row_dimensions[row].height = 32

    for title, distribution in (
        ("SOURCE DISTRIBUTION", report.source_distribution),
        ("ENTRY TYPE DISTRIBUTION", report.entry_type_distribution),
    ):
        row += 2
        section_hdr(ws, row, title, ncols=5)
        for label, count in distribution.items():
            row += 1
            text(ws, row, 1, label, wrap=True)
            num(ws, row, 2, count)
            num(ws, row, 4, count / report.total_records, fmt="0.0%")
            text(ws, row, 5, "INFO")
            ws.row_dimensions[row].height = 28

    for letter, width in {"A": 43, "B": 12, "C": 12, "D": 14, "E": 14}.items():
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = "B6"
    ws.print_options.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:5"
    ws.print_area = f"A1:E{row}"


__all__ = ["write_data_quality_sheet"]
