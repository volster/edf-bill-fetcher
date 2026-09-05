"""Shared utility helpers extracted from edf_collector.py.

Submodules:
- ``formatting`` — number/text formatting helpers (currency, integer, account matching).
- ``pst_resources`` — PST attachment-filename walking and sender-email extraction.
- ``date_utils`` — pandas time-series statistics and evidence-trail builder.
- ``excel_utils`` — openpyxl cell primitives, SAP row index map, text-warning suppression.
- ``pdf_utils`` — placeholder for future PDF-specific helpers.
- ``theme`` — EDF brand colours, greyscale helpers, SAP back-billing fill pairs,
  and the openpyxl ``Side``/``Border`` instances used across the evidence workbook.
"""

from edf_bill_fetcher.helpers.theme import (
    CELL_BORDER,
    DUP_GREY,
    EDF_NAVY,
    EDF_OFFWHITE,
    EDF_ORANGE,
    MEDIUM_GREY,
    NAVY_BLUE,
    ORANGE,
    SAP_BB_DETAIL_FILL_PAIR,
    SAP_BB_MEDIUM_BORDER,
    SAP_BB_SUMMARY_FILL_PAIR,
)

from . import date_utils, excel_utils, formatting, pdf_utils, pst_resources, theme

__all__ = [
    "CELL_BORDER",
    "DUP_GREY",
    "EDF_NAVY",
    "EDF_OFFWHITE",
    "EDF_ORANGE",
    "MEDIUM_GREY",
    "NAVY_BLUE",
    "ORANGE",
    "SAP_BB_DETAIL_FILL_PAIR",
    "SAP_BB_MEDIUM_BORDER",
    "SAP_BB_SUMMARY_FILL_PAIR",
    "date_utils",
    "excel_utils",
    "formatting",
    "pdf_utils",
    "pst_resources",
    "theme",
]
