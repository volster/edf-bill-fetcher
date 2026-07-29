"""Theme constants — colours, fills, and borders used across the evidence workbook.

Extracted from ``edf_collector.py`` as part of the modularization refactor
(Task 2).  This module is the single source of truth for:

- EDF brand colours (``EDF_NAVY``, ``EDF_ORANGE``, ``EDF_OFFWHITE``)
- Greyscale helpers (``MEDIUM_GREY``, ``DUP_GREY``)
- Hex-without-hash aliases (``ORANGE``, ``NAVY_BLUE``) used by the
  SAP back-billing analyser sheets
- SAP back-billing fill pairs (``SAP_BB_SUMMARY_FILL_PAIR``,
  ``SAP_BB_DETAIL_FILL_PAIR``) — alternating row colours for the
  summary and detail sheets
- ``SAP_BB_MEDIUM_BORDER`` — the medium-weight navy border used on
  the SAP back-billing summary sheet
- ``CELL_BORDER`` — the thin grey border applied to every cell on
  every sheet (originally defined in ``excel_utils.py``; moved here
  so all visual primitives live in one place)

All constants are module-level and fully type-annotated.  ``Side``
and ``Border`` are imported from ``openpyxl.styles`` because the
border objects are openpyxl instances, not plain strings.
"""

from __future__ import annotations

from openpyxl.styles import Border, Side

# ---------------------------------------------------------------------------
# EDF brand colours
# ---------------------------------------------------------------------------
EDF_NAVY: str = "#10367A"
EDF_ORANGE: str = "#FE5716"
EDF_OFFWHITE: str = "#F5F5F5"

# ---------------------------------------------------------------------------
# Greyscale helpers
# ---------------------------------------------------------------------------
MEDIUM_GREY: str = "#666666"
DUP_GREY: str = "E0E0E0"

# ---------------------------------------------------------------------------
# Hex-without-hash aliases (used by SAP back-billing analyser sheets)
# ---------------------------------------------------------------------------
ORANGE: str = "FE5716"
NAVY_BLUE: str = "10367A"

# ---------------------------------------------------------------------------
# SAP back-billing fill pairs (alternating row colours)
# ---------------------------------------------------------------------------
SAP_BB_SUMMARY_FILL_PAIR: tuple[str, str] = ("EFF4FB", "ffffff")
SAP_BB_DETAIL_FILL_PAIR: tuple[str, str] = ("F8FAFC", "ffffff")

# ---------------------------------------------------------------------------
# Borders
# ---------------------------------------------------------------------------
SAP_BB_MEDIUM_BORDER: Side = Side(style="medium", color="10367A")

CELL_BORDER: Border = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)


__all__ = [
    "EDF_NAVY",
    "EDF_ORANGE",
    "EDF_OFFWHITE",
    "MEDIUM_GREY",
    "DUP_GREY",
    "ORANGE",
    "NAVY_BLUE",
    "SAP_BB_SUMMARY_FILL_PAIR",
    "SAP_BB_DETAIL_FILL_PAIR",
    "SAP_BB_MEDIUM_BORDER",
    "CELL_BORDER",
]
