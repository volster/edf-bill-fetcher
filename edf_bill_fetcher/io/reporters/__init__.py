"""Reporters — PDF and DOCX report rendering for the evidence bundle.

This package wraps the two historical top-level renderer modules,
``edf_report`` (PDF / reportlab) and ``edf_report_docx`` (DOCX /
python-docx), so callers can use the package layout while the
compat-shim refactor window is open:

    from edf_bill_fetcher.io.reporters import (
        generate_pdf_from_gui,
        generate_docx_from_gui,
    )

Each submodule (``pdf_report``, ``docx_report``) is a thin
re-export shim — the implementation lives in
``edf_report.py`` / ``edf_report_docx.py`` because they predate
the modularization refactor and are imported directly by
``edf_collector.py`` and the test suite.  Task 7 strips the compat
layer and moves the implementations into these submodules.
"""

from __future__ import annotations

from edf_bill_fetcher.io.reporters.docx_report import (
    generate_docx_from_gui,
    generate_ombudsman_docx,
)
from edf_bill_fetcher.io.reporters.pdf_report import (
    REPORT_SECTIONS,
    RenderContext,
    fmt_date,
    fmt_money,
    fmt_number,
    fmt_pct,
    generate_ombudsman_pdf,
    generate_pdf_from_gui,
)

__all__ = [
    "REPORT_SECTIONS",
    "RenderContext",
    "fmt_date",
    "fmt_money",
    "fmt_number",
    "fmt_pct",
    "generate_docx_from_gui",
    "generate_ombudsman_docx",
    "generate_ombudsman_pdf",
    "generate_pdf_from_gui",
]
