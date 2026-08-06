"""Reporters — PDF and DOCX report rendering for the evidence bundle.

The two renderers moved into this package during the modularization:

    from edf_bill_fetcher.io.reporters import (
        generate_pdf_from_gui,
        generate_docx_from_gui,
    )

Each submodule (``pdf_report``, ``docx_report``) is the canonical
implementation; the pre-refactor top-level ``edf_report.py`` and
``edf_report_docx.py`` monoliths were deleted outright (no compat
layer).
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
