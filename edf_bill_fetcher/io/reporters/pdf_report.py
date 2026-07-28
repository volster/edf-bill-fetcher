"""PDF report renderer — re-exports ``edf_report`` generators.

The full implementation lives in the top-level ``edf_report.py``
module (which is the historical home of the PDF renderer and is
imported directly by ``edf_collector.py`` and the test suite).
This submodule exists so callers can reach the same entry points
via the package layout:

    from edf_bill_fetcher.io.reporters.pdf_report import (
        generate_ombudsman_pdf,
        generate_pdf_from_gui,
    )

Following the same thin-shim convention as
``edf_bill_fetcher/writers/analysis.py`` — implementation stays in
``edf_report.py`` until Task 7 strips the compat layer and the
module moves into this package.
"""

from __future__ import annotations

from edf_report import (
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
    "generate_ombudsman_pdf",
    "generate_pdf_from_gui",
]
