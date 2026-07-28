"""DOCX report renderer — re-exports ``edf_report_docx`` generators.

The full implementation lives in the top-level ``edf_report_docx.py``
module (which is the historical home of the DOCX renderer and is
imported directly by ``edf_collector.py`` and the test suite).
This submodule exists so callers can reach the same entry points
via the package layout:

    from edf_bill_fetcher.io.reporters.docx_report import (
        generate_ombudsman_docx,
        generate_docx_from_gui,
    )

Following the same thin-shim convention as
``edf_bill_fetcher/writers/analysis.py`` — implementation stays in
``edf_report_docx.py`` until Task 7 strips the compat layer and the
module moves into this package.
"""

from __future__ import annotations

from edf_report_docx import (
    fmt_money,
    fmt_number,
    generate_docx_from_gui,
    generate_ombudsman_docx,
)

__all__ = [
    "fmt_money",
    "fmt_number",
    "generate_docx_from_gui",
    "generate_ombudsman_docx",
]
