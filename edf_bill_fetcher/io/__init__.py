"""I/O submodules — writers, adapters, reporters, CLI.

The ``io`` namespace groups every framework-coupled boundary of the
evidence pipeline:

- ``writers`` — Excel workbook emission (``openpyxl``)
- ``adapters`` — file reading (PDF, PST, HTM)
- ``reporters`` — PDF + DOCX report rendering
- ``cli`` — ``argparse`` entry points for headless extraction and report generation

Each submodule is independently importable so callers can pick the
narrowest import path they need.
"""

from edf_bill_fetcher.io.cli import (
    _safe_pickle_load,
    main,
    run_cli_docx_report,
    run_cli_extract,
    run_cli_pdf_report,
)

__all__ = [
    "run_cli_extract",
    "run_cli_pdf_report",
    "run_cli_docx_report",
    "main",
    "_safe_pickle_load",
]
