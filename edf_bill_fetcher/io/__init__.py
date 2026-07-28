"""I/O submodules — writers, adapters, reporters, CLI.

The ``io`` namespace groups every framework-coupled boundary of the
evidence pipeline:

- ``writers`` — Excel workbook emission (``openpyxl``)
- ``adapters`` — file reading (PDF, PST, HTM)
- ``reporters`` — PDF + DOCX report rendering
- ``cli`` (planned, Task 7) — ``argparse`` entry points

Each submodule is independently importable so callers can pick the
narrowest import path they need.
"""
