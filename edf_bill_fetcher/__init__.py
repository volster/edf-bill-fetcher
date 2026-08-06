"""edf-bill-fetcher — Python package.

Top-level public API.  The package was carved out of ``edf_collector.py``
as part of the modularization refactor.  Each submodule owns a domain:

- ``collectors`` — extraction orchestrators (EvidenceEngine)
- ``helpers`` — shared utility helpers (formatting, date_utils, excel_utils, pdf_utils)
- ``models`` — typed dataclasses (SapBackBillingEvent, SapEdfMatch)
- ``ui`` — Tkinter application classes (App, ReportOptionsDialog)
- ``writers`` — Excel sheet writers (write_reconciliation_sheet, export_to_excel, etc.)

The top-level ``__init__.py`` does NOT eagerly import from submodules
because doing so triggers a circular import between the submodule
groups.  Callers should use submodule imports directly:

  from edf_bill_fetcher.collectors import EvidenceEngine
  from edf_bill_fetcher.io.writers import export_to_excel
  from edf_bill_fetcher.models import SapBackBillingEvent
"""

__all__: list[str] = []
