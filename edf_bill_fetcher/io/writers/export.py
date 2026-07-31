"""Compat re-export - export writer.

Implementation lives in ``edf_bill_fetcher.writers`` (the
legacy monolith at ``writers/__init__.py``). Real extraction
happens in later phases when the monolith is deleted.
This module exists so callers can ``from
edf_bill_fetcher.io.writers.export import ...``
per the test contract at ``tests/test_io_writers_extraction.py``.

The import from ``edf_bill_fetcher.writers`` is deferred to a PEP 562
module-level ``__getattr__`` to break a circular import:
``writers/__init__.py`` imports from ``io.writers.evidence`` (real
extraction) which loads ``io/writers/__init__.py`` which in turn
imports every shim in this package. A direct ``from
edf_bill_fetcher.writers import export_to_excel`` at module-init time
would re-enter ``writers/__init__.py`` mid-init and raise
``ImportError: partially initialized module``. Deferring to
``__getattr__`` lets Python finish loading the package before the
back-reference is resolved on first attribute access.
"""

from __future__ import annotations

__all__ = [
    "export_to_excel",
    "write_reconciliation_sheet",
]


def __getattr__(name: str):
    if name in ("export_to_excel", "write_reconciliation_sheet"):
        from edf_bill_fetcher.writers import export_to_excel, write_reconciliation_sheet

        return {"export_to_excel": export_to_excel, "write_reconciliation_sheet": write_reconciliation_sheet}[name]
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
