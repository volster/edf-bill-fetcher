"""Compat re-export — back_billing writer.

Implementation lives in ``edf_bill_fetcher.writers`` (the
legacy monolith at ``writers/__init__.py``). Real extraction
happens in later phases when the monolith is deleted.
This module exists so callers can ``from
edf_bill_fetcher.io.writers.back_billing import ...``
per the test contract at ``tests/test_io_writers_extraction.py``.
"""

from __future__ import annotations

from edf_bill_fetcher.writers import write_back_billing_sheet  # noqa: F401

__all__ = [
    "write_back_billing_sheet",
]
