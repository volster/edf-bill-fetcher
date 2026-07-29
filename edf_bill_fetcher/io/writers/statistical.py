"""Compat re-export — statistical writer.

Implementation lives in ``edf_bill_fetcher.writers`` (the
legacy monolith at ``writers/__init__.py``). Real extraction
happens in later phases when the monolith is deleted.
This module exists so callers can ``from
edf_bill_fetcher.io.writers.statistical import ...``
per the test contract at ``tests/test_io_writers_extraction.py``.
"""

from __future__ import annotations

from edf_bill_fetcher.writers import write_statistical_analysis_sheet  # noqa: F401

__all__ = [
    "write_statistical_analysis_sheet",
]
