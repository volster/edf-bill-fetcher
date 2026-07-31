"""Compat re-export - meter writer (with adapter).

Implementation lives in ``edf_bill_fetcher.writers`` (the legacy monolith
at ``writers/__init__.py``). Real extraction happens in later phases
when the monolith is deleted.

The public API exposed here matches the test contract at
``tests/test_io_writers_extraction.py``: ``write_meter_readings_sheet``
takes ``(ws, df)`` only. The underlying implementation accepts additional
optional arguments; sensible defaults (empty rollovers DataFrame) are
supplied by the adapter wrapper below.

The import from ``edf_bill_fetcher.writers`` is deferred into the
adapter function bodies to break a circular import:
``writers/__init__.py`` imports from ``io.writers.evidence`` (real
extraction) which loads ``io/writers/__init__.py`` which in turn
imports every shim in this package. Direct ``from
edf_bill_fetcher.writers import write_contract_history_sheet`` at
module-init time would re-enter ``writers/__init__.py`` mid-init and
raise ``ImportError: partially initialized module``. Deferring into
the function bodies lets Python finish loading the package before the
back-reference is resolved on first invocation.
``write_contract_history_sheet`` itself is exposed as a PEP 562
``__getattr__`` since it has no adapter wrapping.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

import pandas as pd

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet


__all__ = [
    "write_contract_history_sheet",
    "write_meter_readings_sheet",
]


def write_meter_readings_sheet(
    ws,
    df: pd.DataFrame,
    rollovers: pd.DataFrame | None = None,
    account: str = "",
    *,
    evidence_df: pd.DataFrame | None = None,
    evidence_index: dict[str, int] | None = None,
) -> None:
    """Adapter: test contract uses ``(ws, df)``; supply defaults for the rest."""
    from edf_bill_fetcher.writers import write_meter_readings_sheet as _impl

    return _impl(
        ws,
        df,
        rollovers if rollovers is not None else pd.DataFrame(),
        account,
        evidence_df=evidence_df,
        evidence_index=evidence_index,
    )


def __getattr__(name: str):
    if name == "write_contract_history_sheet":
        from edf_bill_fetcher.writers import write_contract_history_sheet

        return write_contract_history_sheet
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
