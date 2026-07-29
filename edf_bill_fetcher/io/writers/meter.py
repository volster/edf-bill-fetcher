"""Compat re-export — meter writer.

Implementation lives in ``edf_bill_fetcher.writers`` (the legacy monolith
at ``writers/__init__.py``). Real extraction happens in later phases
when the monolith is deleted.

The public API exposed here matches the test contract at
``tests/test_io_writers_extraction.py``: ``write_meter_readings_sheet``
takes ``(ws, df)`` only. The underlying implementation accepts additional
optional arguments; sensible defaults (empty rollovers DataFrame) are
supplied by the adapter wrapper below.
"""

from __future__ import annotations

import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet

from edf_bill_fetcher.writers import (  # noqa: F401
    write_contract_history_sheet,
)
from edf_bill_fetcher.writers import (
    write_meter_readings_sheet as _write_meter_readings_sheet_impl,
)


def write_meter_readings_sheet(
    ws: Worksheet,
    df: pd.DataFrame,
    rollovers: pd.DataFrame | None = None,
    account: str = "",
    *,
    evidence_df: pd.DataFrame | None = None,
    evidence_index: dict[str, int] | None = None,
) -> None:
    """Adapter: test contract uses ``(ws, df)``; supply defaults for the rest."""
    return _write_meter_readings_sheet_impl(
        ws,
        df,
        rollovers if rollovers is not None else pd.DataFrame(),
        account,
        evidence_df=evidence_df,
        evidence_index=evidence_index,
    )


__all__ = [
    "write_contract_history_sheet",
    "write_meter_readings_sheet",
]
