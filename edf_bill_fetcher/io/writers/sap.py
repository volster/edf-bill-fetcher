"""Compat re-export — SAP sheet writer.

Implementation lives in ``edf_bill_fetcher.writers`` (the legacy monolith
at ``writers/__init__.py``). Real extraction happens in later phases
when the monolith is deleted.

The public API exposed here matches the test contract at
``tests/test_io_writers_extraction.py``: ``write_sap_contract_history_sheet``
takes ``(ws, df)`` only; the underlying implementation expects
``(ws, rows: list[dict], account: str)``. The adapter converts the
DataFrame to list-of-dicts and supplies an empty default account.
"""

from __future__ import annotations

from typing import Any

import pandas as pd
from openpyxl.worksheet.worksheet import Worksheet

from edf_bill_fetcher.writers import (  # noqa: F401
    _write_sap_bb_events_sheet,
    _write_sap_bb_matches_sheet,
    _write_sap_header_row,
    write_sap_back_billing_sheets,
    write_sap_financial_transactions_sheet,
    write_sap_meter_readings_sheet,
)
from edf_bill_fetcher.writers import (
    write_sap_contract_history_sheet as _write_sap_contract_history_sheet_impl,
)


def write_sap_contract_history_sheet(
    ws: Worksheet,
    df_or_rows: pd.DataFrame | list[dict[str, Any]],
    account: str = "",
) -> None:
    """Adapter: test contract uses ``(ws, df)``; convert DataFrame to rows."""
    if isinstance(df_or_rows, pd.DataFrame):
        rows = df_or_rows.to_dict(orient="records")
    else:
        rows = df_or_rows
    return _write_sap_contract_history_sheet_impl(ws, rows, account)


__all__ = [
    "_write_sap_bb_events_sheet",
    "_write_sap_bb_matches_sheet",
    "_write_sap_header_row",
    "write_sap_back_billing_sheets",
    "write_sap_contract_history_sheet",
    "write_sap_financial_transactions_sheet",
    "write_sap_meter_readings_sheet",
]
