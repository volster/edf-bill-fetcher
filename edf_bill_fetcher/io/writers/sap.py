"""Compat re-export - SAP sheet writer (with adapter).

Implementation lives in ``edf_bill_fetcher.writers`` (the legacy monolith
at ``writers/__init__.py``). Real extraction happens in later phases
when the monolith is deleted.

The public API exposed here matches the test contract at
``tests/test_io_writers_extraction.py``: ``write_sap_contract_history_sheet``
takes ``(ws, df)`` only; the underlying implementation expects
``(ws, rows: list[dict], account: str)``. The adapter converts the
DataFrame to list-of-dicts and supplies an empty default account.

The import from ``edf_bill_fetcher.writers`` is deferred into the
adapter function bodies and a PEP 562 ``__getattr__`` to break a
circular import: ``writers/__init__.py`` imports from
``io.writers.evidence`` (real extraction) which loads
``io/writers/__init__.py`` which in turn imports every shim in this
package. Direct ``from edf_bill_fetcher.writers import ...`` at
module-init time would re-enter ``writers/__init__.py`` mid-init and
raise ``ImportError: partially initialized module``. Deferring lets
Python finish loading the package before the back-reference is
resolved on first access.
"""

from __future__ import annotations

from typing import Any

import pandas as pd

__all__ = [
    "_write_sap_bb_events_sheet",
    "_write_sap_bb_matches_sheet",
    "_write_sap_header_row",
    "write_sap_back_billing_sheets",
    "write_sap_contract_history_sheet",
    "write_sap_financial_transactions_sheet",
    "write_sap_meter_readings_sheet",
]


def write_sap_contract_history_sheet(
    ws,
    df_or_rows: pd.DataFrame | list[dict[str, Any]],
    account: str = "",
) -> None:
    """Adapter: test contract uses ``(ws, df)``; convert DataFrame to rows."""
    from edf_bill_fetcher.writers import write_sap_contract_history_sheet as _impl

    if isinstance(df_or_rows, pd.DataFrame):
        rows = df_or_rows.to_dict(orient="records")
    else:
        rows = df_or_rows
    return _impl(ws, rows, account)


def __getattr__(name: str):
    if name in __all__ and name != "write_sap_contract_history_sheet":
        # Lazy-import from the monolith on first attribute access.
        from edf_bill_fetcher import writers as _w

        return getattr(_w, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
