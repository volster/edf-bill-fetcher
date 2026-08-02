"""Coverage tests for the io.writers PEP 562 lazy-shim package.

Closes the 34-missed-line gap in
``edf_bill_fetcher/io/writers/__init__.py`` by exercising both the
per-name ``__getattr__`` lookup branch (one test per ``__all__``
entry) and the ``raise AttributeError`` fallback branch (one
negative test).
"""

from __future__ import annotations

import pytest

# Pre-import the legacy ``edf_bill_fetcher.writers`` package so that
# the circular import chain documented in
# ``edf_bill_fetcher/io/writers/__init__.py`` is fully resolved before
# the shim's ``__getattr__`` triggers
# ``from edf_bill_fetcher.io.writers.evidence import ...``. Without
# this, ``write_evidence_sheet`` / ``write_summary_sheet`` raise
# ``ImportError: cannot import name 'write_evidence_sheet' from
# partially initialized module 'edf_bill_fetcher.io.writers.evidence'``
# because ``edf_bill_fetcher.writers`` is mid-initialisation when the
# shim re-enters it.
import edf_bill_fetcher.io.writers as pkg  # noqa: F401,I001  (side-effect import)
import edf_bill_fetcher.writers  # noqa: F401,I001  (side-effect import)

SHIM_NAMES: tuple[str, ...] = (
    "write_evidence_sheet",
    "write_summary_sheet",
    "write_statistical_analysis_sheet",
    "write_payment_analysis_sheet",
    "write_forecast_sheet",
    "write_data_quality_sheet",
    "write_tariff_analysis_sheet",
    "export_to_excel",
    "write_reconciliation_sheet",
    "write_back_billing_sheet",
    "write_rebilling_sheet",
    "write_meter_readings_sheet",
    "write_contract_history_sheet",
    "write_sap_back_billing_sheets",
    "write_sap_contract_history_sheet",
    "write_sap_financial_transactions_sheet",
    "write_sap_meter_readings_sheet",
)


@pytest.mark.parametrize("name", SHIM_NAMES)
def test_io_writers_shim_resolves_name(name: str) -> None:
    """The io.writers shim resolves every ``__all__`` name via PEP 562 ``__getattr__``."""
    resolved = getattr(pkg, name)
    assert callable(resolved), f"{name} did not resolve to a callable"


def test_io_writers_shim_raises_for_unknown_name() -> None:
    """The io.writers shim raises ``AttributeError`` with a helpful message for an unknown name."""
    unknown_name = "nonexistent_writers_function_xyz"
    with pytest.raises(AttributeError):
        getattr(pkg, unknown_name)
