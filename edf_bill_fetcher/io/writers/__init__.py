"""Compat re-export - io.writers package.

Re-exports every public writer name from the ``io.writers`` submodules
so callers can use ``from edf_bill_fetcher.io.writers import ...``.

All re-exports are deferred to a PEP 562 module-level ``__getattr__``
(rather than ``from X import Y`` eager imports) to break a circular
import chain:

    caller (e.g. edf_collector.py)
      -> edf_bill_fetcher.writers
        -> edf_bill_fetcher.io.writers.evidence        (real extraction)
          -> edf_bill_fetcher.io.writers (this __init__)
            -> io.writers.back_billing (legacy shim)
              -> edf_bill_fetcher.writers (!cycle!)

Eager ``from edf_bill_fetcher.io.writers.back_billing import ...``
would re-enter ``writers/__init__.py`` mid-init. PEP 562 defers the
attribute resolution to first access, after both packages have
finished initialising.
"""

from __future__ import annotations

from typing import Any

__all__ = [
    # Phase 5A real extractions:
    "write_evidence_sheet",
    "write_summary_sheet",
    # Phase 5B real extractions:
    "write_statistical_analysis_sheet",
    "write_payment_analysis_sheet",
    "write_forecast_sheet",
    "write_data_quality_sheet",
    "write_tariff_analysis_sheet",
    # Legacy shims (still in writers/__init__.py):
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
]


def __getattr__(name: str) -> Any:
    if name in ("write_evidence_sheet", "write_summary_sheet"):
        from edf_bill_fetcher.io.writers.evidence import write_evidence_sheet, write_summary_sheet

        return {"write_evidence_sheet": write_evidence_sheet, "write_summary_sheet": write_summary_sheet}[name]
    if name == "write_statistical_analysis_sheet":
        from edf_bill_fetcher.io.writers.statistical import write_statistical_analysis_sheet

        return write_statistical_analysis_sheet
    if name == "write_payment_analysis_sheet":
        from edf_bill_fetcher.io.writers.payment import write_payment_analysis_sheet

        return write_payment_analysis_sheet
    if name == "write_forecast_sheet":
        from edf_bill_fetcher.io.writers.forecast import write_forecast_sheet

        return write_forecast_sheet
    if name == "write_data_quality_sheet":
        from edf_bill_fetcher.io.writers.data_quality import write_data_quality_sheet

        return write_data_quality_sheet
    if name == "write_tariff_analysis_sheet":
        from edf_bill_fetcher.io.writers.tariff import write_tariff_analysis_sheet

        return write_tariff_analysis_sheet
    if name in ("export_to_excel", "write_reconciliation_sheet"):
        from edf_bill_fetcher.io.writers.export import export_to_excel, write_reconciliation_sheet

        return {"export_to_excel": export_to_excel, "write_reconciliation_sheet": write_reconciliation_sheet}[name]
    if name == "write_back_billing_sheet":
        from edf_bill_fetcher.io.writers.back_billing import write_back_billing_sheet

        return write_back_billing_sheet
    if name == "write_rebilling_sheet":
        from edf_bill_fetcher.io.writers.rebilling import write_rebilling_sheet

        return write_rebilling_sheet
    if name in ("write_meter_readings_sheet", "write_contract_history_sheet"):
        import edf_bill_fetcher.io.writers.meter as m

        return getattr(m, name)
    if name in (
        "write_sap_back_billing_sheets",
        "write_sap_contract_history_sheet",
        "write_sap_financial_transactions_sheet",
        "write_sap_meter_readings_sheet",
    ):
        import edf_bill_fetcher.io.writers.sap as s

        return getattr(s, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
