"""Eager re-exports of the ``io.writers`` sheet writers.

Each name is imported from its canonical submodule (``evidence``,
``statistical``, ``payment``, ``forecast``, ``data_quality``, ``tariff``,
``export``, ``back_billing``, ``rebilling``, ``meter``, ``sap``) so
``from edf_bill_fetcher.io.writers import write_evidence_sheet`` works
without a lazy attribute-resolution indirection layer.
"""

from edf_bill_fetcher.io.writers.back_billing import write_back_billing_sheet
from edf_bill_fetcher.io.writers.data_quality import write_data_quality_sheet
from edf_bill_fetcher.io.writers.evidence import write_evidence_sheet, write_summary_sheet
from edf_bill_fetcher.io.writers.export import export_to_excel, write_reconciliation_sheet
from edf_bill_fetcher.io.writers.forecast import write_forecast_sheet
from edf_bill_fetcher.io.writers.meter import (
    write_contract_history_sheet,
    write_meter_readings_sheet,
)
from edf_bill_fetcher.io.writers.payment import write_payment_analysis_sheet
from edf_bill_fetcher.io.writers.rebilling import write_rebilling_sheet
from edf_bill_fetcher.io.writers.sap import (
    write_sap_back_billing_sheets,
    write_sap_contract_history_sheet,
    write_sap_financial_transactions_sheet,
    write_sap_meter_readings_sheet,
)
from edf_bill_fetcher.io.writers.statistical import write_statistical_analysis_sheet
from edf_bill_fetcher.io.writers.superseded import write_superseded_reconciliation_sheet
from edf_bill_fetcher.io.writers.tariff import write_tariff_analysis_sheet

try:
    import importlib.util

    HAS_SCIPY = importlib.util.find_spec("scipy") is not None
except ImportError:
    HAS_SCIPY = False

__all__ = [
    "export_to_excel",
    "write_back_billing_sheet",
    "write_contract_history_sheet",
    "write_data_quality_sheet",
    "write_evidence_sheet",
    "write_forecast_sheet",
    "write_meter_readings_sheet",
    "write_payment_analysis_sheet",
    "write_rebilling_sheet",
    "write_reconciliation_sheet",
    "write_sap_back_billing_sheets",
    "write_sap_contract_history_sheet",
    "write_sap_financial_transactions_sheet",
    "write_sap_meter_readings_sheet",
    "write_statistical_analysis_sheet",
    "write_summary_sheet",
    "write_superseded_reconciliation_sheet",
    "write_tariff_analysis_sheet",
    "HAS_SCIPY",
]
