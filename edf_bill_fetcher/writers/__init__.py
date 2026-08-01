"""Writer functions for the EDF evidence workbook.

Extracted from ``edf_collector.py`` as part of the modularization
refactor (Task 5).  Each function writes one or more Excel sheets
using openpyxl.
"""

from __future__ import annotations

try:
    import tkinter as tk  # noqa: F401

    HAS_TK = True
except ImportError:
    HAS_TK = False

try:
    import pypff  # noqa: F401

    HAS_PYPFF = True
except ImportError:
    HAS_PYPFF = False

try:
    import importlib.util

    HAS_SCIPY = importlib.util.find_spec("scipy") is not None
except ImportError:
    HAS_SCIPY = False

try:
    importlib.util.find_spec("statsmodels.tsa.holtwinters")

    HAS_STATSMODELS = True
except ImportError:
    HAS_STATSMODELS = False

try:
    HAS_PDF_REPORT = importlib.util.find_spec("edf_report") is not None
    HAS_DOCX_REPORT = importlib.util.find_spec("edf_report_docx") is not None
except ImportError:
    HAS_PDF_REPORT = False
    HAS_DOCX_REPORT = False
from edf_bill_fetcher.helpers.excel_utils import (  # noqa: E402,F401,I001
    _TEXT_SUPPRESSION_QUEUE,
    CELL_BORDER,
)
from edf_bill_fetcher.helpers.excel_utils import (  # noqa: F401
    open_pdf_hyperlink_cell as _open_pdf_hyperlink_cell,
)
from edf_bill_fetcher.helpers.formatting import (  # noqa: E402,F401,I001
    account_number_matches as _account_number_matches,
)
from edf_bill_fetcher.io.adapters.pdf import legal_context  # noqa: E402,F401,I001
from edf_bill_fetcher.io.writers.evidence import (  # noqa: E402,F401
    write_evidence_sheet,
    write_summary_sheet,
)
from edf_bill_fetcher.writers._helpers import (  # noqa: E402,F401,I001
    _SOURCE_PRECEDENCE,
    DUP_GREY,
    EDF_NAVY,
    EDF_OFFWHITE,
    EDF_ORANGE,
    EST_YELLOW,
    JUMP_RED,
    MEDIUM_GREY,
    _analyze_tariff_impact,
    _compute_volatility,
    _data_quality_report,
    _detect_payment_patterns,
    _disclosed_label,
    _holt_winters_forecast,
    _holt_winters_forecast_pair,
    _iqr_anomalies,
    _linear_forecast,
    _linear_forecast_pair,
    _parse_amount_for_event,
    _reading_type_to_aem,
    _recon_hyperlink,
    _zscore_anomalies,
    build_evidence_index,
    compute_dispute_flags,
    detect_sap_back_billing_events,
    match_sap_events_to_edf,
)
from evidence_bundle import build_bundle_index, save_evidence_files  # noqa: E402,F401

__all__ = [
    "_analyze_tariff_impact",
    "_compute_volatility",
    "_data_quality_report",
    "_detect_payment_patterns",
    "_disclosed_label",
    "_holt_winters_forecast",
    "_holt_winters_forecast_pair",
    "_iqr_anomalies",
    "_linear_forecast",
    "_linear_forecast_pair",
    "_recon_hyperlink",
    "_write_sap_bb_events_sheet",
    "_write_sap_bb_matches_sheet",
    "_write_sap_header_row",
    "build_evidence_index",
    "compute_dispute_flags",
    "detect_back_billing",
    "detect_sap_back_billing_events",
    "export_to_excel",
    "match_sap_events_to_edf",
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
    "write_tariff_analysis_sheet",
]


from edf_bill_fetcher.io.writers.sap import (  # noqa: E402,F401,I001
    _bb_invoice_value,
    _write_sap_bb_events_sheet,
    _write_sap_bb_matches_sheet,
    _write_sap_header_row,
    _write_sap_contract_history_sheet_impl,
    write_sap_back_billing_sheets,
    write_sap_contract_history_sheet,
    write_sap_financial_transactions_sheet,
    write_sap_meter_readings_sheet,
)

def __getattr__(name: str) -> object:
    if name == "run_analysers":
        from edf_bill_fetcher.io.writers.analysis import run_analysers

        return run_analysers
    if name in ("write_evidence_sheet", "write_summary_sheet"):
        from edf_bill_fetcher.io.writers.evidence import (
            write_evidence_sheet,
            write_summary_sheet,
        )

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
    if name == "export_to_excel":
        from edf_bill_fetcher.io.writers.export import export_to_excel

        return export_to_excel
    if name == "write_reconciliation_sheet":
        from edf_bill_fetcher.io.writers.reconciliation import write_reconciliation_sheet

        return write_reconciliation_sheet
    if name in (
        "write_back_billing_sheet",
        "_assess_reason",
        "detect_back_billing",
    ):
        import edf_bill_fetcher.io.writers.back_billing as _m_back_billing

        return getattr(_m_back_billing, name)
    if name in (
        "write_rebilling_sheet",
        "_reversal_match",
        "detect_rebilling",
    ):
        import edf_bill_fetcher.io.writers.rebilling as _m_rebilling

        return getattr(_m_rebilling, name)
    if name in (
        "detect_meter_rollover",
        "infer_contracts",
        "write_meter_readings_sheet",
        "write_contract_history_sheet",
    ):
        import edf_bill_fetcher.io.writers.meter as _m_meter

        return getattr(_m_meter, name)
    if name in (
        "write_sap_back_billing_sheets",
        "_bb_invoice_value",
        "_write_sap_bb_events_sheet",
        "_write_sap_bb_matches_sheet",
        "_write_sap_header_row",
        "_write_sap_contract_history_sheet_impl",
        "write_sap_contract_history_sheet",
        "write_sap_financial_transactions_sheet",
        "write_sap_meter_readings_sheet",
    ):
        import edf_bill_fetcher.io.writers.sap as _m_sap

        return getattr(_m_sap, name)
    if name in (
        "_recon_amount_to_float",
        "_recon_parse_iso_date",
    ):
        import edf_bill_fetcher.io.writers.reconciliation as _m_recon

        return getattr(_m_recon, name)
    if name in (
        "detect_sap_back_billing_events",
        "match_sap_events_to_edf",
        "compute_dispute_flags",
        "build_evidence_index",
        "_SOURCE_PRECEDENCE",
    ):
        import edf_bill_fetcher.writers._helpers as _m_helpers

        return getattr(_m_helpers, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")

