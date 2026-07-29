#!/usr/bin/env python3
# ruff: noqa: I001
"""
EDF Master Evidence Collector
Collects billing data from PST/OST files, local PDF folders, and HTM account exports.
Fixed version: correct Excel date serials, dynamic range references, new PDF format support.

This module is now a thin compat re-export layer. The implementation has been
extracted into the ``edf_bill_fetcher`` package; this file re-exports the
public names so existing ``from edf_collector import X`` call sites continue
to work.
"""

from __future__ import annotations

import re
import pdfplumber  # noqa: F401  — test_engine_sources.py + test_multi_invoice_pdf_dispatch.py patch edf_collector.pdfplumber via unittest.mock.patch


# Tkinter is only needed for the GUI dialog.  Importing it at module
# level would crash on headless / CI machines that lack a display, so
# we guard it and set a flag that downstream GUI code checks.
try:
    import tkinter as tk  # noqa: F401
    from tkinter import filedialog, messagebox, ttk  # noqa: F401

    HAS_TK = True
except ImportError:
    HAS_TK = False

# ---------------------------------------------------------------------------
# Inline definitions — names that have not yet been extracted to
# ``edf_bill_fetcher`` submodules.  When they are extracted these
# blocks will be replaced with compat re-exports.
# ---------------------------------------------------------------------------

# Try to import pypff (PST parser) with graceful fallback
try:
    import pypff  # noqa: F401

    HAS_PYPFF = True
except ImportError:
    HAS_PYPFF = False

# Try to detect scipy for advanced stats (graceful fallback)
import importlib.util

try:
    HAS_SCIPY = importlib.util.find_spec("scipy") is not None
except ImportError:
    HAS_SCIPY = False

_SOURCE_PRECEDENCE: dict[str, int] = {
    "HTM Account History": 0,
    "Local PDF Folder": 1,
    "Statement Reconciliation": 1,
    "PST PDF Attachment": 2,
    "Email Body": 3,
    "Email Body (RTF)": 3,
}

AMOUNT_PATTERNS: list[tuple[str, re.Pattern[str]]] = [
    (
        "current_balance_debit",
        re.compile(r"current balance\s+£\s?([\d,]+(?:\.\d{2})?)\s*(?:in\s+)?debit", re.IGNORECASE),
    ),
    (
        "total_charges_period",
        re.compile(
            r"total charges for this period\s+£\s?([\d,]+(?:\.\d{2})?)\s*(?:in\s+)?debit",
            re.IGNORECASE,
        ),
    ),
    (
        "total_credits_bill",
        re.compile(
            r"total credits for this bill\s+£\s?([\d,]+(?:\.\d{2})?)(?:\s*(?:in\s+)?credit)?",
            re.IGNORECASE,
        ),
    ),
    (
        "total_charges_within",
        re.compile(r"total charges[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)", re.IGNORECASE),
    ),
    (
        "total_amount_due_within",
        re.compile(r"total amount due[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)", re.IGNORECASE),
    ),
    (
        "amount_to_pay_within",
        re.compile(r"amount to pay[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)", re.IGNORECASE),
    ),
    (
        "your_new_account_balance",
        re.compile(
            r"your new account balance\s+£\s?([\d,]+(?:\.\d{2})?)(?:\s*(?:in\s+)?(?:credit|debit))?",
            re.IGNORECASE,
        ),
    ),
    (
        "balance_within",
        re.compile(r"balance[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)", re.IGNORECASE),
    ),
    (
        "current_balance_within",
        re.compile(r"current balance[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)", re.IGNORECASE),
    ),
    (
        "pound_amount_debit",
        re.compile(r"£\s?([\d,]+(?:\.\d{2})?)\s*(?:in\s+)?debit", re.IGNORECASE),
    ),
    (
        "pound_amount_credit",
        re.compile(r"£\s?([\d,]+(?:\.\d{2})?)\s*credit", re.IGNORECASE),
    ),
]

_AMOUNT_PATTERN_NEW_BILL: frozenset[str] = frozenset(
    {
        "current_balance_debit",
        "total_charges_period",
        "total_credits_bill",
        "total_charges_within",
        "total_amount_due_within",
        "amount_to_pay_within",
        "pound_amount_debit",
    }
)
_AMOUNT_PATTERN_ONGOING_BALANCE: frozenset[str] = frozenset(
    {
        "your_new_account_balance",
        "balance_within",
        "current_balance_within",
        "pound_amount_credit",
    }
)
for _name, _ in AMOUNT_PATTERNS:
    assert _name in _AMOUNT_PATTERN_NEW_BILL or _name in _AMOUNT_PATTERN_ONGOING_BALANCE, (
        f"AMOUNT_PATTERNS entry {_name!r} has no entry-type bucket — "
        "add it to either _AMOUNT_PATTERN_NEW_BILL or _AMOUNT_PATTERN_ONGOING_BALANCE."
    )

READING_PATTERNS: dict[str, re.Pattern[str]] = {
    "Estimated": re.compile(r"estimated|est\.|estimate", re.IGNORECASE),
    "Smart": re.compile(r"smart meter|automated reading|smart reading", re.IGNORECASE),
    "Actual": re.compile(
        r"actual reading|customer reading|your reading|"
        r"reading was actual|reading is actual|"
        r"actual\s+reading\s*[-:]\s*\d|"
        r"meter\s+reading\s+was\s+actual",
        re.IGNORECASE,
    ),
}

# ---------------------------------------------------------------------------
# Compat re-exports — helpers
# ---------------------------------------------------------------------------
from edf_bill_fetcher.helpers.date_utils import (  # noqa: E402,F401,I001
    build_evidence_trail as _build_evidence_trail,
    completeness_score as _completeness_score,
    compute_ema as _compute_ema,
    compute_momentum as _compute_momentum,
    compute_rolling_stats as _compute_rolling_stats,
    _ISO_DATE_RE as _ISO_DATE_RE,
    _safe_to_datetime as _safe_to_datetime,
    parse_to_display_date as parse_to_display_date,
    parse_to_sort_date as parse_to_sort_date,
    to_excel_date as to_excel_date,
)
from edf_bill_fetcher.helpers.excel_utils import (  # noqa: E402,F401,I001
    _TEXT_SUPPRESSION_QUEUE,
    CELL_BORDER,
    build_sap_row_index_map as _build_sap_row_index_map,
    hcell as _hcell,
    money as _money,
    set_column_widths_from_spec as _set_column_widths_from_spec,
    suppress_text_warning as _suppress_text_warning,
    suppress_text_warnings_post_save as _suppress_text_warnings_post_save,
    text as _text,
)
from edf_bill_fetcher.helpers.formatting import (  # noqa: E402,F401,I001
    _is_populated as _is_populated,
    _amalgamate_cluster as _amalgamate_cluster,
    apply_currency_format as _apply_currency_format,
    apply_int_format as _apply_int_format,
)

# ---------------------------------------------------------------------------
# Compat re-exports — adapters (PDF / HTML / PST)
# ---------------------------------------------------------------------------
from edf_bill_fetcher.io.adapters.pdf import (  # noqa: E402,F401,I001
    ADMIT_RE,
    INV_BOUNDARY_RE,
    LEGAL_CONTEXT,
    PAGE1_BOUNDARY_RE,
    extract_admit_phrase,
    legal_context,
    slice_pdf_pages,
)
from edf_bill_fetcher.io.adapters.html import (  # noqa: E402,F401,I001
    htm_excerpt,
    parse_htm_account_history,
)
from edf_bill_fetcher.io.adapters.pst import (  # noqa: E402,F401,I001
    EMAIL_ADDR_RE,
    FROM_HEADER_RE,
    PST_PR_ATTACH_FILENAME,
    PST_PR_ATTACH_LONG_FILENAME,
    extract_sender_email,
    matches_domain_filter,
    pst_attachment_filename,
)

# Private alias for legacy _ADMIT_RE import
_ADMIT_RE = ADMIT_RE
_INV_BOUNDARY_RE = INV_BOUNDARY_RE

# ---------------------------------------------------------------------------
# Compat re-exports — UI classes
# ---------------------------------------------------------------------------
from edf_bill_fetcher.ui.app import App, ReportOptionsDialog  # noqa: E402,F401,I001

# ---------------------------------------------------------------------------
# Compat re-exports — CLI functions
# ---------------------------------------------------------------------------
from edf_bill_fetcher.io.cli import (  # noqa: E402,F401,I001
    _RestrictedUnpickler,
    _safe_pickle_load,
    main,
    run_cli_docx_report,
    run_cli_extract,
    run_cli_pdf_report,
)

# ---------------------------------------------------------------------------
# Compat re-exports — processors
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Compat re-exports — Models
# ---------------------------------------------------------------------------
from edf_bill_fetcher.models.events import SapBackBillingEvent  # noqa: E402,F401,I001

# ---------------------------------------------------------------------------
# Compat re-exports — EvidenceEngine & engine helpers
# ---------------------------------------------------------------------------
from edf_bill_fetcher.collectors.engine import (  # noqa: E402,F401,I001
    EvidenceEngine,
    account_number_matches as _account_number_matches,
    _extract_sender_email as _extract_sender_email,
    _fallback_amount as _fallback_amount,
    _fallback_inv_num as _fallback_inv_num,
    _fallback_period_from as _fallback_period_from,
    _fallback_period_to as _fallback_period_to,
    _matches_domain_filter as _matches_domain_filter,
    _pst_attachment_filename as _pst_attachment_filename,
)

# ---------------------------------------------------------------------------
# Compat re-exports — Writers shared helpers (theme colours)
# ---------------------------------------------------------------------------
from edf_bill_fetcher.writers._helpers import (  # noqa: E402,F401,I001
    DUP_GREY,
    EDF_NAVY,
    EDF_ORANGE,
    MEDIUM_GREY,
)

from edf_bill_fetcher.processors.detection import (  # noqa: E402,F401,I001
    detect_back_billing,
    detect_rebilling,
    detect_meter_rollover,
    detect_pdf_format,
    _assess_reason,
    _disclosed_label,
    _reversal_match,
    _DEFAULT_ROLLOVER_THRESHOLD,
    _KCR_PRESENCE_RE,
    _KI_PRESENCE_RE,
    _SAP_DDMMYYYY_RE,
    write_back_billing_sheet,
)
from edf_bill_fetcher.processors.matching import (  # noqa: E402,F401,I001
    infer_contracts,
    match_sap_events_to_edf,
    build_evidence_index,
    _confidence_band,
    _PST_PR_ATTACH_FILENAME,
    _PST_PR_ATTACH_LONG_FILENAME,
    _RECON_MONTH_MAP,
)
from edf_bill_fetcher.processors.analysis import (  # noqa: E402,F401,I001
    compute_dispute_flags,
    _reading_type_to_aem,
    _analyze_tariff_impact,
    _data_quality_report,
    _detect_payment_patterns,
)
from edf_bill_fetcher.processors.reconciliation import (  # noqa: E402,F401,I001
    _recon_parse_iso_date,
    _recon_amount_to_float,
    _recon_hyperlink,
)
from edf_bill_fetcher.processors.forecasting import (  # noqa: E402,F401,I001
    _compute_volatility,
    _zscore_anomalies,
    _iqr_anomalies,
    _linear_forecast,
    _linear_forecast_pair,
    _holt_winters_forecast,
    _holt_winters_forecast_pair,
)
from edf_bill_fetcher.processors.patterns import (  # noqa: E402,F401,I001
    _BILLING_PERIOD_RE,
    _INV_NUMBER_RE,
    _PERIOD_CHARGE_RE,
    _CREDIT_NUMBER_RE,
    _CREDIT_TOTAL_RE,
    _FROM_HEADER_RE,
    _COVER_BLOCK_INV_RE,
    _COVER_BLOCK_PERIOD_RE,
    _FALLBACK_AMOUNT_RE,
    _FALLBACK_INV_RE,
    _POUND_AMOUNT_FALLBACK_RE,
    _EMAIL_ADDR_RE,
    PERIOD_RE,
)
from edf_bill_fetcher.processors.sap_parsers import (  # noqa: E402,F401,I001
    _ACC_NUM_RE,
    _CURRENT_BAL_RE,
    _DATE_ISSUED_RE,
    _STANDING_CHARGE_RE,
    _TARIFF_NAME_RE,
    _UNITS_USED_RE,
    detect_reconciliation_statement,
    detect_sap_dump,
    extract_new_credit_fields,
    extract_new_invoice_fields,
    extract_reconciliation_statement_rows,
    parse_sap_contract_history,
    parse_sap_financial_transactions,
    parse_sap_meter_read_history,
)
from edf_bill_fetcher.collectors.engine import (  # noqa: E402,F401,I001
    _ACCOUNT_BALANCE_LANG_RE,
    _BILL_INDICATORS_RE,
    _BILL_MARKERS_RE,
    _OLD_PDF_DATE_RE,
    _OLD_PDF_INV_RE,
    _OLD_PDF_KWH_RE,
    _OLD_PDF_PERIOD_CHARGE_RE,
    _OLD_PDF_STANDING_RE,
)

# ---------------------------------------------------------------------------
# Compat re-exports — writers
# ---------------------------------------------------------------------------
from edf_bill_fetcher.writers import (  # noqa: E402,F401,I001
    _bb_invoice_value,
    _write_sap_bb_events_sheet,
    _write_sap_bb_matches_sheet,
    _write_sap_header_row,
    export_to_excel,
    run_analysers,
    write_contract_history_sheet,
    write_data_quality_sheet,
    write_evidence_sheet,
    write_forecast_sheet,
    write_meter_readings_sheet,
    write_payment_analysis_sheet,
    write_rebilling_sheet,
    write_reconciliation_sheet,
    write_sap_back_billing_sheets,
    write_sap_contract_history_sheet,
    write_sap_financial_transactions_sheet,
    write_sap_meter_readings_sheet,
    write_statistical_analysis_sheet,
    write_summary_sheet,
    write_tariff_analysis_sheet,
    EST_YELLOW,
    JUMP_RED,
    detect_sap_back_billing_events,
)

# ---------------------------------------------------------------------------
# __all__
# ---------------------------------------------------------------------------
__all__ = [
    # helpers
    "_build_evidence_trail",
    "_completeness_score",
    "_compute_ema",
    "_compute_momentum",
    "_compute_rolling_stats",
    "_TEXT_SUPPRESSION_QUEUE",
    "CELL_BORDER",
    "_build_sap_row_index_map",
    "_hcell",
    "_money",
    # adapters
    "ADMIT_RE",
    "INV_BOUNDARY_RE",
    "LEGAL_CONTEXT",
    "PAGE1_BOUNDARY_RE",
    "extract_admit_phrase",
    "legal_context",
    "slice_pdf_pages",
    "htm_excerpt",
    "parse_htm_account_history",
    "EMAIL_ADDR_RE",
    "FROM_HEADER_RE",
    "PST_PR_ATTACH_FILENAME",
    "PST_PR_ATTACH_LONG_FILENAME",
    "extract_sender_email",
    "matches_domain_filter",
    "pst_attachment_filename",
    # UI
    "App",
    "ReportOptionsDialog",
    # CLI
    "run_cli_extract",
    "run_cli_pdf_report",
    "run_cli_docx_report",
    "main",
    "_safe_pickle_load",
    "_RestrictedUnpickler",
    # EvidenceEngine
    "EvidenceEngine",
    # processors / detection
    "detect_back_billing",
    "detect_rebilling",
    "detect_meter_rollover",
    "detect_pdf_format",
    "_assess_reason",
    "_disclosed_label",
    "_reversal_match",
    "_DEFAULT_ROLLOVER_THRESHOLD",
    "_KCR_PRESENCE_RE",
    "_KI_PRESENCE_RE",
    "_SAP_DDMMYYYY_RE",
    "write_back_billing_sheet",
    # processors / matching
    "infer_contracts",
    "match_sap_events_to_edf",
    "build_evidence_index",
    "_confidence_band",
    "_PST_PR_ATTACH_FILENAME",
    "_PST_PR_ATTACH_LONG_FILENAME",
    "_RECON_MONTH_MAP",
    # processors / analysis
    "compute_dispute_flags",
    "_reading_type_to_aem",
    "_analyze_tariff_impact",
    "_data_quality_report",
    "_detect_payment_patterns",
    # processors / reconciliation
    "_recon_parse_iso_date",
    "_recon_amount_to_float",
    "_recon_hyperlink",
    "write_reconciliation_sheet",
    # processors / forecasting
    "_compute_volatility",
    "_zscore_anomalies",
    "_iqr_anomalies",
    "_linear_forecast",
    "_linear_forecast_pair",
    "_holt_winters_forecast",
    "_holt_winters_forecast_pair",
    # processors / patterns
    "_BILLING_PERIOD_RE",
    "_INV_NUMBER_RE",
    "_PERIOD_CHARGE_RE",
    "_CREDIT_NUMBER_RE",
    "_CREDIT_TOTAL_RE",
    "_FROM_HEADER_RE",
    "_COVER_BLOCK_INV_RE",
    "_COVER_BLOCK_PERIOD_RE",
    "_FALLBACK_AMOUNT_RE",
    "_FALLBACK_INV_RE",
    "_POUND_AMOUNT_FALLBACK_RE",
    "_EMAIL_ADDR_RE",
    "PERIOD_RE",
    # SAP parsers
    "_ACC_NUM_RE",
    "_CURRENT_BAL_RE",
    "_DATE_ISSUED_RE",
    "_STANDING_CHARGE_RE",
    "_TARIFF_NAME_RE",
    "_UNITS_USED_RE",
    # collectors / engine
    "_ACCOUNT_BALANCE_LANG_RE",
    "_BILL_INDICATORS_RE",
    "_BILL_MARKERS_RE",
    "_OLD_PDF_DATE_RE",
    "_OLD_PDF_INV_RE",
    "_OLD_PDF_KWH_RE",
    "_OLD_PDF_PERIOD_CHARGE_RE",
    "_OLD_PDF_STANDING_RE",
    # inline defs (not yet extracted)
    "AMOUNT_PATTERNS",
    "HAS_PYPFF",
    "HAS_SCIPY",
    "READING_PATTERNS",
    "_AMOUNT_PATTERN_NEW_BILL",
    "_AMOUNT_PATTERN_ONGOING_BALANCE",
    "_SOURCE_PRECEDENCE",
    # models
    "SapBackBillingEvent",
    # EvidenceEngine
    "EvidenceEngine",
    # collectors / engine helpers
    "_account_number_matches",
    "_extract_sender_email",
    "_fallback_amount",
    "_fallback_inv_num",
    "_fallback_period_from",
    "_fallback_period_to",
    "_matches_domain_filter",
    "_pst_attachment_filename",
    # helpers
    "_build_evidence_trail",
    "_completeness_score",
    "_compute_ema",
    "_compute_momentum",
    "_compute_rolling_stats",
    "_TEXT_SUPPRESSION_QUEUE",
    "CELL_BORDER",
    "_build_sap_row_index_map",
    "_hcell",
    "_money",
    "_set_column_widths_from_spec",
    "_suppress_text_warning",
    "_suppress_text_warnings_post_save",
    "_text",
    "_apply_currency_format",
    "_apply_int_format",
    "_is_populated",
    "_amalgamate_cluster",
    "_ISO_DATE_RE",
    "_safe_to_datetime",
    "parse_to_display_date",
    "parse_to_sort_date",
    "to_excel_date",
    # theme
    "DUP_GREY",
    "EDF_NAVY",
    "EDF_ORANGE",
    "MEDIUM_GREY",
    # adapters
    "ADMIT_RE",
    "INV_BOUNDARY_RE",
    "LEGAL_CONTEXT",
    "PAGE1_BOUNDARY_RE",
    "extract_admit_phrase",
    "legal_context",
    "slice_pdf_pages",
    "htm_excerpt",
    "parse_htm_account_history",
    "EMAIL_ADDR_RE",
    "FROM_HEADER_RE",
    "PST_PR_ATTACH_FILENAME",
    "PST_PR_ATTACH_LONG_FILENAME",
    "extract_sender_email",
    "matches_domain_filter",
    "pst_attachment_filename",
    # UI
    "App",
    "ReportOptionsDialog",
    # CLI
    "run_cli_extract",
    "run_cli_pdf_report",
    "run_cli_docx_report",
    "main",
    "_safe_pickle_load",
    "_RestrictedUnpickler",
    # processors / detection
    "detect_back_billing",
    "detect_rebilling",
    "detect_meter_rollover",
    "detect_pdf_format",
    "_assess_reason",
    "_disclosed_label",
    "_reversal_match",
    "_DEFAULT_ROLLOVER_THRESHOLD",
    "_KCR_PRESENCE_RE",
    "_KI_PRESENCE_RE",
    "_SAP_DDMMYYYY_RE",
    "write_back_billing_sheet",
    # processors / matching
    "infer_contracts",
    "match_sap_events_to_edf",
    "build_evidence_index",
    "_confidence_band",
    "_PST_PR_ATTACH_FILENAME",
    "_PST_PR_ATTACH_LONG_FILENAME",
    "_RECON_MONTH_MAP",
    # processors / analysis
    "compute_dispute_flags",
    "_reading_type_to_aem",
    "_analyze_tariff_impact",
    "_data_quality_report",
    "_detect_payment_patterns",
    # processors / reconciliation
    "_recon_parse_iso_date",
    "_recon_amount_to_float",
    "_recon_hyperlink",
    "write_reconciliation_sheet",
    # processors / forecasting
    "_compute_volatility",
    "_zscore_anomalies",
    "_iqr_anomalies",
    "_linear_forecast",
    "_linear_forecast_pair",
    "_holt_winters_forecast",
    "_holt_winters_forecast_pair",
    # processors / patterns
    "_BILLING_PERIOD_RE",
    "_INV_NUMBER_RE",
    "_PERIOD_CHARGE_RE",
    "_CREDIT_NUMBER_RE",
    "_CREDIT_TOTAL_RE",
    "_FROM_HEADER_RE",
    "_COVER_BLOCK_INV_RE",
    "_COVER_BLOCK_PERIOD_RE",
    "_FALLBACK_AMOUNT_RE",
    "_FALLBACK_INV_RE",
    "_POUND_AMOUNT_FALLBACK_RE",
    "_EMAIL_ADDR_RE",
    "PERIOD_RE",
    # SAP parsers
    "_ACC_NUM_RE",
    "_CURRENT_BAL_RE",
    "_DATE_ISSUED_RE",
    "_STANDING_CHARGE_RE",
    "_TARIFF_NAME_RE",
    "_UNITS_USED_RE",
    "detect_reconciliation_statement",
    "detect_sap_dump",
    "extract_new_credit_fields",
    "extract_new_invoice_fields",
    "extract_reconciliation_statement_rows",
    "parse_sap_contract_history",
    "parse_sap_financial_transactions",
    "parse_sap_meter_read_history",
    # collectors / engine
    "_ACCOUNT_BALANCE_LANG_RE",
    "_BILL_INDICATORS_RE",
    "_BILL_MARKERS_RE",
    "_OLD_PDF_DATE_RE",
    "_OLD_PDF_INV_RE",
    "_OLD_PDF_KWH_RE",
    "_OLD_PDF_PERIOD_CHARGE_RE",
    "_OLD_PDF_STANDING_RE",
    # writers
    "_bb_invoice_value",
    "_write_sap_bb_events_sheet",
    "_write_sap_bb_matches_sheet",
    "_write_sap_header_row",
    "export_to_excel",
    "run_analysers",
    "write_contract_history_sheet",
    "write_data_quality_sheet",
    "write_evidence_sheet",
    "write_forecast_sheet",
    "write_meter_readings_sheet",
    "write_payment_analysis_sheet",
    "write_rebilling_sheet",
    "write_sap_back_billing_sheets",
    "write_sap_contract_history_sheet",
    "write_sap_financial_transactions_sheet",
    "write_sap_meter_readings_sheet",
    "write_statistical_analysis_sheet",
    "write_summary_sheet",
    "write_tariff_analysis_sheet",
    "EST_YELLOW",
    "JUMP_RED",
    "detect_sap_back_billing_events",
]
