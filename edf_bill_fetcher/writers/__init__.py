"""Shared writer helpers re-exported from ``writers._helpers``.

Legacy flat namespace for the ``_helpers`` module (extracted from
``edf_collector.py`` during the modularization).  The sheet-writing
functions moved to ``edf_bill_fetcher.io.writers`` — import them from
there.  This module exists only to keep ``from
edf_bill_fetcher.writers import <helper>`` call sites working.
"""

from edf_bill_fetcher.writers._helpers import (  # noqa: F401
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

__all__ = [
    "_SOURCE_PRECEDENCE",
    "DUP_GREY",
    "EDF_NAVY",
    "EDF_OFFWHITE",
    "EDF_ORANGE",
    "EST_YELLOW",
    "JUMP_RED",
    "MEDIUM_GREY",
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
    "_parse_amount_for_event",
    "_reading_type_to_aem",
    "_recon_hyperlink",
    "_zscore_anomalies",
    "build_evidence_index",
    "compute_dispute_flags",
    "detect_sap_back_billing_events",
    "match_sap_events_to_edf",
]
