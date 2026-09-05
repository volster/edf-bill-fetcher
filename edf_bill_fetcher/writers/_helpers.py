"""Writer-specific helpers and constants extracted from edf_collector.py."""

from __future__ import annotations

import openpyxl
import pandas as pd
from openpyxl.styles import Font

try:
    from statsmodels.tsa.holtwinters import ExponentialSmoothing  # noqa: F401

    HAS_STATSMODELS = True
except ImportError:
    HAS_STATSMODELS = False

from edf_bill_fetcher.helpers.excel_utils import build_evidence_index  # noqa: F401
from edf_bill_fetcher.helpers.formatting import parse_amount
from edf_bill_fetcher.helpers.theme import EDF_NAVY, EDF_OFFWHITE, EDF_ORANGE  # noqa: F401
from edf_bill_fetcher.models.events import SapBackBillingEvent

# Re-export the canonical SAP<->EDF matcher from processors.matching so
# the two copies are unified into a single implementation.  Existing
# importers (writers/__init__.py, io/writers/export.py) keep working.
from edf_bill_fetcher.processors.matching import match_sap_events_to_edf  # noqa: F401

EST_YELLOW = "FFFF99"
JUMP_RED = "FF9999"
DUP_GREY = "E0E0E0"
MEDIUM_GREY = "#666666"

# Source precedence for deduplication in export_to_excel.
# Lower number = higher precedence.
_SOURCE_PRECEDENCE: dict[str, int] = {
    "HTM Account History": 0,
    "Local PDF Folder": 1,
    "Statement Reconciliation": 1,
    "PST PDF Attachment": 2,
    "Email Body": 3,
    "Email Body (RTF)": 3,
}


def _analyze_tariff_impact(df: pd.DataFrame) -> dict[str, object]:
    """Alias for the shared computed model (see models/report_models.py)."""
    from edf_bill_fetcher.models.report_models import compute_tariff_analysis

    result = compute_tariff_analysis(df)
    if result.empty:
        return {}
    return {
        "tariff_stats": result.stats,
        "num_tariffs": result.num_tariffs,
        "tariff_changes": result.tariff_changes,
    }


def _data_quality_report(df: pd.DataFrame) -> dict[str, object]:
    """Alias for the shared computed model (see models/report_models.py)."""
    from edf_bill_fetcher.models.report_models import compute_data_quality_report

    return compute_data_quality_report(df).to_dict()


def _disclosed_label(admitted: bool, overlaps: bool) -> str:
    """Return the human-readable value of the 'Cancel/Rebill Disclosed' cell used on the Back-billing and Rebilling tabs.

    The disclosed column joins two independent signals:
      * admit-phrase (the cover-page wording 'we've recently
        cancelled some charges for you'), captured as a bool on the
        record; and
      * period overlap, flagged by :func:`detect_rebilling`.
    """
    if admitted and overlaps:
        return "Admitted + overlap"
    if admitted:
        return "Admitted phrase"
    if overlaps:
        return "Period overlap"
    return ""


def _reading_type_to_aem(reading_value: str) -> str:
    """Map the Reading column's value (Actual/Estimated/Smart/Unknown) to the single-letter A/E/M code used on the Meter Readings tab."""
    if reading_value == "Actual":
        return "A"
    if reading_value == "Estimated":
        return "E"
    if reading_value == "Smart":
        return "A"
    return "E"


def _recon_hyperlink(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    row: int,
    col: int,
    sheet: str,
    target_row: int,
) -> None:
    cell = ws.cell(row=row, column=col)
    location = f"'{sheet}'!A{target_row}"
    cell.hyperlink = openpyxl.worksheet.hyperlink.Hyperlink(
        ref=cell.coordinate,
        location=location,
        display="\u2192",
        tooltip=f"Jump to {sheet}!A{target_row}",
    )
    cell.value = "\u2192"
    cell.font = Font(name="Calibri", size=10, color="0563C1", underline="single")


# The forecast/anomaly primitives below are aliases to the canonical
# ``processors.forecasting`` implementations (removing the old backwards
# ``processors -> writers`` dependency). Re-exporting them here preserves every
# legacy ``from edf_bill_fetcher.writers._helpers import ...`` call site.
from edf_bill_fetcher.processors.forecasting import (  # noqa: E402,F401,I001
    _compute_volatility,
    _holt_winters_forecast,
    _holt_winters_forecast_pair,
    _iqr_anomalies,
    _linear_forecast,
    _linear_forecast_pair,
    _zscore_anomalies,
)


def _detect_payment_patterns(df):
    """Compat alias for the shared compute (see models/report_models.py)."""
    from edf_bill_fetcher.models.report_models import compute_payment_analysis

    pa = compute_payment_analysis(df)
    if pa.count == 0:
        return {}
    return {
        "count": pa.count,
        "total_paid": pa.total_paid,
        "avg_payment": pa.avg_payment,
        "median_payment": pa.median_payment,
        "max_payment": pa.largest_payment,
        "min_payment": pa.smallest_payment,
        "avg_interval_days": pa.avg_interval_days,
        "median_interval_days": pa.median_interval_days,
        "last_payment_date": pa.last_payment_date,
        "last_payment_amount": pa.last_payment_amount,
    }


# SAP back-billing constants
_SAP_DEBT_MGMT_FLAG_VALUE = "Installment Plan Item"
_SAP_MIN_CLUSTER_SIZE = 4
_SAP_MATCH_DAY_BANDS = ((0, 50), (3, 25), (14, 5))
_SAP_MATCH_AMOUNT_BANDS = ((0.05, 40), (0.25, 20), (0.50, 5))
_SAP_CONFIDENCE_BANDS = (("High", 75), ("Medium", 40), ("Low", 10))


_parse_amount_for_event = parse_amount


def _confidence_band(score: int) -> str | None:
    """Map a numeric match score to High/Medium/Low/None (Unmatched)."""
    for band, threshold in _SAP_CONFIDENCE_BANDS:
        if score >= threshold:
            return band
    return None


def detect_sap_back_billing_events(
    sap_rows: list[dict],
    *,
    min_cluster_size: int = _SAP_MIN_CLUSTER_SIZE,
) -> list:
    """Cluster SAP Financial Transaction rows into back-billing events."""
    from collections import Counter

    from edf_bill_fetcher.helpers.date_utils import build_evidence_trail
    from edf_bill_fetcher.models.events import SapBackBillingEvent

    if not sap_rows:
        return []

    filtered = [
        r
        for r in sap_rows
        if str(r.get("Statistical Key Flag", "")).strip() != _SAP_DEBT_MGMT_FLAG_VALUE
    ]

    clusters: dict[str, list[dict]] = {}
    for r in filtered:
        cd = str(r.get("Clearing Document", "")).strip()
        if not cd or cd in ("NA", "None", "*"):
            continue
        clusters.setdefault(cd, []).append(r)

    clusters_kept = {cd: rows for cd, rows in clusters.items() if len(rows) >= min_cluster_size}

    events: list = []
    for cd, rows in clusters_kept.items():
        clear_dates = [
            pd.to_datetime(str(r.get("Clearing Date", "")).strip(), errors="coerce")
            for r in rows
            if str(r.get("Clearing Date", "")).strip()
            and str(r.get("Clearing Date", "")).strip() not in ("NA", "None")
        ]
        clear_dates = [d for d in clear_dates if not pd.isna(d)]
        clearing_date = min(clear_dates) if clear_dates else pd.NaT

        reasons = [
            str(r.get("Clearing Reason", "")).strip()
            for r in rows
            if str(r.get("Clearing Reason", "")).strip()
        ]
        clearing_reason = Counter(reasons).most_common(1)[0][0] if reasons else ""

        net_amount = sum(_parse_amount_for_event(r.get("Amount")) for r in rows)
        has_credit = any(
            "Credit for Consum Billing" in str(r.get("Transaction Text", "")) for r in rows
        )
        has_acct_maint = any(
            str(r.get("Transaction Text", "")).strip() == "Account maintenance" for r in rows
        )
        amounts = [_parse_amount_for_event(r.get("Amount")) for r in rows]
        non_zero = [a for a in amounts if abs(a) > 0.001]
        largest = max(non_zero, key=lambda x: abs(x)) if non_zero else 0.0

        post_dates = [
            str(r.get("Posting Date", "")).strip()
            for r in rows
            if str(r.get("Posting Date", "")).strip()
            and str(r.get("Posting Date", "")).strip() not in ("NA", "None")
        ]
        posting_date_range = (min(post_dates), max(post_dates)) if post_dates else ("", "")
        evidence_trail = build_evidence_trail(rows)

        events.append(
            SapBackBillingEvent(
                clearing_doc=cd,
                clearing_date=clearing_date,
                clearing_reason=clearing_reason,
                rows=rows,
                net_amount=round(net_amount, 2),
                has_credit_for_consum_billing=has_credit,
                has_account_maintenance=has_acct_maint,
                largest_single_posting=round(largest, 2),
                posting_date_range=posting_date_range,
                evidence_trail=evidence_trail,
            )
        )

    events_sorted = sorted(
        events,
        key=lambda ev: (
            ev.clearing_date if not pd.isna(ev.clearing_date) else pd.Timestamp.max,
            ev.clearing_doc,
        ),
    )
    return events_sorted


def handle_cluster_unmatched(
    sap_event: SapBackBillingEvent,
    clusters: list[dict],
) -> dict | None:
    """Tag a SAP event as an internal mechanism of a back-billing cluster.

    When a SAP back-billing event's Posting Date falls inside a known
    back-billing cluster's posting-date window but no invoice in that
    cluster achieves amount-band agreement with the event's net amount,
    return a match dict tagging the event as an internal mechanism of
    that cluster.  Returns ``None`` when the event's posting-date range
    is empty, when no cluster window contains the posting date, or when
    an in-cluster invoice matches on the amount band (within 50%).

    The amount-band agreement test mirrors the spec §3.3 matcher's
    outer band: a SAP event and an EDF invoice agree when their amounts
    are within 50% of each other.  ``sap_event.net_amount`` is compared
    against each cluster invoice's ``Period Charge (£)``.

    Args:
        sap_event: A ``SapBackBillingEvent`` dataclass instance.  The
            event's ``posting_date_range`` (a ``(start, end)`` tuple of
            ISO date strings) supplies the posting date; the first
            non-empty bound is used as the comparison date.
        clusters: A list of cluster dicts, each with keys ``name``,
            ``posting_date_start``, ``posting_date_end``, and
            ``invoices`` (a list of dicts with ``Invoice #`` and
            ``Period Charge (£)``).

    Returns:
        A match dict with keys ``Matched EDF Invoice #``, ``Confidence``,
        ``Notes``, and ``Evidence Trail``; or ``None`` when the event
        should not be tagged as a cluster-unmatched internal mechanism.

    """
    posting_start, posting_end = sap_event.posting_date_range
    posting_date = posting_start or posting_end
    if not posting_date:
        return None

    if sap_event.net_amount == 0 and sap_event.largest_single_posting is not None:
        comparison_amount = abs(sap_event.largest_single_posting)
    else:
        comparison_amount = sap_event.net_amount

    for cluster in clusters:
        cluster_start = cluster.get("posting_date_start", "")
        cluster_end = cluster.get("posting_date_end", "")
        if not cluster_start or not cluster_end:
            continue
        if not (cluster_start <= posting_date <= cluster_end):
            continue

        # Posting Date is inside this cluster's window.  Check whether any
        # in-cluster invoice achieves amount-band agreement (within 50%,
        # mirroring the spec §3.3 matcher's outer band).
        for invoice in cluster.get("invoices", []):
            inv_amount = float(invoice.get("Period Charge (£)", 0) or 0)
            if inv_amount <= 0:
                continue
            if abs(comparison_amount - inv_amount) / inv_amount <= 0.50:
                # Amount agreement exists → not cluster-unmatched.
                return None

        # No amount agreement in-cluster → tag as internal mechanism.
        return {
            "Matched EDF Invoice #": f"{cluster.get('name', '')} internal mechanism",
            "Confidence": 0,
            "Notes": (
                "Posting Date inside cluster window but no amount agreement "
                f"with any in-cluster invoice (SAP net £{comparison_amount:.2f})"
            ),
            "Evidence Trail": (
                f"Posting Date: {posting_date}, "
                f"cluster window: {cluster_start}..{cluster_end}, "
                f"cluster: {cluster.get('name', '')}"
            ),
        }

    return None


# compute_dispute_flags is the canonical dispute-flag detector.  It
# previously had an independent (divergent) implementation here; that
# copy has been removed.  This module re-exports the canonical
# ``processors.analysis.compute_dispute_flags`` so Excel, PDF, DOCX and
# HTML surfaces can never disagree about flag amounts or severities.
#
# The import is LAZY (inside the function body) because
# ``processors.analysis`` imports ``_disclosed_label`` and
# ``_reading_type_to_aem`` from THIS module at module scope — a
# top-level import here would create a circular import.  Deferring the
# import to call time breaks the cycle while keeping the re-export
# surface identical (callers can still ``from edf_bill_fetcher.writers
# import compute_dispute_flags`` or mock.patch the name here).
def compute_dispute_flags(dfc: pd.DataFrame, mean_daily: float = 0.0) -> tuple[list, dict]:
    """Re-export of the canonical detector (lazy import — see comment above)."""
    from edf_bill_fetcher.processors.analysis import compute_dispute_flags as _canon

    return _canon(dfc, mean_daily)
