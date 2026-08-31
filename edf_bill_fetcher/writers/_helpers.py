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

from edf_bill_fetcher.helpers.dispute_flags import (
    BALANCE_REDUCTION_AMOUNT,
    BILLING_GAP_HIGH_DAYS,
    BILLING_GAP_MIN_DAYS,
    ESTIMATED_RUN_MIN,
    HIGH_DAILY_RATE_HIGH_RATIO,
    HIGH_DAILY_RATE_RATIO,
    LARGE_JUMP_HIGH_PCT,
    LARGE_JUMP_MAX_DAYS,
    LARGE_JUMP_PCT,
    RECON_HIGH_PCT,
    RECON_MIN_TOLERANCE,
    RECON_PCT_TOLERANCE,
)
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


def compute_dispute_flags(dfc: pd.DataFrame, mean_daily: float = 0.0) -> tuple[list, dict]:
    """Compute dispute flags from a sorted DataFrame."""
    flags: list = []
    n = len(dfc)
    if n < 2:
        return flags, {"HIGH": 0, "MEDIUM": 0, "INFO": 0}

    # 1. LARGE JUMP: >25% increase within 90 days
    for i in range(1, n):
        p = dfc.iloc[i - 1]
        c_ = dfc.iloc[i]
        try:
            chg = float(c_["Amount (£)"]) - float(p["Amount (£)"])
            pct = chg / float(p["Amount (£)"]) if float(p["Amount (£)"]) > 0 else 0
            days = (c_["_dt"] - p["_dt"]).days
            if pct > LARGE_JUMP_PCT and 0 < days <= LARGE_JUMP_MAX_DAYS:
                flags.append(
                    (
                        "LARGE JUMP",
                        c_["Date"],
                        c_["Amount (£)"],
                        f"+£{chg:,.2f} (+{pct * 100:.1f}%) in {days} days (from {p['Date']}: £{p['Amount (£)']:,.2f})",
                        "HIGH" if pct > LARGE_JUMP_HIGH_PCT else "MEDIUM",
                    )
                )
        except (ValueError, TypeError, KeyError):
            pass

    # 2. BILLING GAP: >60 days without a bill
    for i in range(1, n):
        p = dfc.iloc[i - 1]
        c_ = dfc.iloc[i]
        try:
            days = (c_["_dt"] - p["_dt"]).days
            if days > BILLING_GAP_MIN_DAYS:
                flags.append(
                    (
                        "BILLING GAP",
                        c_["Date"],
                        c_["Amount (£)"],
                        f"{days} days without a bill (previous: {p['Date']}). Balance accumulated unchecked.",
                        "HIGH" if days > BILLING_GAP_HIGH_DAYS else "MEDIUM",
                    )
                )
        except (ValueError, TypeError, KeyError):
            pass

    # 3. ESTIMATED RUN: 3+ consecutive estimated readings
    if "Reading" in dfc.columns:
        run = 0
        run_start = None
        for i, rv in enumerate(dfc["Reading"].tolist()):
            if str(rv).lower() in ("estimated", "est."):
                run += 1
                if run == 1:
                    run_start = dfc.iloc[i]["Date"]
            else:
                if run >= ESTIMATED_RUN_MIN:
                    flags.append(
                        (
                            "ESTIMATED RUN",
                            run_start,
                            None,
                            f"{run} consecutive estimated readings from {run_start}.",
                            "HIGH",
                        )
                    )
                run = 0
                run_start = None
        if run >= ESTIMATED_RUN_MIN:
            flags.append(
                (
                    "ESTIMATED RUN",
                    run_start,
                    None,
                    f"{run} consecutive estimated readings from {run_start} (ongoing).",
                    "HIGH",
                )
            )

    # 4. HIGH DAILY RATE: daily rate significantly above average
    if mean_daily > 0:
        for i in range(1, n):
            p = dfc.iloc[i - 1]
            c_ = dfc.iloc[i]
            try:
                days = (c_["_dt"] - p["_dt"]).days
                charge = float(c_["Amount (£)"]) - float(p["Amount (£)"])
                if days > 0 and charge > 0:
                    daily = charge / days
                    ratio = daily / mean_daily
                    if ratio > HIGH_DAILY_RATE_RATIO:
                        flags.append(
                            (
                                "HIGH DAILY RATE",
                                c_["Date"],
                                c_["Amount (£)"],
                                f"£{daily:,.2f}/day ({ratio:.1f}× avg £{mean_daily:,.2f}/day) over {days} days",
                                "HIGH" if ratio > HIGH_DAILY_RATE_HIGH_RATIO else "MEDIUM",
                            )
                        )
            except (ValueError, TypeError, KeyError, ZeroDivisionError):
                pass

    # 5. BALANCE REDUCTION: payment/credit > £500
    for i in range(1, n):
        p = dfc.iloc[i - 1]
        c_ = dfc.iloc[i]
        try:
            chg = float(c_["Amount (£)"]) - float(p["Amount (£)"])
            if chg < -BALANCE_REDUCTION_AMOUNT:
                flags.append(
                    (
                        "BALANCE REDUCTION",
                        c_["Date"],
                        c_["Amount (£)"],
                        f"Balance fell £{abs(chg):,.2f} (from £{p['Amount (£)']:,.2f} to £{c_['Amount (£)']:,.2f}).",
                        "INFO",
                    )
                )
        except (ValueError, TypeError, KeyError):
            pass

    # 6. RECONCILIATION MISMATCH: balance delta vs period charge
    if "Period Charge (£)" in dfc.columns:
        for i in range(1, n):
            p = dfc.iloc[i - 1]
            c_ = dfc.iloc[i]
            try:
                if str(c_.get("Entry Type", "")) == "New Bill" and str(p.get("Entry Type", "")) in (
                    "New Bill",
                    "Ongoing Balance",
                ):
                    pc = c_.get("Period Charge (£)")
                    try:
                        pc_val = float(pc)
                    except (ValueError, TypeError):
                        continue
                    balance_delta = float(c_["Amount (£)"]) - float(p["Amount (£)"])
                    diff = abs(balance_delta - pc_val)
                    threshold = (
                        max(pc_val * RECON_PCT_TOLERANCE, RECON_MIN_TOLERANCE)
                        if pc_val > 0
                        else RECON_MIN_TOLERANCE
                    )
                    if diff > threshold:
                        flags.append(
                            (
                                "RECONCILIATION MISMATCH",
                                c_["Date"],
                                c_["Amount (£)"],
                                f"Balance delta £{balance_delta:,.2f} vs period charge £{pc_val:,.2f} "
                                f"(difference: £{diff:,.2f}). Possible payment, credit, or billing error "
                                f"between {p['Date']} and {c_['Date']}.",
                                "HIGH" if diff > pc_val * RECON_HIGH_PCT else "MEDIUM",
                            )
                        )
            except (ValueError, TypeError, KeyError):
                pass

    counts = {s: sum(1 for f in flags if f[4] == s) for s in ("HIGH", "MEDIUM", "INFO")}
    return flags, counts
