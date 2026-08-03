"""Writer-specific helpers and constants extracted from edf_collector.py."""

from __future__ import annotations

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.styles import Font

try:
    from statsmodels.tsa.holtwinters import ExponentialSmoothing

    HAS_STATSMODELS = True
except ImportError:
    HAS_STATSMODELS = False

from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date

# Colour constants used across the writer functions.
EDF_ORANGE = "#FE5716"
EDF_NAVY = "#10367A"
EDF_OFFWHITE = "#F5F5F5"
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
    """Analyze the impact of tariff changes on unit rates and charges."""
    if "Tariff" not in df.columns or "Unit Rate (p/kWh)" not in df.columns:
        return {}

    tariff_data = df[df["Tariff"].notna() & (df["Tariff"] != "N/A")].copy()
    if tariff_data.empty:
        return {}

    # Convert unit rate to numeric
    tariff_data["unit_rate_num"] = pd.to_numeric(
        tariff_data["Unit Rate (p/kWh)"], errors="coerce"
    )
    tariff_data = tariff_data.dropna(subset=["unit_rate_num"])

    if tariff_data.empty:
        return {}

    # Group by tariff
    tariff_stats = (
        tariff_data.groupby("Tariff")
        .agg(
            count=("unit_rate_num", "count"),
            avg_unit_rate=("unit_rate_num", "mean"),
            median_unit_rate=("unit_rate_num", "median"),
            min_unit_rate=("unit_rate_num", "min"),
            max_unit_rate=("unit_rate_num", "max"),
            avg_charge=(
                "Period Charge (£)",
                lambda x: pd.to_numeric(x, errors="coerce").mean(),
            ),
        )
        .reset_index()
    )

    # Find tariff changes
    tariff_data = tariff_data.sort_values(
        "_dt" if "_dt" in tariff_data.columns else "Date"
    )
    tariff_changes = tariff_data["Tariff"].ne(tariff_data["Tariff"].shift()).cumsum()

    return {
        "tariff_stats": tariff_stats,
        "num_tariffs": int(tariff_data["Tariff"].nunique()),
        "tariff_changes": int(tariff_changes.max()) if not tariff_changes.empty else 0,
    }


def _data_quality_report(df: pd.DataFrame) -> dict[str, object]:
    """Generate a comprehensive data quality report.

    Works on a *copy* of the input DataFrame so the caller's data is
    never mutated (previously this added ``_dt_parsed`` as a side-effect
    on the caller's df, which broke downstream code that re-used the
    same DataFrame for other purposes).
    """
    df = df.copy()
    total_records = len(df)
    if total_records == 0:
        return {}

    # Date parsing success
    from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date

    df["_dt_parsed"] = df["Date"].apply(parse_to_sort_date)
    date_parsed = df["_dt_parsed"].notna().sum()
    date_failed = total_records - date_parsed

    # Amount completeness
    amt_complete = df["Amount (£)"].notna().sum()
    amt_missing = total_records - amt_complete

    # Period info completeness
    period_from_complete = (df["Period From"] != "N/A").sum()
    _ = (df["Period To"] != "N/A").sum()  # Not used, but computed for completeness
    period_complete = period_from_complete  # At least from date

    # Reading classification
    reading_classified = (df["Reading"] != "N/A").sum() if "Reading" in df.columns else 0

    # Unit rate computable — count numeric values only.
    ur_computable = df["Unit Rate (p/kWh)"].apply(
        lambda x: isinstance(x, int | float)
    ).sum()

    # Duplicates (same date + amount)
    dup_count = df.duplicated(subset=["Date", "Amount (£)"]).sum()

    # Source distribution
    source_dist = df["Source"].value_counts().to_dict()

    # Entry type distribution
    entry_dist = (
        df["Entry Type"].value_counts().to_dict()
        if "Entry Type" in df.columns
        else {}
    )

    return {
        "total_records": total_records,
        "date_parsed": int(date_parsed),
        "date_failed": int(date_failed),
        "date_parse_rate": date_parsed / total_records if total_records > 0 else 0,
        "amt_complete": int(amt_complete),
        "amt_missing": int(amt_missing),
        "period_complete": int(period_complete),
        "period_completeness_rate": (
            period_complete / total_records if total_records > 0 else 0
        ),
        "reading_classified": int(reading_classified),
        "reading_classify_rate": (
            reading_classified / total_records if total_records > 0 else 0
        ),
        "ur_computable": int(ur_computable),
        "ur_computable_rate": (
            ur_computable / total_records if total_records > 0 else 0
        ),
        "duplicate_count": int(dup_count),
        "duplicate_rate": dup_count / total_records if total_records > 0 else 0,
        "source_distribution": source_dist,
        "entry_type_distribution": entry_dist,
    }


def _disclosed_label(admitted: bool, overlaps: bool) -> str:
    """Return the human-readable value of the 'Cancel/Rebill Disclosed'.

    cell used on the Back-billing and Rebilling tabs.

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
    """Map the Reading column's value (Actual/Estimated/Smart/Unknown).

    to the single-letter A/E/M code used on the Meter Readings tab.
    """
    if reading_value == "Actual":
        return "A"
    if reading_value == "Estimated":
        return "E"
    if reading_value == "Smart":
        return "A"
    return "E"


def build_evidence_index(
    df: pd.DataFrame, header_row_offset: int = 1
) -> dict[str, int]:
    """Map match-key signatures to the Excel row on the Evidence Report."""
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return {}
    index: dict[str, int] = {}
    rows_iter = df.iterrows()
    for i, r in rows_iter:
        row_no = header_row_offset + 1 + i  # Excel row (header row + i + 1)
        inv = r.get("Invoice #", "")
        if isinstance(inv, str) and inv and inv != "N/A":
            key = f"inv:{inv}"
            index.setdefault(key, row_no)
        amt = r.get("Amount (£)", "")
        pf = pd.to_datetime(r.get("Period From"), dayfirst=True, errors="coerce")
        pt = pd.to_datetime(r.get("Period To"), dayfirst=True, errors="coerce")
        if pd.isna(pf) or pd.isna(pt):
            continue
        try:
            amt_f = float(amt)
        except (TypeError, ValueError):
            continue
        days = str((pt - pf).days)
        index.setdefault(f"amt_days:{amt_f:.2f}|{days}", row_no)
    return index


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


def _compute_volatility(series, window=6):
    """Compute rolling volatility (std of returns)."""
    returns = series.pct_change()
    return returns.rolling(window=window, min_periods=1).std()


def _zscore_anomalies(series, threshold=2.5):
    """Detect anomalies using z-score method."""
    if len(series) < 3:
        return pd.Series(False, index=series.index)
    mean = series.mean()
    std = series.std()
    if std == 0:
        return pd.Series(False, index=series.index)
    z_scores = np.abs((series - mean) / std)
    return z_scores > threshold


def _iqr_anomalies(series, multiplier=1.5):
    """Detect anomalies using IQR method."""
    if len(series) < 4:
        return pd.Series(False, index=series.index)
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    if iqr == 0:
        return pd.Series(False, index=series.index)
    lower = q1 - multiplier * iqr
    upper = q3 + multiplier * iqr
    return (series < lower) | (series > upper)


def _linear_forecast_pair(series, steps=6):
    """Simple linear regression: returns (fitted, future) values."""
    if len(series) < 3:
        return None, None
    x = np.arange(len(series))
    y = series.values
    mask = ~np.isnan(y)
    if mask.sum() < 3:
        return None, None
    x_clean = x[mask]
    y_clean = y[mask]
    try:
        coeffs = np.polyfit(x_clean, y_clean, 1)
        fitted = np.polyval(coeffs, x)
        future_x = np.arange(len(series), len(series) + steps)
        forecast = np.polyval(coeffs, future_x)
        return fitted, forecast
    except Exception:
        return None, None


def _holt_winters_forecast_pair(series, steps=6, seasonal_periods=None):
    """Holt-Winters: returns (fitted, future) values (if statsmodels available)."""
    if not HAS_STATSMODELS or len(series) < 4:
        return None, None
    try:
        clean_series = series.dropna()
        if len(clean_series) < 4:
            return None, None
        if seasonal_periods is None:
            seasonal_periods = (
                min(12, len(clean_series) // 2)
                if len(clean_series) >= 8
                else None
            )
        model = ExponentialSmoothing(
            clean_series,
            trend="add",
            seasonal="add" if seasonal_periods else None,
            seasonal_periods=seasonal_periods,
            initialization_method="estimated",
        )
        fitted_model = model.fit(optimized=True)
        fitted_vals = fitted_model.fittedvalues.reindex(series.index)
        forecast = fitted_model.forecast(steps).values
        return fitted_vals.values, forecast
    except Exception:
        return None, None


def _linear_forecast(series, steps=6):
    """Simple linear regression forecast (forward-only legacy entry point)."""
    _, forecast = _linear_forecast_pair(series, steps)
    return forecast


def _holt_winters_forecast(series, steps=6, seasonal_periods=None):
    """Holt-Winters forward-only legacy entry point."""
    _, forecast = _holt_winters_forecast_pair(series, steps, seasonal_periods)
    return forecast


def _detect_payment_patterns(df):
    """Analyze payment/credit patterns in the data."""
    payments = df[df["Entry Type"].isin(["Payment", "Credit"])].copy()
    if payments.empty:
        return {}

    payments["_dt"] = payments["Date"].apply(parse_to_sort_date)
    payments = payments.sort_values("_dt")

    pay_dates = payments["_dt"].dropna()
    intervals = pay_dates.diff().dt.days.dropna()

    if "Period Charge (£)" in payments.columns:
        pc_numeric = pd.to_numeric(payments["Period Charge (£)"], errors="coerce")
    else:
        pc_numeric = pd.Series(
            [float("nan")] * len(payments), index=payments.index
        )
    amt_numeric = pd.to_numeric(payments["Amount (£)"], errors="coerce")
    pay_amounts = pc_numeric.where(pc_numeric.notna() & (pc_numeric > 0), amt_numeric)

    return {
        "count": len(payments),
        "total_paid": abs(pay_amounts.sum()),
        "avg_payment": abs(pay_amounts.mean()),
        "median_payment": abs(pay_amounts.median()),
        "max_payment": abs(pay_amounts.max()),
        "min_payment": abs(pay_amounts.min()),
        "avg_interval_days": float(intervals.mean()) if len(intervals) > 0 else None,
        "median_interval_days": float(intervals.median()) if len(intervals) > 0 else None,
        "last_payment_date": payments.iloc[-1]["Date"] if len(payments) > 0 else None,
        "last_payment_amount": (
            abs(pay_amounts.iloc[-1]) if len(pay_amounts) > 0 else None
        ),
    }


def _assess_reason(invoice, days, admitted, period_from, period_to):
    """Return a short, deterministic narrative for the Reason Assessment column."""
    pf = period_from.strftime("%d %b %Y")
    pt = period_to.strftime("%d %b %Y")
    excess = days - 365
    if admitted:
        return (
            f"Invoice {invoice} billed {days} days ({pf} to {pt}), "
            f"{excess} days past the 12-month back-billing limit. "
            "EDF's cover page admits a cancellation/reversal, which is "
            "direct evidence the bill is a back-billing remedy."
        )
    return (
        f"Invoice {invoice} billed {days} days ({pf} to {pt}), "
        f"{excess} days past the 12-month back-billing limit. No "
        "admit-phrase was found on the cover page."
    )


# SAP back-billing constants
_SAP_DEBT_MGMT_FLAG_VALUE = "Installment Plan Item"
_SAP_MIN_CLUSTER_SIZE = 4
_SAP_MATCH_DAY_BANDS = ((0, 50), (3, 25), (14, 5))
_SAP_MATCH_AMOUNT_BANDS = ((0.05, 40), (0.25, 20), (0.50, 5))
_SAP_CONFIDENCE_BANDS = (("High", 75), ("Medium", 40), ("Low", 10))


def _parse_amount_for_event(v: object) -> float:
    """Parse the SAP ``Amount`` field (often a string with commas)."""
    if v is None:
        return 0.0
    try:
        s = str(v).strip().lstrip("£").replace(",", "")
        if not s:
            return 0.0
        return float(s)
    except ValueError:
        return 0.0


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


def match_sap_events_to_edf(events: list, edf_records: list[dict]) -> list:
    """Fuzzy-match SAP events to EDF Evidence Report rows."""
    from edf_bill_fetcher.helpers.date_utils import _safe_to_datetime
    from edf_bill_fetcher.models.events import SapEdfMatch

    if not events or not edf_records:
        return []

    parsed_edf: list = []
    for i, rec in enumerate(edf_records):
        invoice = str(rec.get("Invoice #", "")).strip()
        if not invoice or invoice in ("N/A", "None"):
            continue
        pf = _safe_to_datetime(rec.get("Period From"))
        pt = _safe_to_datetime(rec.get("Period To"))
        amt_raw = rec.get("Amount (£)")
        try:
            amt = float(str(amt_raw).replace(",", "").lstrip("£").strip())
        except (TypeError, ValueError):
            amt = 0.0
        parsed_edf.append((i, pf, pt, amt, invoice))

    matches: list = []
    for ev in events:
        if pd.isna(ev.clearing_date):
            continue
        ev_cd = pd.Timestamp(ev.clearing_date)
        for idx, pf, pt, edf_amt, _invoice in parsed_edf:
            if not pd.isna(pt):
                date_delta_days = int(abs((ev_cd - pd.Timestamp(pt)).days))
            elif not pd.isna(pf):
                date_delta_days = int(abs((ev_cd - pd.Timestamp(pf)).days))
            else:
                continue

            amount_score = 0
            if abs(ev.net_amount) < 1.0 and edf_amt > 0:
                best_rel_delta = float("inf")
                for r in ev.rows:
                    row_amt = _parse_amount_for_event(r.get("Amount"))
                    if abs(row_amt) < 1 or abs(edf_amt) < 1:
                        continue
                    rel_delta = abs(abs(row_amt) - edf_amt) / max(abs(edf_amt), 0.01)
                    if rel_delta < best_rel_delta:
                        best_rel_delta = rel_delta
                if best_rel_delta != float("inf"):
                    for band_amt, band_score in _SAP_MATCH_AMOUNT_BANDS:
                        if best_rel_delta <= band_amt:
                            amount_score = band_score
                            break
            elif ev.net_amount != 0 and edf_amt > 0:
                ratio = ev.net_amount / edf_amt
                if 0.95 <= ratio <= 1.05:
                    amount_score = 40
                elif 0.75 <= ratio <= 1.25:
                    amount_score = 20
                elif 0.50 <= ratio <= 1.50:
                    amount_score = 5

            date_score = 0
            date_in_span = False
            if (
                not pd.isna(pf)
                and not pd.isna(pt)
                and pd.Timestamp(pf) <= ev_cd <= pd.Timestamp(pt)
            ):
                date_in_span = True
                if amount_score > 0:
                    date_score = 50
                else:
                    delta_to_pf = abs((ev_cd - pd.Timestamp(pf)).days)
                    delta_to_pt = abs((ev_cd - pd.Timestamp(pt)).days)
                    nearest_delta = min(delta_to_pf, delta_to_pt)
                    for band_days, band_score in _SAP_MATCH_DAY_BANDS:
                        if nearest_delta <= band_days:
                            date_score = band_score
                            break
            else:
                for band_days, band_score in _SAP_MATCH_DAY_BANDS:
                    if date_delta_days <= band_days:
                        date_score = band_score
                        break

            total_score = date_score + amount_score
            band = _confidence_band(total_score)
            if band in ("High", "Medium") and amount_score == 0:
                band = "Low" if total_score >= 10 else None
            if band is None:
                continue

            amt_delta = round(ev.net_amount - edf_amt, 2)
            if band == "High":
                notes = (
                    "Clearing date inside EDF period + amount within 5%"
                    if date_in_span
                    else f"Within {date_delta_days}d of period-end + amount within 5%"
                )
            elif band == "Medium":
                if date_in_span:
                    notes = "Clearing date inside EDF period + amount within 25%"
                else:
                    notes = f"Within {date_delta_days}d of period-end + amount within 25%"
            else:
                if date_in_span and amount_score == 0:
                    notes = (
                        "Clearing date inside EDF period but amounts do not "
                        "correspond — likely coincidental"
                    )
                else:
                    notes = f"Within {date_delta_days}d of period-end; may be coincidental"

            matches.append(
                SapEdfMatch(
                    event=ev,
                    edf_record=edf_records[idx],
                    confidence_band=band,
                    confidence_score=total_score,
                    amount_delta=amt_delta,
                    date_delta_days=date_delta_days,
                    notes=notes,
                )
            )
    return matches


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
            if pct > 0.25 and 0 < days <= 90:
                flags.append(
                    (
                        "LARGE JUMP",
                        c_["Date"],
                        c_["Amount (£)"],
                        f"+£{chg:,.2f} (+{pct * 100:.1f}%) in {days} days (from {p['Date']}: £{p['Amount (£)']:,.2f})",
                        "HIGH" if pct > 0.5 else "MEDIUM",
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
            if days > 60:
                flags.append(
                    (
                        "BILLING GAP",
                        c_["Date"],
                        c_["Amount (£)"],
                        f"{days} days without a bill (previous: {p['Date']}). Balance accumulated unchecked.",
                        "HIGH" if days > 120 else "MEDIUM",
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
                if run >= 3:
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
        if run >= 3:
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
                    if ratio > 2.5:
                        flags.append(
                            (
                                "HIGH DAILY RATE",
                                c_["Date"],
                                c_["Amount (£)"],
                                f"£{daily:,.2f}/day ({ratio:.1f}× avg £{mean_daily:,.2f}/day) over {days} days",
                                "HIGH" if ratio > 4 else "MEDIUM",
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
            if chg < -500:
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
                    threshold = max(pc_val * 0.10, 50.0) if pc_val > 0 else 50.0
                    if diff > threshold:
                        flags.append(
                            (
                                "RECONCILIATION MISMATCH",
                                c_["Date"],
                                c_["Amount (£)"],
                                f"Balance delta £{balance_delta:,.2f} vs period charge £{pc_val:,.2f} "
                                f"(difference: £{diff:,.2f}). Possible payment, credit, or billing error "
                                f"between {p['Date']} and {c_['Date']}.",
                                "HIGH" if diff > pc_val * 0.5 else "MEDIUM",
                            )
                        )
            except (ValueError, TypeError, KeyError):
                pass

    counts = {s: sum(1 for f in flags if f[4] == s) for s in ("HIGH", "MEDIUM", "INFO")}
    return flags, counts
