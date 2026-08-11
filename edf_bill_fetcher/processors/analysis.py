"""Pure-pandas analysis helpers for dispute flags, payment patterns, tariff impact, data quality.

Extracted from ``edf_collector.py`` as part of the modularization refactor
(Task 5 - Phase 4).  ``run_analysers`` (the workbook orchestrator) stays in
``edf_collector.py`` for now; it moves to ``io/writers/analysis.py`` in Task 6.

Compat re-exports live in ``edf_collector.py`` so callers using
``from edf_collector import compute_dispute_flags`` continue to work;
stripped by Task 7.
"""

from __future__ import annotations

import warnings

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import (
    _safe_to_datetime,
    parse_to_sort_date,
)


def compute_dispute_flags(dfc: pd.DataFrame, mean_daily: float = 0.0) -> tuple[list, dict]:
    """Compute dispute flags from a sorted DataFrame.

    Returns:
        tuple: (flags_list, flag_counts_dict)
        - flags_list: list of (type, date, amount, detail, severity) tuples
        - flag_counts_dict: dict with HIGH, MEDIUM, INFO counts

    Issues a :func:`warnings.warn` for any row that fails to evaluate
    under each heuristic (parse error, missing key, etc.).  Previously
    those rows were silently swallowed and the report lost the
    surrounding evidence — turning them into warnings surfaces a
    developer-visible signal without breaking the run.

    """

    def _flag_or_warn(
        row_idx: int,
        flag_name: str,
        exc: BaseException,
    ) -> None:
        warnings.warn(
            (
                f"compute_dispute_flags[{flag_name}] could not evaluate "
                f"row index {row_idx}: {exc!r}; row silently skipped."
            ),
            stacklevel=3,
        )

    flags: list[tuple[str, str | float | None, float | None, str, str]] = []
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
                        chg,  # delta (jump size), not the running balance
                        f"+£{chg:,.2f} (+{pct * 100:.1f}%) in {days} days (from {p['Date']}: £{p['Amount (£)']:,.2f})",
                        "HIGH" if pct > 0.5 else "MEDIUM",
                    )
                )
        except (ValueError, TypeError, KeyError) as exc:
            _flag_or_warn(i, "LARGE_JUMP", exc)

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
        except (ValueError, TypeError, KeyError) as exc:
            _flag_or_warn(i, "BILLING_GAP", exc)

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
            except (ValueError, TypeError, KeyError, ZeroDivisionError) as exc:
                _flag_or_warn(i, "HIGH_DAILY_RATE", exc)

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
                        abs(chg),  # reduction size, not the running balance
                        f"Balance fell £{abs(chg):,.2f} (from £{p['Amount (£)']:,.2f} to £{c_['Amount (£)']:,.2f}).",
                        "INFO",
                    )
                )
        except (ValueError, TypeError, KeyError) as exc:
            _flag_or_warn(i, "BALANCE_REDUCTION", exc)

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
            except (ValueError, TypeError, KeyError) as exc:
                _flag_or_warn(i, "RECONCILIATION_MISMATCH", exc)

    # Count by severity
    counts = {s: sum(1 for f in flags if f[4] == s) for s in ("HIGH", "MEDIUM", "INFO")}
    return flags, counts


# ---------------------------------------------------------------------------
# Write evidence sheet
# ---------------------------------------------------------------------------


def _detect_payment_patterns(df):
    """Analyze payment/credit patterns in the data.

    The per-row transaction amount (the customer's actual payment or
    EDF's actual credit) lives in ``Period Charge (£)`` for HTM
    Payment/Credit rows. ``Amount (£)`` on those rows carries the
    *running balance after the transaction* -- using it as the
    "payment amount" used to flood the Payment Analysis sheet with
    huge balance figures masquerading as payments. Prefer
    ``Period Charge (£)`` when the row has a numeric value there,
    falling back to ``Amount (£)`` for legacy / PST-only rows that
    never populated ``Period Charge (£)``.
    """
    payments = df[df["Entry Type"].isin(["Payment", "Credit"])].copy()
    if payments.empty:
        return {}

    payments["_dt"] = payments["Date"].apply(parse_to_sort_date)
    payments = payments.sort_values("_dt")

    # Calculate days between payments
    pay_dates = payments["_dt"].dropna()
    intervals = pay_dates.diff().dt.days.dropna()

    # Per-row transaction amount: prefer Period Charge (£) (the actual
    # payment / credit), fall back to Amount (£) when Period Charge is
    # missing or non-numeric (legacy rows that never populated it, or
    # older callers passing a DataFrame without the column).
    if "Period Charge (£)" in payments.columns:
        pc_numeric = pd.to_numeric(payments["Period Charge (£)"], errors="coerce")
    else:
        pc_numeric = pd.Series([float("nan")] * len(payments), index=payments.index)
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
        "last_payment_amount": abs(pay_amounts.iloc[-1]) if len(pay_amounts) > 0 else None,
    }


def _analyze_tariff_impact(df):
    """Analyze the impact of tariff changes on unit rates and charges."""
    if "Tariff" not in df.columns or "Unit Rate (p/kWh)" not in df.columns:
        return {}

    tariff_data = df[df["Tariff"].notna() & (df["Tariff"] != "N/A")].copy()
    if tariff_data.empty:
        return {}

    # Convert unit rate to numeric
    tariff_data["unit_rate_num"] = pd.to_numeric(tariff_data["Unit Rate (p/kWh)"], errors="coerce")
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
            avg_charge=("Period Charge (£)", lambda x: pd.to_numeric(x, errors="coerce").mean()),
        )
        .reset_index()
    )

    # Find tariff changes
    tariff_data = tariff_data.sort_values("_dt" if "_dt" in tariff_data.columns else "Date")
    tariff_changes = tariff_data["Tariff"].ne(tariff_data["Tariff"].shift()).cumsum()

    return {
        "tariff_stats": tariff_stats,
        "num_tariffs": tariff_data["Tariff"].nunique(),
        "tariff_changes": int(tariff_changes.max()) if not tariff_changes.empty else 0,
    }


def _data_quality_report(df):
    """Generate a comprehensive data quality report.

    Works on a *copy* of the input DataFrame so the caller's data is
    never mutated (previously this added ``_dt_parsed`` as a side-effect
    on the caller's df, which broke downstream code that re-used the
    same DataFrame for other purposes).
    """
    # Work on a copy to avoid mutating the caller's DataFrame
    df = df.copy()
    total_records = len(df)
    if total_records == 0:
        return {}

    # Date parsing success
    df["_dt_parsed"] = df["Date"].apply(parse_to_sort_date)
    date_parsed = df["_dt_parsed"].notna().sum()
    date_failed = total_records - date_parsed

    # Amount completeness
    amt_complete = df["Amount (£)"].notna().sum()
    amt_missing = total_records - amt_complete

    # Period info completeness — a row counts as period-complete only
    # when BOTH Period From and Period To are present.
    period_complete = ((df["Period From"] != "N/A") & (df["Period To"] != "N/A")).sum()

    # Reading classification
    # Reading classification — "N/A" is the sentinel for unclassified readings
    reading_classified = (df["Reading"] != "N/A").sum() if "Reading" in df.columns else 0

    # Unit rate computable — count numeric values only. The unit
    # rate column can hold `int | float | "N/A"`; only numerics can be
    # used downstream by tariff charts, so other values are excluded.
    # The older draft guarded this with `and x != "N/A"`, which is
    # unreachable for an already-typed numeric — pinned here so a
    # future careless refactor cannot silently change this branch
    # back into a no-op-or-true tautology that overcounts.
    ur_computable = df["Unit Rate (p/kWh)"].apply(lambda x: isinstance(x, int | float)).sum()

    # Duplicates (same date + amount)
    dup_count = df.duplicated(subset=["Date", "Amount (£)"]).sum()

    # Source distribution
    source_dist = df["Source"].value_counts().to_dict()

    # Entry type distribution
    entry_dist = df["Entry Type"].value_counts().to_dict() if "Entry Type" in df.columns else {}

    return {
        "total_records": total_records,
        "date_parsed": date_parsed,
        "date_failed": date_failed,
        "date_parse_rate": date_parsed / total_records if total_records > 0 else 0,
        "amt_complete": amt_complete,
        "amt_missing": amt_missing,
        "period_complete": period_complete,
        "period_completeness_rate": period_complete / total_records if total_records > 0 else 0,
        "reading_classified": reading_classified,
        "reading_classify_rate": reading_classified / total_records if total_records > 0 else 0,
        "ur_computable": ur_computable,
        "ur_computable_rate": ur_computable / total_records if total_records > 0 else 0,
        "duplicate_count": int(dup_count),
        "duplicate_rate": dup_count / total_records if total_records > 0 else 0,
        "source_distribution": source_dist,
        "entry_type_distribution": entry_dist,
    }


# ---------------------------------------------------------------------------
# NEW ANALYSIS TAB WRITERS
# ---------------------------------------------------------------------------


def _disclosed_label(
    admitted: bool,
    overlaps: bool,
) -> str:
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


def _reversal_match(
    evidence_df: pd.DataFrame | None,
    killed_inv: str,
    killed_amount: float | None,
    killed_pf: pd.Timestamp,
    killed_pt: pd.Timestamp,
) -> bool:
    """Return whether a reversal-credit row in *evidence_df* matches the killed invoice well enough to count as rebilling evidence.

    Spec ref: 2026-07-16 §11. A reversal credit accepts the killed
    invoice when its amount is within ±£0.50 AND either its period
    overlaps the killed period by ≥ 30 days OR its period is
    unparseable (so we accept on amount alone, Entry Type == Credit).
    """
    if evidence_df is None or evidence_df.empty:
        return False
    if "Entry Type" not in evidence_df.columns:
        return False
    try:
        amount = abs(float(killed_amount or 0.0))
    except (TypeError, ValueError):
        return False
    matching = evidence_df[evidence_df["Entry Type"].isin(["Credit", "Payment"])]
    for _, row in matching.iterrows():
        try:
            row_amt = abs(float(row.get("Amount (£)", 0) or 0))
        except (TypeError, ValueError):
            continue
        if abs(row_amt - amount) > 0.50:
            continue
        rpf = _safe_to_datetime(row.get("Period From"))
        rpt = _safe_to_datetime(row.get("Period To"))
        if pd.isna(rpf) or pd.isna(rpt):
            return True
        overlap = (min(killed_pt, rpt) - max(killed_pf, rpf)).days
        if overlap >= 30:
            return True
    return False


def _reading_type_to_aem(reading_value: str) -> str:
    """Map the Reading column's value (Actual/Estimated/Smart/Unknown) to the single-letter A/E/M code used on the Meter Readings tab."""
    if reading_value == "Actual":
        return "A"
    if reading_value == "Estimated":
        return "E"
    if reading_value == "Smart":
        return "A"
    return "E"


__all__ = [
    "compute_dispute_flags",
    "_data_quality_report",
    "_detect_payment_patterns",
    "_analyze_tariff_impact",
    "_disclosed_label",
    "_reversal_match",
    "_reading_type_to_aem",
]
