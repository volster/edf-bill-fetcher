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
    RECON_MIN_TOLERANCE,
    RECON_PCT_TOLERANCE,
)
from edf_bill_fetcher.processors.detection import _ambiguous_invoice_references, _reversal_match
from edf_bill_fetcher.writers._helpers import _disclosed_label, _reading_type_to_aem


def compute_dispute_flags(dfc: pd.DataFrame, mean_daily: float = 0.0) -> tuple[list, dict]:
    """Compute scoped findings from explicit balances and billed daily charges."""
    from edf_bill_fetcher.processors.financial import bill_daily_rates, mean_daily_rate
    from edf_bill_fetcher.processors.preparation import (
        balance_frame,
        identity,
        normalize_frame,
        number,
        scope,
    )

    for i, raw in enumerate(dfc.to_dict("records")):
        value = raw.get("Amount (£)")
        if identity(value) and number(value) is None:
            warnings.warn(
                f"Balance heuristics could not evaluate invalid amount at row index {i + 1}",
                UserWarning,
                stacklevel=2,
            )

    work = normalize_frame(dfc)
    flags = []
    work["_scope"] = [scope(r) for r in work.to_dict("records")]
    for _, group in work.groupby("_scope", sort=False):
        for invoice in sorted(_ambiguous_invoice_references(group)):
            observations = group[group["Invoice #"].map(identity) == invoice]
            dated = observations[observations["_dt"].notna()].sort_values("_dt", kind="stable")
            date = dated.iloc[-1]["Date"] if not dated.empty else "N/A"
            flags.append(
                (
                    "INVOICE IDENTITY CONFLICT",
                    date,
                    None,
                    f"Invoice {invoice} has differing bill observations within the same ledger. "
                    "Automatic supersession is withheld; review the original documents.",
                    "MEDIUM",
                )
            )
        bills = group[group["Entry Type"].isin(["New Bill", "Ongoing Balance"])].sort_values("_dt")
        dated_bills = bills[bills["_dt"].notna()]
        for i in range(1, len(dated_bills)):
            previous, current = dated_bills.iloc[i - 1], dated_bills.iloc[i]
            days = (current["_dt"] - previous["_dt"]).days
            if days > BILLING_GAP_MIN_DAYS:
                flags.append(
                    (
                        "BILLING GAP",
                        current["Date"],
                        current["Balance (£)"],
                        f"{days} days between recorded bills (previous: {previous['Date']}). Check for missing evidence.",
                        "HIGH" if days > BILLING_GAP_HIGH_DAYS else "MEDIUM",
                    )
                )
        run = []
        for _, row in bills.iterrows():
            if str(row["Reading"]).lower() in {"estimated", "est."}:
                run.append(row)
            else:
                if len(run) >= ESTIMATED_RUN_MIN:
                    flags.append(
                        (
                            "ESTIMATED RUN",
                            run[0]["Date"],
                            None,
                            f"{len(run)} consecutive estimated readings from {run[0]['Date']}.",
                            "HIGH",
                        )
                    )
                run = []
        if len(run) >= ESTIMATED_RUN_MIN:
            flags.append(
                (
                    "ESTIMATED RUN",
                    run[0]["Date"],
                    None,
                    f"{len(run)} consecutive estimated readings from {run[0]['Date']} (ongoing).",
                    "HIGH",
                )
            )
        baseline = mean_daily_rate(bills)
        for _, row in bill_daily_rates(bills).iterrows():
            ratio = row["_daily"] / baseline if baseline > 0 else 0
            if ratio > HIGH_DAILY_RATE_RATIO:
                flags.append(
                    (
                        "HIGH DAILY RATE",
                        row["Date"],
                        row["Period Charge (£)"],
                        f"£{row['_daily']:,.2f}/day ({ratio:.1f}× avg £{baseline:,.2f}/day) over {int(row['_days'])} consumption days",
                        "HIGH" if ratio > HIGH_DAILY_RATE_HIGH_RATIO else "MEDIUM",
                    )
                )
        balances = balance_frame(group)
        for i in range(1, len(balances)):
            previous, current = balances.iloc[i - 1], balances.iloc[i]
            delta = current["Balance (£)"] - previous["Balance (£)"]
            days = (current["_dt"] - previous["_dt"]).days
            pct = delta / previous["Balance (£)"] if previous["Balance (£)"] > 0 else 0
            if pct > LARGE_JUMP_PCT and 0 < days <= LARGE_JUMP_MAX_DAYS:
                flags.append(
                    (
                        "LARGE JUMP",
                        current["Date"],
                        delta,
                        f"+£{delta:,.2f} (+{pct * 100:.1f}%) in {days} days (from {previous['Date']}: £{previous['Balance (£)']:,.2f})",
                        "HIGH" if pct > LARGE_JUMP_HIGH_PCT else "MEDIUM",
                    )
                )
            if delta < -BALANCE_REDUCTION_AMOUNT:
                flags.append(
                    (
                        "BALANCE REDUCTION",
                        current["Date"],
                        abs(delta),
                        f"Balance fell £{abs(delta):,.2f}; verify the underlying payments or adjustments.",
                        "INFO",
                    )
                )
            between = group[(group["_dt"] > previous["_dt"]) & (group["_dt"] <= current["_dt"])]
            movements = pd.to_numeric(between["Transaction Amount (£)"], errors="coerce").dropna()
            if days > 0 and not movements.empty:
                expected = float(movements.sum())
                difference = abs(delta - expected)
                if difference > max(abs(expected) * RECON_PCT_TOLERANCE, RECON_MIN_TOLERANCE):
                    flags.append(
                        (
                            "RECONCILIATION MISMATCH",
                            current["Date"],
                            difference,
                            f"Observed balance movement £{delta:,.2f} vs recorded transactions £{expected:,.2f}. "
                            "The ledger may be incomplete; this alone does not establish a billing error.",
                            "INFO",
                        )
                    )
    counts = {s: sum(f[4] == s for f in flags) for s in ("HIGH", "MEDIUM", "INFO")}
    return flags, counts


# ---------------------------------------------------------------------------
# Write evidence sheet
# ---------------------------------------------------------------------------


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


def _analyze_tariff_impact(df):
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


def _data_quality_report(df):
    """Alias for the shared computed model (see models/report_models.py)."""
    from edf_bill_fetcher.models.report_models import compute_data_quality_report

    return compute_data_quality_report(df).to_dict()


# ---------------------------------------------------------------------------
# NEW ANALYSIS TAB WRITERS
# ---------------------------------------------------------------------------


__all__ = [
    "compute_dispute_flags",
    "_data_quality_report",
    "_detect_payment_patterns",
    "_analyze_tariff_impact",
    "_disclosed_label",
    "_reversal_match",
    "_reading_type_to_aem",
]
