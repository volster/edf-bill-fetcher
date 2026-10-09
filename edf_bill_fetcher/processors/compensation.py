"""Indicative compensation from live charges and nonoverlapping credit episodes.

Interest is a user-specified scenario, not an asserted statutory entitlement.
No interest is assumed unless a finite non-negative annual rate is supplied.
"""

from __future__ import annotations

from typing import Any

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.processors.detection import (
    compute_transitive_domination,
    compute_unlawful_union_total,
    detect_back_billing,
    detect_rebilling,
)
from edf_bill_fetcher.processors.preparation import normalize_frame, number, scope

DISCLAIMER = (
    "Indicative calculations only. Back-billing eligibility requires review of the "
    "original evidence and applicable rules. Credit interest is a configured scenario, "
    "not a determination of entitlement."
)


def _parse_date(value: object) -> pd.Timestamp | None:
    """Use the same date interpretation as the rest of the pipeline."""
    parsed = parse_to_sort_date(value)
    return None if pd.isna(parsed) else parsed


def _as_of(config: dict[str, Any]) -> pd.Timestamp:
    """Use the run's recorded valuation date."""
    parsed = _parse_date(config.get("as_of"))
    return parsed if parsed is not None else pd.Timestamp.today().normalize()


def estimate_compensation(
    records_df: pd.DataFrame | None, config: dict[str, Any] | None = None
) -> list[dict[str, Any]]:
    """Calculate live back-billing charges and one interest amount per credit episode."""
    if records_df is None or records_df.empty:
        return []
    cfg = config or {}
    rate = number(cfg.get("credit_interest_rate", 0.0))
    hold_days = number(cfg.get("credit_hold_days", 90))
    if rate is None or rate < 0 or hold_days is None or hold_days < 0:
        raise ValueError("Credit rate and hold days must be finite and non-negative")
    valuation = _as_of(cfg)
    work = normalize_frame(records_df)
    work["_scope"] = [scope(r) for r in work.to_dict("records")]
    rows: list[dict[str, Any]] = []
    for _, ledger in work.groupby("_scope", sort=False):
        bb = detect_back_billing(ledger)
        rebills = detect_rebilling(ledger, evidence_df=ledger)
        dominated = compute_transitive_domination(rebills, bb)
        if not bb.empty:
            live = bb[~bb["Invoice #"].astype(str).isin(dominated)].reset_index(drop=True)
            previous_total = 0.0
            for i, bill in live.iterrows():
                union = compute_unlawful_union_total(live.iloc[: i + 1])
                allocated = max(0.0, round(union - previous_total, 2))
                previous_total = union
                if allocated <= 0:
                    continue
                rows.append(
                    {
                        "category": "back_billing_excess",
                        "invoice_ref": str(bill["Invoice #"]),
                        "date": parse_to_sort_date(bill["Bill Date"]).date().isoformat(),
                        "base_amount": float(bill["Period Charge (£)"]),
                        "days": int(bill["Excess Days"]),
                        "rate": None,
                        "indicative_amount": allocated,
                        "legal_basis": "Indicative consumption outside the back-billing window; live union allocation",
                        "disclaimer": DISCLAIMER,
                    }
                )
        if rate == 0:
            continue
        rows.extend(_credit_episodes(ledger, rate, hold_days, valuation))
    return rows


__all__ = ["DISCLAIMER", "estimate_compensation"]


def _credit_episodes(ledger, rate, hold_days, valuation):
    """Accrue one nonoverlapping interest scenario for each observed credit episode."""
    rows = []
    events = ledger[ledger["_dt"].notna() & (ledger["_dt"] <= valuation)].sort_values(
        "_dt", kind="stable"
    )
    balance: float | None = None
    last_date: pd.Timestamp | None = None
    episode_start: pd.Timestamp | None = None
    episode_invoice = ""
    interest = 0.0
    accrued_days = 0
    principal_peak = 0.0

    def accrue(end: pd.Timestamp) -> None:
        nonlocal interest, accrued_days
        if balance is None or balance >= 0 or last_date is None or episode_start is None:
            return
        eligible_start = max(last_date, episode_start + pd.Timedelta(days=int(hold_days)))
        days = max(0, (end - eligible_start).days)
        interest += abs(balance) * rate * days / 365
        accrued_days += days

    def finish() -> None:
        nonlocal interest, accrued_days, principal_peak
        if interest > 0 and episode_start is not None:
            rows.append(
                {
                    "category": "credit_hold_interest",
                    "invoice_ref": episode_invoice,
                    "date": episode_start.date().isoformat(),
                    "base_amount": round(principal_peak, 2),
                    "days": accrued_days,
                    "rate": rate,
                    "indicative_amount": round(interest, 2),
                    "legal_basis": "User-configured simple-interest scenario on observed credit; entitlement unassessed",
                    "disclaimer": DISCLAIMER,
                }
            )
        interest = 0.0
        accrued_days = 0
        principal_peak = 0.0

    for _, event in events.iterrows():
        current_date = event["_dt"]
        accrue(current_date)
        observed = number(event["Balance (£)"])
        transaction = number(event["Transaction Amount (£)"])
        new_balance = observed if observed is not None else balance
        if observed is None and balance is not None and transaction is not None:
            new_balance = round(balance + transaction, 2)
        if new_balance is not None and new_balance < 0:
            if balance is None or balance >= 0:
                episode_start = current_date
                episode_invoice = str(event["Invoice #"])
            principal_peak = max(principal_peak, abs(new_balance))
        elif balance is not None and balance < 0:
            finish()
            episode_start = None
        balance = new_balance
        last_date = current_date
    accrue(valuation)
    finish()
    return rows
