"""Financial computations independent of document rendering."""

from __future__ import annotations

from decimal import Decimal

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.processors.preparation import balance_frame, normalize_frame, number, scope


def financial_totals(df: pd.DataFrame) -> tuple[float, float]:
    """Sum charges and payment/credit movements, never account balances."""
    work = normalize_frame(df)
    charges = Decimal("0")
    payments = Decimal("0")
    for _, row in work.iterrows():
        transaction = number(row["Transaction Amount (£)"])
        if transaction is None:
            continue
        value = Decimal(str(transaction))
        if row["Entry Type"] == "Refund":
            payments -= value
        elif row["Entry Type"] in {"Payment", "Credit"} or (
            transaction < 0 and row["Entry Type"] in {"New Bill", "Ongoing Balance", "Charge"}
        ):
            payments += abs(value)
        elif transaction > 0 and row["Entry Type"] in {
            "New Bill",
            "Ongoing Balance",
            "Late Payment",
            "Charge",
        }:
            charges += value
    return float(charges), float(payments)


def balance_extremes(df: pd.DataFrame) -> tuple[float | None, float | None]:
    """First and last observed balance; transactions cannot supply a balance."""
    work = balance_frame(df)
    if work.empty or len({scope(r) for r in work.to_dict("records")}) > 1:
        return None, None
    return float(work.iloc[0]["Balance (£)"]), float(work.iloc[-1]["Balance (£)"])


def bill_daily_rates(df: pd.DataFrame) -> pd.DataFrame:
    """Actual billed charges per consumption day, with consistent units."""
    work = normalize_frame(df)
    start = pd.to_datetime(work["Period From"].apply(parse_to_sort_date))
    end = pd.to_datetime(work["Period To"].apply(parse_to_sort_date))
    days = (end - start).dt.days
    # One-day periods are valid; reversed or missing periods are not.
    days = days.mask(days == 0, 1)
    charge = pd.to_numeric(work["Period Charge (£)"], errors="coerce")
    eligible = (
        work["Entry Type"].isin(["New Bill", "Ongoing Balance", "Charge"])
        & (days > 0)
        & (charge > 0)
    )
    result = work[eligible].copy()
    result["_days"] = days[eligible]
    result["_daily"] = charge[eligible] / days[eligible]
    return result


def mean_daily_rate(df: pd.DataFrame) -> float:
    """Time-weighted billed daily rate, without a second monthly conversion."""
    rates = bill_daily_rates(df)
    if rates.empty:
        return 0.0
    return float((rates["_daily"] * rates["_days"]).sum() / rates["_days"].sum())
