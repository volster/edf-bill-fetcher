"""Typed computed-report results shared by the Excel writer and the PDF/DOCX reporters."""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.helpers.payment_figures import payment_amounts


@dataclass
class PaymentAnalysis:  # noqa: D101
    count: int
    total_paid: float
    avg_payment: float
    median_payment: float
    largest_payment: float
    smallest_payment: float
    avg_interval_days: float | None
    median_interval_days: float | None
    last_payment_date: str | None
    last_payment_amount: float | None
    chronology: pd.DataFrame = field(default_factory=lambda: pd.DataFrame())


def compute_payment_analysis(df: pd.DataFrame) -> PaymentAnalysis:
    """Compute payment/credit statistics for the analysis surfaces."""
    payments = df[df["Entry Type"].isin(["Payment", "Credit"])].copy()
    if payments.empty:
        return PaymentAnalysis(
            count=0,
            total_paid=0.0,
            avg_payment=0.0,
            median_payment=0.0,
            largest_payment=0.0,
            smallest_payment=0.0,
            avg_interval_days=None,
            median_interval_days=None,
            last_payment_date=None,
            last_payment_amount=None,
            chronology=payments,
        )

    payments["_dt"] = payments["Date"].apply(parse_to_sort_date)
    payments = payments.sort_values("_dt").reset_index(drop=True)
    # `_amount` holds the RAW (signed) per-row transaction figure so
    # chronology rendering matches the old `payment_amount(row)[0]`
    # behavior exactly; the abs() mirroring the old dict is applied at
    # the STAT level (abs(sum), abs(mean), ...), not element-wise.
    payments["_amount"] = payment_amounts(payments)

    pay_dates = payments["_dt"].dropna()
    intervals = pay_dates.diff().dt.days.dropna()

    return PaymentAnalysis(
        count=len(payments),
        total_paid=float(abs(payments["_amount"].sum())),
        avg_payment=float(abs(payments["_amount"].mean())),
        median_payment=float(abs(payments["_amount"].median())),
        largest_payment=float(abs(payments["_amount"].max())),
        smallest_payment=float(abs(payments["_amount"].min())),
        avg_interval_days=float(intervals.mean()) if len(intervals) > 0 else None,
        median_interval_days=float(intervals.median()) if len(intervals) > 0 else None,
        last_payment_date=payments.iloc[-1]["Date"] if len(payments) > 0 else None,
        last_payment_amount=float(payments["_amount"].iloc[-1]) if len(payments) > 0 else None,
        chronology=payments,
    )
