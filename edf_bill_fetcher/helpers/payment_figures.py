"""Shared payment and credit figure selection."""

from __future__ import annotations

import pandas as pd


def payment_amount(row: pd.Series) -> tuple[float, str]:
    """Return the transaction amount and its source for one payment row."""
    period_charge = pd.to_numeric(row.get("Period Charge (£)"), errors="coerce")
    if pd.notna(period_charge):
        return float(period_charge), "Period Charge (£)"
    amount = pd.to_numeric(row.get("Amount (£)"), errors="coerce")
    if pd.notna(amount):
        return float(amount), "Amount (£) fallback"
    return 0.0, "Unavailable"


def payment_amounts(rows: pd.DataFrame) -> pd.Series:
    """Return transaction amounts, preferring period charge over balance."""
    values = [payment_amount(row)[0] for _, row in rows.iterrows()]
    return pd.Series(values, index=rows.index, dtype="float64")
