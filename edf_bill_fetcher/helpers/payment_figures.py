"""Shared payment and credit figure selection."""

from __future__ import annotations

import pandas as pd


def payment_amount(row: pd.Series) -> tuple[float, str]:
    """Return the magnitude of an actual payment/credit, never its balance."""
    from edf_bill_fetcher.processors.preparation import normalize_record, number

    normalized = normalize_record(row.to_dict() if isinstance(row, pd.Series) else dict(row))
    transaction = number(normalized["Transaction Amount (£)"])
    return (
        (abs(transaction), "Transaction Amount (£)")
        if transaction is not None
        else (0.0, "Unavailable")
    )


def payment_amounts(rows: pd.DataFrame) -> pd.Series:
    """Return transaction amounts, preferring period charge over balance."""
    values = [payment_amount(row)[0] for _, row in rows.iterrows()]
    return pd.Series(values, index=rows.index, dtype="float64")
