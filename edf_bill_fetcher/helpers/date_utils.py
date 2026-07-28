"""Date and statistics helpers for the evidence workbook.

Extracted from ``edf_collector.py`` as part of the modularization refactor
(Task 3).  These cover:

- ``completeness_score`` — counts populated substantive fields on a
  record row (used as the primary dedup sort key).
- ``compute_ema`` / ``compute_momentum`` / ``compute_rolling_stats`` —
  pandas-powered time-series statistics.
- ``build_evidence_trail`` — human-readable one-line cluster narrative.
"""

from __future__ import annotations

from typing import Any

COMPLETENESS_FIELDS: tuple[str, ...] = (
    "Date",
    "Period From",
    "Period To",
    "Invoice #",
    "Period Charge (£)",
    "Unit Rate (p/kWh)",
    "Entry Type",
    "Reading",
    "Units (kWh)",
    "Standing Chg (p/day)",
    "Tariff",
)


def completeness_score(row: Any) -> int:
    """Count populated substantive fields on a record row.

    Used as the primary sort key in the dedup pass so the row with the
    most populated ``COMPLETENESS_FIELDS`` ends up first (and thus
    survives ``keep="first"``).  Lower score = sparser row; ties
    fall through to source precedence and then date.

    A value counts as "populated" if it is not None, not NaN, and (for
    strings) not ``""`` and not ``"N/A"``.

    Deliberately excluded: ``% Change``, ``Anomaly Flag``, ``Duplicate Of``,
    and ``Logic Used`` (don't reflect user data) and ``Amount (£)``
    (the dedup key — every sibling has it by definition).
    """
    import math

    count = 0
    for f in COMPLETENESS_FIELDS:
        if f not in row.index:
            continue
        v = row[f]
        if v is None:
            continue
        if isinstance(v, float) and math.isnan(v):
            continue
        if isinstance(v, str):
            s = v.strip()
            if s == "" or s == "N/A":
                continue
        count += 1
    return count


def compute_rolling_stats(series: Any, window: int = 6) -> dict[str, Any]:
    """Compute rolling statistics for a time series."""
    return {
        "mean": series.rolling(window=window, min_periods=1).mean(),
        "std": series.rolling(window=window, min_periods=1).std(),
        "min": series.rolling(window=window, min_periods=1).min(),
        "max": series.rolling(window=window, min_periods=1).max(),
        "median": series.rolling(window=window, min_periods=1).median(),
    }


def compute_ema(series: Any, span: int = 6) -> Any:
    """Compute Exponential Moving Average."""
    return series.ewm(span=span, adjust=False).mean()


def compute_momentum(series: Any, period: int = 3) -> Any:
    """Compute momentum (rate of change) of a series."""
    return series.diff(period)


def build_evidence_trail(rows: list[dict[str, Any]]) -> str:
    """Stitch a human-readable one-line narrative of the cluster.

    Each row contributes its ``Evidence Trail`` or one synthesized from
    its clearing doc + amount + posting date; rows are joined with
    ``; `` separators.
    """
    parts: list[str] = []
    for r in rows:
        trail = r.get("Evidence Trail") or r.get("evidence_trail") or ""
        if not trail:
            cd = r.get("Clearing doc", r.get("clearing_doc", ""))
            amt = r.get("Net amount", r.get("net_amount", ""))
            pd_ = r.get("Posting date", r.get("posting_date", ""))
            trail = f"{cd} £{amt} ({pd_})"
        parts.append(str(trail))
    return "; ".join(parts)


__all__ = [
    "completeness_score",
    "compute_rolling_stats",
    "compute_ema",
    "compute_momentum",
    "build_evidence_trail",
]
