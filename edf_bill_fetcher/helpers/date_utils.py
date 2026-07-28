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

import re as _re

import pandas as pd

_ISO_DATE_RE = _re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _safe_to_datetime(value: object, *, dayfirst: bool = True) -> pd.Timestamp | pd.Series:
    """Parse ``value`` as a date without triggering Pandas UserWarning noise.

    Uses ``dayfirst=True`` (UK convention) first; for scalars that produce
    NaT, falls back to ``dayfirst=False``.  Passed as the hot-loop parser
    throughout the detection and analysis code where each call would otherwise
    emit a ``UserWarning`` on mixed-format strings.
    """
    if isinstance(value, pd.Series | pd.Index):
        import warnings as _w
        with _w.catch_warnings():
            _w.simplefilter("ignore", UserWarning)
            s = pd.to_datetime(value, dayfirst=dayfirst, errors="coerce")
        return s
    try:
        dt = pd.to_datetime(value, dayfirst=dayfirst, errors="coerce")
    except (TypeError, ValueError):
        return pd.NaT
    if pd.isna(dt) and dayfirst:
        try:
            dt = pd.to_datetime(value, dayfirst=False, errors="coerce")
        except (TypeError, ValueError):
            return pd.NaT
    return dt


def parse_to_sort_date(date_input):
    s = str(date_input).strip() if date_input else ""
    if not s or s in ("Unknown", "N/A", ""):
        return pd.NaT
    try:
        if _ISO_DATE_RE.match(s):
            return pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")
        dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            dt = pd.to_datetime(s, dayfirst=False, errors="coerce")
        return dt
    except Exception:
        return pd.NaT


def parse_to_display_date(date_input):
    dt = parse_to_sort_date(date_input)
    return dt.strftime("%d/%m/%Y") if not pd.isna(dt) else str(date_input)


def to_excel_date(date_input):
    """Return a Python datetime for openpyxl to write as a true Excel date serial."""
    dt = parse_to_sort_date(date_input)
    if pd.isna(dt):
        return None
    return dt.to_pydatetime()


__all__ += [
    "parse_to_sort_date",
    "parse_to_display_date",
    "to_excel_date",
]
