"""Date and time helpers for parsing, formatting, and comparing EDF billing dates.

Date parsing contract
---------------------
ISO dates retain ISO ordering, numeric dates use UK day/month/year, and
missing or invalid dates remain missing. Scalar and vector paths share one
parser; format inference never chooses the interpretation of a whole series.
"""

from __future__ import annotations  # noqa: I001

import re as _re
from datetime import date, datetime
from email.utils import parsedate_to_datetime
from typing import Any, TypeAlias

import pandas as pd

"""Date helpers extracted from edf_collector.py for the modularization refactor.
"""


# pandas exposes no public NaTType; ``pd.NaT`` (the "not a time" sentinel) is
# the runtime value produced by the parsers here, and consumers guard with
# ``pd.isna()``.  Historically the annotation referenced the private
# ``pandas._libs.tslibs.nattype.NaTType`` path, which can move between pandas
# versions; ``TimestampOrNaT`` is the stable, public replacement.
TimestampOrNaT: TypeAlias = "pd.Timestamp | Any"


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
    "Attachment Name",
    "Details",
)


def completeness_score(row: Any) -> int:
    """Count populated fields in a record row across the canonical completeness set."""

    def _is_populated(v: object) -> bool:
        if v is None:
            return False
        try:
            if isinstance(v, float) and pd.isna(v):
                return False
        except (TypeError, ValueError):
            pass
        if isinstance(v, str):
            s = v.strip()
            return s.casefold() not in {"", "n/a", "unknown", "none", "nan", "<na>"}
        return True

    return sum(1 for f in COMPLETENESS_FIELDS if _is_populated(row.get(f)))


def compute_rolling_stats(series: Any, window: int = 6) -> dict[str, Any]:
    """Compute rolling mean/std/min/max/median over a fixed window."""
    if len(series) < window:
        return {}
    return {
        "mean": series.rolling(window=window, min_periods=1).mean(),
        "std": series.rolling(window=window, min_periods=1).std(),
        "min": series.rolling(window=window, min_periods=1).min(),
        "max": series.rolling(window=window, min_periods=1).max(),
        "median": series.rolling(window=window, min_periods=1).median(),
    }


def compute_ema(series: Any, span: int = 6) -> Any:
    """Compute the exponential moving average of a series."""
    return series.ewm(span=span, adjust=False).mean()


def compute_momentum(series: Any, period: int = 3) -> Any:
    """Compute the period-over-period delta of a series."""
    return series - series.shift(period)


def build_evidence_trail(rows: list[dict[str, Any]]) -> str:
    """Build a one-line evidence-trail summary from a list of record rows."""
    if not rows:
        return "No rows"
    f = rows[0]
    return f"{len(rows)} rows from {f.get('Source', '')}"


_ISO_DATE_RE = _re.compile(r"^\d{4}-\d{2}-\d{2}$")


def parse_to_sort_date(date_input):
    """Parse ISO, UK numeric, named-month, and email dates consistently."""
    if isinstance(date_input, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(date_input).tz_localize(None)
    if date_input is None:
        return pd.NaT
    s = str(date_input).strip()
    if s.lower() in {"", "unknown", "n/a", "none", "nan", "nat", "<na>"}:
        return pd.NaT
    try:
        if _re.match(r"^\d{4}-\d{2}-\d{2}(?:$|[T ])", s):
            return pd.Timestamp(datetime.fromisoformat(s.replace("Z", "+00:00"))).tz_localize(None)
        if _re.match(r"^\d{1,2}[/.-]\d{1,2}[/.-]\d{4}", s):
            parts = _re.split(r"[/.-]", s.split()[0])
            return pd.Timestamp(datetime(int(parts[2]), int(parts[1]), int(parts[0])))
        try:
            return pd.Timestamp(parsedate_to_datetime(s)).tz_localize(None)
        except (TypeError, ValueError):
            pass
        if _re.search(r"[A-Za-z]", s):
            return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except (ValueError, TypeError, OverflowError):
        pass
    return pd.NaT


def _safe_to_datetime(value: object, *, dayfirst: bool = True) -> pd.Timestamp | pd.Series:
    """Parse scalar and mixed-format vectors with the same explicit date rules."""
    if isinstance(value, (pd.Series, pd.Index)):
        index = value.index if isinstance(value, pd.Series) else None
        return pd.Series(
            [parse_to_sort_date(v) for v in value], index=index, dtype="datetime64[ns]"
        )
    return parse_to_sort_date(value)


def parse_to_display_date(date_input):
    """Parse a date string into a DD/MM/YYYY display string."""
    dt = parse_to_sort_date(date_input)
    return dt.strftime("%d/%m/%Y") if not pd.isna(dt) else str(date_input)


def to_excel_date(date_input):
    """Parse a date string into a Python datetime suitable for openpyxl cells."""
    dt = parse_to_sort_date(date_input)
    return None if pd.isna(dt) else dt.to_pydatetime()
