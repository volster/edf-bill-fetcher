from __future__ import annotations  # noqa: I001
from __future__ import annotations

import re as _re
import warnings as _w

import pandas as pd
from typing import Any

"""Date helpers extracted from edf_collector.py for the modularization refactor.
"""



COMPLETENESS_FIELDS: tuple[str, ...] = (
    "Date", "Period From", "Period To", "Invoice #",
    "Period Charge (£)", "Unit Rate (p/kWh)", "Entry Type",
    "Reading", "Units (kWh)", "Standing Chg (p/day)", "Tariff",
    "Attachment Name", "Details",
)

def completeness_score(row: Any) -> int:
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
            return s != "" and s != "N/A"
        return True
    return sum(1 for f in COMPLETENESS_FIELDS if _is_populated(row.get(f)))

def compute_rolling_stats(series: Any, window: int = 6) -> dict[str, Any]:
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
    return series.ewm(span=span, adjust=False).mean()

def compute_momentum(series: Any, period: int = 3) -> Any:
    return series - series.shift(period)

def build_evidence_trail(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "No rows"
    f = rows[0]
    return f'{len(rows)} rows from {f.get("Source", "")} totalling {f.get("Amount (£)", "")}'

_ISO_DATE_RE = _re.compile(r"^\d{4}-\d{2}-\d{2}$")

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

def _safe_to_datetime(value: object, *, dayfirst: bool = True) -> pd.Timestamp | pd.Series:
    if isinstance(value, pd.Series | pd.Index):
        with _w.catch_warnings():
            _w.simplefilter("ignore", UserWarning)
            return pd.to_datetime(value, dayfirst=dayfirst, errors="coerce")
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

def parse_to_display_date(date_input):
    dt = parse_to_sort_date(date_input)
    return dt.strftime("%d/%m/%Y") if not pd.isna(dt) else str(date_input)

def to_excel_date(date_input):
    dt = parse_to_sort_date(date_input)
    return None if pd.isna(dt) else dt.to_pydatetime()

