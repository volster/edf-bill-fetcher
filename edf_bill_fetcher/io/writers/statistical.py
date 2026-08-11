"""Statistical analysis sheet writer — extracted from writers/__init__.py.

Contains: write_statistical_analysis_sheet — renders the "Statistical Analysis"
worksheet with anomaly detection, distribution moments, and JB/Shapiro tests.
"""

from __future__ import annotations

import importlib.util

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.helpers.excel_utils import (
    hcell as _hcell,
)
from edf_bill_fetcher.helpers.excel_utils import (
    money as _money,
)
from edf_bill_fetcher.helpers.excel_utils import (
    num as _num,
)
from edf_bill_fetcher.helpers.excel_utils import (
    section_hdr as _section_hdr,
)
from edf_bill_fetcher.helpers.excel_utils import (
    text as _text,
)
from edf_bill_fetcher.helpers.theme import CELL_BORDER
from edf_bill_fetcher.writers._helpers import _iqr_anomalies, _zscore_anomalies

# Local scipy-availability probe — mirrors the module-level
# ``HAS_SCIPY`` constant that lived in ``writers/__init__.py`` before
# this function was extracted. Kept private to this module because
# no other writer needs it.
_HAS_SCIPY = importlib.util.find_spec("scipy") is not None


# --- write_statistical_analysis_sheet (was writers/__init__.py L1721-1956) ---


def write_statistical_analysis_sheet(ws, dfc, config):
    """Write Statistical Analysis tab with advanced pandas analytics."""
    ws.title = "Statistical Analysis"

    NAVY = "10367A"
    ORANGE = "FE5716"
    AMBER = "FFD166"
    LGREY = "F0F0F0"
    DGREY = "888888"

    # Prepare data
    dfc = dfc.copy()
    dfc["_dt"] = dfc["Date"].apply(parse_to_sort_date)
    dfc = dfc.sort_values("_dt").reset_index(drop=True)
    amounts = dfc["Amount (£)"].astype(float).values
    dates = dfc["Date"].tolist()
    n = len(amounts)

    if n < 3:
        _hcell(ws, 1, 1, "Insufficient data for statistical analysis", bg=NAVY)
        ws.column_dimensions["A"].width = 50
        return

    # Headers
    headers = [
        "Metric",
        "Value",
        "Notes",
    ]
    for col, h in enumerate(headers, 1):
        _hcell(ws, 2, col, h, bg=NAVY)
    ws.row_dimensions[2].height = 28

    # Title
    tc = ws.cell(row=1, column=1, value="EDF ENERGY DISPUTE  —  STATISTICAL ANALYSIS")
    tc.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    tc.fill = PatternFill("solid", start_color=ORANGE)
    tc.border = CELL_BORDER
    tc.alignment = Alignment(horizontal="left", vertical="center")
    for c in [2, 3]:
        x = ws.cell(row=1, column=c)
        x.fill = PatternFill("solid", start_color=ORANGE)
        x.border = CELL_BORDER

    # Summary stats
    r = 3
    _section_hdr(ws, r, "DESCRIPTIVE STATISTICS")

    amounts_series = pd.Series(amounts)
    stats_data = [
        ("Count", len(amounts), "#,##0", "Number of billing records"),
        ("Mean (£)", float(amounts_series.mean()), "£#,##0.00", "Average balance"),
        ("Median (£)", float(amounts_series.median()), "£#,##0.00", "Median balance"),
        ("Std Dev (£)", float(amounts_series.std()), "£#,##0.00", "Standard deviation"),
        ("Min (£)", float(amounts_series.min()), "£#,##0.00", "Minimum balance"),
        ("Max (£)", float(amounts_series.max()), "£#,##0.00", "Maximum balance"),
        ("Range (£)", float(amounts_series.max() - amounts_series.min()), "£#,##0.00", "Max - Min"),
        (
            "Skewness",
            float(amounts_series.skew()) if hasattr(amounts_series, "skew") else 0,
            "0.00",
            "Asymmetry of distribution",
        ),
        (
            "Kurtosis",
            float(amounts_series.kurtosis()) if hasattr(amounts_series, "kurtosis") else 0,
            "0.00",
            "Tailedness of distribution",
        ),
        (
            "CV (%)",
            float(amounts_series.std() / amounts_series.mean() * 100)
            if amounts_series.mean() > 0
            else 0,
            "0.00",
            "Coefficient of variation",
        ),
    ]

    for label, value, fmt, note in stats_data:
        r += 1
        bg = LGREY if r % 2 == 0 else None
        _text(ws, r, 1, label, fill_hex=bg)
        if fmt == "£":
            _money(ws, r, 2, value, fill_hex=bg)
        elif fmt == "%":
            _num(ws, r, 2, value, fmt="0.0%", fill_hex=bg)
        else:
            _num(ws, r, 2, value, fmt=fmt, fill_hex=bg)
        _text(ws, r, 3, note, fill_hex=bg, color=DGREY)

    # Rolling statistics
    r += 1
    _section_hdr(ws, r, "ROLLING STATISTICS (6-period window)")
    r += 1
    _text(ws, r, 1, "Current 6-Period Rolling Mean (£)", bold=True)
    rolling_mean = float(pd.Series(amounts).rolling(6, min_periods=1).mean().iloc[-1])
    _money(ws, r, 2, rolling_mean)

    r += 1
    _text(ws, r, 1, "Current 6-Period Rolling Std (£)", bold=True)
    rolling_std = float(pd.Series(amounts).rolling(6, min_periods=1).std().iloc[-1])
    _money(ws, r, 2, rolling_std)

    r += 1
    _text(ws, r, 1, "Current 6-Period Rolling Min (£)", bold=True)
    rolling_min = float(pd.Series(amounts).rolling(6, min_periods=1).min().iloc[-1])
    _money(ws, r, 2, rolling_min)

    r += 1
    _text(ws, r, 1, "Current 6-Period Rolling Max (£)", bold=True)
    rolling_max = float(pd.Series(amounts).rolling(6, min_periods=1).max().iloc[-1])
    _money(ws, r, 2, rolling_max)

    r += 1
    _text(ws, r, 1, "Current 6-Period Rolling Median (£)", bold=True)
    rolling_median = float(pd.Series(amounts).rolling(6, min_periods=1).median().iloc[-1])
    _money(ws, r, 2, rolling_median)

    # Exponential Moving Average
    r += 1
    _section_hdr(ws, r, "EXPONENTIAL MOVING AVERAGE")
    r += 1
    _text(ws, r, 1, "Current EMA (span=6) (£)", bold=True)
    ema = float(pd.Series(amounts).ewm(span=6, adjust=False).mean().iloc[-1])
    _money(ws, r, 2, ema)

    r += 1
    _text(ws, r, 1, "EMA vs Simple SMA Difference (£)", bold=True)
    sma = float(pd.Series(amounts).rolling(6, min_periods=1).mean().iloc[-1])
    _money(ws, r, 2, ema - sma)

    # Momentum & Volatility
    r += 1
    _section_hdr(ws, r, "MOMENTUM & VOLATILITY")
    r += 1
    mom = float(pd.Series(amounts).diff(3).iloc[-1]) if n >= 4 else 0
    _text(ws, r, 1, "3-Period Momentum (£)", bold=True)
    _money(ws, r, 2, mom)

    r += 1
    vol = (
        float(pd.Series(amounts).pct_change().rolling(6, min_periods=1).std().iloc[-1])
        if n >= 3
        else 0
    )
    _text(ws, r, 1, "6-Period Volatility (σ of returns)", bold=True)
    _num(ws, r, 2, vol, fmt="0.00%")

    # Anomaly Detection
    r += 1
    _section_hdr(ws, r, "ANOMALY DETECTION")
    series = pd.Series(amounts, index=pd.to_datetime(dates, dayfirst=True, errors="coerce"))

    z_anoms = _zscore_anomalies(series, threshold=2.5)
    iqr_anoms = _iqr_anomalies(series, multiplier=1.5)

    z_count = int(z_anoms.sum())
    iqr_count = int(iqr_anoms.sum())

    r += 1
    _text(ws, r, 1, "Z-Score Anomalies (threshold=2.5σ)", bold=True)
    _num(ws, r, 2, z_count, fmt="#,##0")

    r += 1
    _text(ws, r, 1, "IQR Anomalies (multiplier=1.5)", bold=True)
    _num(ws, r, 2, iqr_count, fmt="#,##0")

    # List detected anomalies
    if z_count > 0:
        r += 1
        _text(ws, r, 1, "Z-Score Anomaly Dates:", bold=True)
        anom_dates = series[z_anoms].index
        for dt in anom_dates:
            r += 1
            amount_val = series[dt]
            if isinstance(amount_val, pd.Series):
                amount_val = amount_val.iloc[0]
            _text(
                ws,
                r,
                1,
                f"  • {dt.strftime('%d/%m/%Y') if hasattr(dt, 'strftime') else dt} ({amount_val:,.2f})",
            )

    if iqr_count > 0:
        r += 1
        _text(ws, r, 1, "IQR Anomaly Dates:", bold=True)
        anom_dates = series[iqr_anoms].index
        for dt in anom_dates:
            r += 1
            amount_val = series[dt]
            if isinstance(amount_val, pd.Series):
                amount_val = amount_val.iloc[0]
            _text(
                ws,
                r,
                1,
                f"  • {dt.strftime('%d/%m/%Y') if hasattr(dt, 'strftime') else dt} ({amount_val:,.2f})",
            )

    # Normality test (if scipy available)
    r += 1
    _section_hdr(ws, r, "DISTRIBUTION TESTS")
    if _HAS_SCIPY:
        try:
            from scipy import stats as sp_stats

            shapiro_stat, shapiro_p = sp_stats.shapiro(amounts_series.dropna())
            r += 1
            _text(ws, r, 1, "Shapiro-Wilk Test (Normality)", bold=True)
            _num(ws, r, 2, shapiro_stat, fmt="0.0000")
            _text(
                ws,
                r,
                3,
                f"p-value: {shapiro_p:.4f} — {'Normal' if shapiro_p > 0.05 else 'Non-normal'}",
            )

            # Jarque-Bera
            jb_stat, jb_p = sp_stats.jarque_bera(amounts_series.dropna())
            r += 1
            _text(ws, r, 1, "Jarque-Bera Test (Normality)", bold=True)
            _num(ws, r, 2, jb_stat, fmt="0.00")
            _text(ws, r, 3, f"p-value: {jb_p:.4f} — {'Normal' if jb_p > 0.05 else 'Non-normal'}")
        except Exception:
            r += 1
            _text(ws, r, 1, "Scipy tests failed", fill_hex=AMBER)
    else:
        r += 1
        _text(ws, r, 1, "Scipy not available — install for normality tests", fill_hex=AMBER)

    # Column widths
    for col_letter, width in zip(["A", "B", "C"], [45, 22, 80], strict=False):
        ws.column_dimensions[col_letter].width = width
    ws.freeze_panes = "A3"


__all__ = ["write_statistical_analysis_sheet"]
