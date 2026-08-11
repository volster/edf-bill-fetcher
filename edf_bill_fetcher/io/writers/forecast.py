"""Forecast sheet writer — extracted from writers/__init__.py.

Contains: write_forecast_sheet — renders the "Forecast" worksheet with
linear and Holt-Winters forecasts + forecast chart.
"""

from __future__ import annotations

from datetime import timedelta

import numpy as np
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
from edf_bill_fetcher.processors.forecasting import HAS_STATSMODELS, _compute_ema  # noqa: F401
from edf_bill_fetcher.writers._helpers import (
    _holt_winters_forecast_pair,
    _linear_forecast,
    _linear_forecast_pair,
)

# --- write_forecast_sheet (was writers/__init__.py L2175-2422) ---


def write_forecast_sheet(ws, dfc):
    """Write Forecast/Projection tab with multiple forecasting methods."""
    ws.title = "Forecast & Projection"

    NAVY = "10367A"
    ORANGE = "FE5716"
    AMBER = "FFD166"
    LGREY = "F0F0F0"
    DGREY = "888888"

    dfc = dfc.copy()
    dfc["_dt"] = dfc["Date"].apply(parse_to_sort_date)
    dfc = dfc.sort_values("_dt").reset_index(drop=True)
    amounts = dfc["Amount (£)"].astype(float).values
    dates = dfc["Date"].tolist()
    n = len(amounts)

    if n < 3:
        _hcell(ws, 1, 1, "Insufficient data for forecasting (need 3+ records)", bg=NAVY)
        ws.column_dimensions["A"].width = 60
        return

    # ``Date`` + the canonical six forecast columns + ``Forecast Δ
    # (Actual − Linear)``.  The Δ column is what makes the tab
    # useful as evidence: a reviewer sees *by how much* each bill
    # diverged from what the model would call average.  Historical
    # rows carry a per-row back-painted prediction; future rows
    # carry forward-looking projections; the divider between the
    # two is a separator row.
    headers = [
        "Date",
        "Actual (£)",
        "Linear Forecast (£)",
        "Holt-Winters (£)",
        "EMA Projection (£)",
        "Confidence (±£)",
        "Forecast Δ (Actual − Linear)",
    ]
    for col, h in enumerate(headers, 1):
        _hcell(ws, 2, col, h, bg=NAVY)
    ws.row_dimensions[2].height = 28

    tc = ws.cell(row=1, column=1, value="EDF ENERGY DISPUTE  —  BALANCE FORECAST")
    tc.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    tc.fill = PatternFill("solid", start_color=ORANGE)
    tc.border = CELL_BORDER
    tc.alignment = Alignment(horizontal="left", vertical="center")
    for c in range(2, 8):
        x = ws.cell(row=1, column=c)
        x.fill = PatternFill("solid", start_color=ORANGE)
        x.border = CELL_BORDER

    # Generate forecasts (6 steps ahead).  We use the *_pair helper
    # variants to also obtain the in-sample fitted-values array so
    # every historical row carries a real prediction column rather
    # than the previous "—" placeholders.  This is what makes the
    # tab show model-vs-actual divergence across the full data range.
    forecast_steps = 6
    series = pd.Series(amounts, index=pd.to_datetime(dates, dayfirst=True, errors="coerce"))

    # ``linear_fitted[i]`` is the straight-line prediction at row i
    # (uses ALL n historical points); ``linear_fc[i]`` is the future
    # value i steps past the last historical row.  Both come from
    # the same fit, so the in-sample and out-of-sample columns
    # share one model.
    linear_fitted, linear_fc = _linear_forecast_pair(series, forecast_steps)
    hw_fitted, hw_fc = _holt_winters_forecast_pair(series, forecast_steps)
    # EMA trajectory: per-row exponentially-weighted moving average.
    # We expand the existing ``_compute_ema`` helper into a length-n
    # series so every historical row gets the right EMA *as of that
    # row*, not the last-window mean.
    ema_series = _compute_ema(series, span=6)
    ema_last = ema_series.iloc[-1] if n >= 2 else amounts[-1]
    # Forward EMA projection extends the last EMA flat-forecast for
    # future rows; historical rows just carry the historical EMA.
    ema_future = [ema_last] * forecast_steps

    # Historical volatility for confidence intervals.
    # ``hist_vol`` is the std-dev of monthly *returns* (pct_change),
    # which is what we multiply against the predicted value to
    # produce a ±2σ confidence band.  With only one historical bill
    # we fall back to a sensible default.
    returns = pd.Series(amounts).pct_change().dropna()
    hist_vol = returns.std() if len(returns) > 1 else 0.05

    def _model_value(fitted_array, fc_array, i, n_total):
        """Pick the in-sample fitted value at historical index i or ``N/A`` if the model didn't fit (not enough data)."""
        if fitted_array is None:
            return None
        # Defensive index guard — the fitted array has the same
        # length as ``series`` per the *_pair helpers, but a
        # statsmodels-index misalignment is always possible.
        if i < len(fitted_array):
            val = fitted_array[i]
            return val if not pd.isna(val) else None
        return None

    # === Historical block: back-paint every forecast column ===
    # The y-axis of the forecast table now spans the *entire* data
    # range — each historical row carries the model's prediction at
    # that point, and the Forecast Δ column quantifies how far the
    # actual bill landed above (positive) or below (negative) the
    # linear-trend baseline.  The future block (after the separator
    # row) shows 6 forward projection rows.  Together they answer
    # "given what you've paid historically, what should you have
    # paid each month, and where did the bill diverge?".
    r = 3
    for i in range(n):
        bg = LGREY if i % 2 == 0 else None
        _text(ws, r, 1, dates[i], fill_hex=bg)
        _money(ws, r, 2, float(amounts[i]), fill_hex=bg)
        # Linear forecast — back-painted fitted value (not "—").
        lin_val = _model_value(linear_fitted, linear_fc, i, n)
        if lin_val is not None:
            _money(ws, r, 3, float(lin_val), fill_hex=bg)
        else:
            _text(ws, r, 3, "N/A", fill_hex=bg)
        # Holt-Winters — back-painted fitted value (still "N/A"
        # when statsmodels is unavailable or the series is too
        # short for the additive-trend fit).
        hw_val = _model_value(hw_fitted, hw_fc, i, n)
        if hw_val is not None:
            _money(ws, r, 4, float(hw_val), fill_hex=bg)
        else:
            _text(ws, r, 4, "N/A", fill_hex=bg)
        # EMA — per-row exponentially-weighted moving average
        # (historical anchored to row i's position in the series).
        ema_at_i = float(ema_series.iloc[i]) if not pd.isna(ema_series.iloc[i]) else None
        if ema_at_i is not None:
            _money(ws, r, 5, ema_at_i, fill_hex=bg)
        else:
            _text(ws, r, 5, "N/A", fill_hex=bg)
        # Confidence band — ±2σ around the fitted value.  When the
        # model didn't fit we fall back to the predicted value of
        # the actual bill (i.e. confidence = 0) — visually faithful
        # but not concealing data.
        if lin_val is not None:
            conf = abs(float(lin_val)) * hist_vol * 2
            _money(ws, r, 6, conf, fill_hex=bg)
        else:
            _text(ws, r, 6, "N/A", fill_hex=bg)
        # Forecast Δ = actual − fitted linear.  This is the
        # ombudsman-facing signal: a row with ``£50`` actual and a
        # fitted linear value of ``£200`` writes ``−£150`` here,
        # i.e. the bill landed £150 below what the trend expected
        # (favourable).  Conversely an actual bill above fitted
        # writes a positive number the reviewer can see as the
        # over-billing flag.
        if lin_val is not None:
            delta = float(amounts[i]) - float(lin_val)
            _money(ws, r, 7, delta, fill_hex=bg)
        else:
            _text(ws, r, 7, "N/A", fill_hex=bg)
        r += 1

    # Separator
    ws.cell(row=r, column=1, value="— " * 20).font = Font(bold=True, color=DGREY)
    r += 1

    # === Forward forecast block: 6 steps past the last historical ===
    forecast_dates = []
    last_date = parse_to_sort_date(dates[-1])

    if not pd.isna(last_date):
        for i in range(1, forecast_steps + 1):
            next_date = last_date + timedelta(days=30 * i)  # Approximate monthly
            forecast_dates.append(next_date.strftime("%d/%m/%Y"))
    else:
        forecast_dates = [f"Forecast +{i + 1}" for i in range(forecast_steps)]

    for i in range(forecast_steps):
        bg = AMBER
        _text(ws, r, 1, forecast_dates[i], fill_hex=bg, bold=True)
        _text(ws, r, 2, "—", fill_hex=bg)  # No actual
        lin_val = linear_fc[i] if linear_fc is not None else None
        hw_val = hw_fc[i] if hw_fc is not None else None
        if lin_val is not None:
            _money(ws, r, 3, float(lin_val), fill_hex=bg)
        else:
            _text(ws, r, 3, "N/A", fill_hex=bg)
        if hw_val is not None:
            _money(ws, r, 4, float(hw_val), fill_hex=bg)
        else:
            _text(ws, r, 4, "N/A", fill_hex=bg)
        _money(ws, r, 5, ema_future[i], fill_hex=bg)
        # Confidence band on the future prediction is the *predicted
        # value's* ±2σ — same shape as on the historical rows but
        # at the forecasted level so the reviewer sees the
        # widening band as the horizon extends.
        if lin_val is not None:
            conf = abs(float(lin_val)) * hist_vol * 2
            _money(ws, r, 6, conf, fill_hex=bg)
        else:
            _text(ws, r, 6, "N/A", fill_hex=bg)
        # Forecast Δ is intentionally "—" for future rows: there
        # is no actual bill yet to subtract from.
        _text(ws, r, 7, "—", fill_hex=bg)
        r += 1

    # Model comparison
    r += 1
    _section_hdr(ws, r, "MODEL COMPARISON")
    r += 1
    _text(ws, r, 1, "Linear Trend", bold=True)
    _text(ws, r, 2, "Simple linear regression on time index")
    r += 1
    _text(ws, r, 1, "Holt-Winters", bold=True)
    _text(
        ws, r, 2, "Exponential smoothing with trend" + (" + seasonality" if HAS_STATSMODELS else "")
    )
    r += 1
    _text(ws, r, 1, "EMA Projection", bold=True)
    _text(ws, r, 2, "Extends last Exponential Moving Average (span=6)")
    r += 1
    _text(ws, r, 1, "Historical Volatility", bold=True)
    _num(ws, r, 2, hist_vol, fmt="0.00%")
    _text(ws, r, 3, "Monthly return std used for confidence bands")

    # Accuracy metrics (in-sample)
    r += 1
    _section_hdr(ws, r, "IN-SAMPLE ACCURACY (Last 6 periods)")
    if n >= 7:
        test_series = pd.Series(amounts[:-6])
        true_vals = amounts[-6:]
        lin_hist = _linear_forecast(test_series, 6)
        if lin_hist is not None:
            mae = np.mean(np.abs(lin_hist - true_vals))
            rmse = np.sqrt(np.mean((lin_hist - true_vals) ** 2))
            mape = np.mean(np.abs((lin_hist - true_vals) / true_vals)) * 100

            r += 1
            _text(ws, r, 1, "Linear Forecast MAE (£)", bold=True)
            _money(ws, r, 2, mae)
            r += 1
            _text(ws, r, 1, "Linear Forecast RMSE (£)", bold=True)
            _money(ws, r, 2, rmse)
            r += 1
            _text(ws, r, 1, "Linear Forecast MAPE (%)", bold=True)
            _num(ws, r, 2, mape, fmt="0.00%")

    for col_letter, width in zip(
        ["A", "B", "C", "D", "E", "F", "G"], [14, 16, 18, 18, 18, 16, 22], strict=False
    ):
        ws.column_dimensions[col_letter].width = width
    ws.freeze_panes = "A3"


__all__ = ["write_forecast_sheet"]
