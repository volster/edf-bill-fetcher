"""Forecast / anomaly detection helpers for the evidence workbook.

Extracted from ``edf_collector.py`` as part of the modularization refactor
(Task 5 — Phase 4).  Pure-pandas helpers with optional ``statsmodels``
fallback for Holt-Winters exponential smoothing.
"""

from __future__ import annotations

import importlib.util

import numpy as np
import pandas as pd

HAS_STATSMODELS = importlib.util.find_spec("statsmodels.tsa.holtwinters") is not None

try:
    from statsmodels.tsa.holtwinters import ExponentialSmoothing  # type: ignore[import-not-found]
except ImportError:
    ExponentialSmoothing = None  # type: ignore[assignment,misc]

from edf_bill_fetcher.helpers.date_utils import (  # noqa: E402,I001
    compute_ema as _compute_ema,
)
from edf_bill_fetcher.helpers.date_utils import (  # noqa: E402,I001
    compute_momentum as _compute_momentum,
)
from edf_bill_fetcher.helpers.date_utils import (  # noqa: E402,I001
    compute_rolling_stats as _compute_rolling_stats,
)
from edf_bill_fetcher.writers._helpers import _zscore_anomalies  # noqa: E402,I001


def _compute_volatility(series, window=6):
    """Compute rolling volatility (std of returns)."""
    returns = series.pct_change()
    return returns.rolling(window=window, min_periods=1).std()


def _iqr_anomalies(series, multiplier=1.5):
    """Detect anomalies using IQR method."""
    if len(series) < 4:
        return pd.Series(False, index=series.index)
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    if iqr == 0:
        return pd.Series(False, index=series.index)
    lower = q1 - multiplier * iqr
    upper = q3 + multiplier * iqr
    return (series < lower) | (series > upper)


def _linear_forecast_pair(series, steps=6):
    """Compute a simple linear regression and return (fitted, future) values.

    The fitted series is the model's prediction at each historical
    point — this lets the Forecast tab back-paint predictions onto
    historical rows so the reader sees actual-vs-predicted for the
    whole data range, not only at a 6-step future horizon.

    Linear regression in this codebase uses ``np.polyfit``.  The
    fitted value at index ``i`` is simply ``np.polyval(coeffs, i)``
    computed against the same coefficients used for the future
    forecast, so the in-sample and out-of-sample predictions share
    a single model — meaning the historical vs forward columns
    reflect exactly the same fit.

    Returns ``(None, None)`` for insufficient data.
    """
    if len(series) < 3:
        return None, None
    x = np.arange(len(series))
    y = series.values
    # Handle NaN values
    mask = ~np.isnan(y)
    if mask.sum() < 3:
        return None, None
    x_clean = x[mask]
    y_clean = y[mask]
    try:
        coeffs = np.polyfit(x_clean, y_clean, 1)
        # Fitted values for every historical index — back-pained
        # by the same straight line that drives the future window.
        fitted = np.polyval(coeffs, x)
        future_x = np.arange(len(series), len(series) + steps)
        forecast = np.polyval(coeffs, future_x)
        return fitted, forecast
    except Exception:
        return None, None


def _holt_winters_forecast_pair(series, steps=6, seasonal_periods=None):
    """Holt-Winters: returns (fitted, future) values (if statsmodels available).

    Mirrors ``_linear_forecast_pair`` for the ExponentialSmoothing
    path.  Statsmodels's ``fit()`` returns a fitted-ness model whose
    ``.fittedvalues`` attribute carries the one-step-ahead in-sample
    prediction at every historical index — exactly what we need to
    back-paint the forecast tab so the reader sees actual vs
    predicted divergence for the whole data range.

    Returns ``(None, None)`` when statsmodels is unavailable, the
    series is too short, or fitting fails.
    """
    if not HAS_STATSMODELS or len(series) < 4:
        return None, None
    try:
        clean_series = series.dropna()
        if len(clean_series) < 4:
            return None, None

        if seasonal_periods is None:
            seasonal_periods = min(12, len(clean_series) // 2) if len(clean_series) >= 8 else None

        model = ExponentialSmoothing(
            clean_series,
            trend="add",
            seasonal="add" if seasonal_periods else None,
            seasonal_periods=seasonal_periods,
            initialization_method="estimated",
        )
        fitted_model = model.fit(optimized=True)
        # In-sample fitted: statsmodels returns the one-step-ahead
        # prediction for each historical point the model was fit
        # against.  We reindex onto the original series (which may
        # include NaN gaps) so row N in the call sites lines up
        # with row N in the user's data.
        fitted_vals = fitted_model.fittedvalues.reindex(series.index)
        forecast = fitted_model.forecast(steps).values
        return fitted_vals.values, forecast
    except Exception:
        return None, None


def _linear_forecast(series, steps=6):
    """Produce a forward-only linear regression forecast (legacy entry point).

    See ``_linear_forecast_pair`` for the (fitted, future) form that
    the Forecast tab now uses.  This single-value shim is kept for
    any callers that imported the previous-shape return value (we
    don't have any in-tree callers anymore, but a user
    may have downstream code that does).
    """
    _, forecast = _linear_forecast_pair(series, steps)
    return forecast


def _holt_winters_forecast(series, steps=6, seasonal_periods=None):
    """Holt-Winters forward-only legacy entry point.  See ``_holt_winters_forecast_pair``."""
    _, forecast = _holt_winters_forecast_pair(series, steps, seasonal_periods)
    return forecast


__all__ = [
    "_compute_volatility",
    "_zscore_anomalies",
    "_iqr_anomalies",
    "_linear_forecast_pair",
    "_holt_winters_forecast_pair",
    "_linear_forecast",
    "_holt_winters_forecast",
    "_compute_ema",
    "_compute_momentum",
    "_compute_rolling_stats",
]
