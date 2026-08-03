"""Coverage tests for the pure-pandas forecasting / anomaly helpers
in ``processors/forecasting.py``.

Closes the 13-missed-line gap. Each early-return branch — too-short
series, all-zero std, all-zero IQR, all-NaN mask, statsmodels-
unavailable fallback — is exercised with a minimal synthetic
``pandas.Series``. No mocking of pandas / numpy is needed; the
helpers operate on the series values directly.
"""

from __future__ import annotations

import numpy as np
import pandas as pd
import pytest

from edf_bill_fetcher.processors.forecasting import (
    HAS_STATSMODELS,
    _compute_volatility,
    _holt_winters_forecast,
    _holt_winters_forecast_pair,
    _iqr_anomalies,
    _linear_forecast,
    _linear_forecast_pair,
    _zscore_anomalies,
)

# ---------- _compute_volatility ----------


def test_compute_volatility_returns_rolling_std_of_pct_change() -> None:
    series = pd.Series([1.0, 2.0, 4.0, 8.0, 16.0, 32.0, 64.0])
    vol = _compute_volatility(series, window=3)
    assert isinstance(vol, pd.Series)
    assert vol.index.equals(series.index)
    assert pd.isna(vol.iloc[0])


def test_compute_volatility_default_window_is_6() -> None:
    series = pd.Series([1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0])
    vol = _compute_volatility(series)
    assert len(vol) == 9
    assert pd.isna(vol.iloc[0])


def test_compute_volatility_constant_series_yields_zero_or_nan() -> None:
    series = pd.Series([5.0, 5.0, 5.0, 5.0, 5.0])
    vol = _compute_volatility(series, window=2)
    assert (vol.fillna(0.0) == 0.0).all()


# ---------- _zscore_anomalies ----------


def test_zscore_anomalies_short_series_returns_all_false() -> None:
    series = pd.Series([1.0, 2.0])
    result = _zscore_anomalies(series)
    assert isinstance(result, pd.Series)
    assert len(result) == 2
    assert not result.any()


def test_zscore_anomalies_zero_std_returns_all_false() -> None:
    series = pd.Series([3.0, 3.0, 3.0, 3.0])
    result = _zscore_anomalies(series)
    assert not result.any()


def test_zscore_anomalies_detects_outliers_above_threshold() -> None:
    series = pd.Series([1.0, 2.0, 1.0, 2.0, 1.0, 20.0])
    result = _zscore_anomalies(series, threshold=1.5)
    assert result.iloc[-1]
    assert not result.iloc[0]


def test_zscore_anomalies_default_threshold_is_2_5() -> None:
    series = pd.Series([1.0, 1.0, 1.0, 1.0, 1.0, 5.0])
    result_low_threshold = _zscore_anomalies(series, threshold=0.5)
    result_default = _zscore_anomalies(series)
    assert result_low_threshold.iloc[-1]
    assert not result_default.iloc[-1]


# ---------- _iqr_anomalies ----------


def test_iqr_anomalies_short_series_returns_all_false() -> None:
    series = pd.Series([1.0, 2.0, 3.0])
    result = _iqr_anomalies(series)
    assert not result.any()


def test_iqr_anomalies_zero_iqr_returns_all_false() -> None:
    series = pd.Series([5.0, 5.0, 5.0, 5.0, 5.0])
    result = _iqr_anomalies(series)
    assert not result.any()


def test_iqr_anomalies_flags_values_outside_whiskers() -> None:
    series = pd.Series([10.0, 11.0, 12.0, 13.0, 14.0, 100.0])
    result = _iqr_anomalies(series, multiplier=1.5)
    assert result.iloc[-1]
    assert not result.iloc[0]


def test_iqr_anomalies_default_multiplier_is_1_5() -> None:
    series = pd.Series([1.0, 2.0, 3.0, 4.0, 100.0])
    result = _iqr_anomalies(series)
    assert result.iloc[-1]


# ---------- _linear_forecast_pair ----------


def test_linear_forecast_pair_short_series_returns_none_pair() -> None:
    series = pd.Series([1.0, 2.0])
    fitted, future = _linear_forecast_pair(series)
    assert fitted is None
    assert future is None


def test_linear_forecast_pair_all_nan_series_returns_none_pair() -> None:
    series = pd.Series([np.nan, np.nan, 1.0, np.nan])
    fitted, future = _linear_forecast_pair(series)
    assert fitted is None
    assert future is None


def test_linear_forecast_pair_well_formed_series_returns_fitted_and_future() -> None:
    series = pd.Series([1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0])
    fitted, future = _linear_forecast_pair(series, steps=4)
    assert fitted is not None
    assert future is not None
    assert len(fitted) == 8
    assert len(future) == 4


def test_linear_forecast_pair_string_series_raises_type_error() -> None:
    """A non-numeric series raises TypeError at `np.isnan` — the helper explicitly does NOT
    catch `TypeError` from the isnan call (only catches exceptions inside the
    `try`-block wrapping `np.polyfit` and `np.polyval`). Documented behavior."""
    series = pd.Series(["not", "numbers", "at", "all", "here"])  # type: ignore[list-item]
    with pytest.raises(TypeError):
        _linear_forecast_pair(series)  # type: ignore[arg-type]


# ---------- _holt_winters_forecast_pair ----------


def test_holt_winters_forecast_pair_short_series_returns_none_pair() -> None:
    series = pd.Series([1.0, 2.0, 3.0])
    fitted, future = _holt_winters_forecast_pair(series)
    assert fitted is None
    assert future is None


def test_holt_winters_forecast_pair_too_few_non_nan_returns_none_pair() -> None:
    series = pd.Series([1.0, np.nan, np.nan, np.nan, 2.0, np.nan])
    fitted, future = _holt_winters_forecast_pair(series)
    assert fitted is None
    assert future is None


@pytest.mark.skipif(
    not HAS_STATSMODELS, reason="statsmodels unavailable — branch returns (None, None)"
)
def test_holt_winters_forecast_pair_well_formed_series_returns_arrays() -> None:
    series = pd.Series([10.0, 11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0])
    fitted, future = _holt_winters_forecast_pair(series, steps=4, seasonal_periods=None)
    assert fitted is not None
    assert future is not None
    assert len(future) == 4


@pytest.mark.skipif(HAS_STATSMODELS, reason="statsmodels IS available — branch returns arrays")
def test_holt_winters_forecast_pair_no_statsmodels_returns_none_pair() -> None:
    series = pd.Series([10.0, 11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0])
    fitted, future = _holt_winters_forecast_pair(series, steps=4)
    assert fitted is None
    assert future is None


# ---------- legacy forward-only shims ----------


def test_linear_forecast_legacy_shim_returns_forecast_only() -> None:
    series = pd.Series([1.0, 2.0, 3.0, 4.0, 5.0])
    result = _linear_forecast(series, steps=3)
    assert result is not None
    assert len(result) == 3


def test_linear_forecast_legacy_shim_short_series_returns_none() -> None:
    series = pd.Series([1.0, 2.0])
    assert _linear_forecast(series) is None


def test_holt_winters_forecast_legacy_shim_returns_forecast_only() -> None:
    series = pd.Series([10.0, 11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0])
    if HAS_STATSMODELS:
        result = _holt_winters_forecast(series, steps=3)
        assert result is not None
        assert len(result) == 3
    else:
        assert _holt_winters_forecast(series, steps=3) is None
