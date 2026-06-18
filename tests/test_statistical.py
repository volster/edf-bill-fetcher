"""Tests for statistical analysis functions to improve coverage."""

import sys

sys.path.insert(0, "C:/Users/matthew/edf-bill-fetcher")

import pandas as pd
import pytest

from edf_collector import (
    _compute_ema,
    _compute_momentum,
    _compute_rolling_stats,
    _compute_volatility,
    _holt_winters_forecast,
    _iqr_anomalies,
    _linear_forecast,
    _zscore_anomalies,
)


class TestStatisticalFunctions:
    """Tests for statistical analysis helper functions."""

    @pytest.fixture
    def sample_series(self):
        return pd.Series([100.0, 200.0, 150.0, 300.0, 250.0, 400.0, 350.0, 500.0])

    @pytest.fixture
    def flat_series(self):
        return pd.Series([100.0] * 8)

    @pytest.fixture
    def anomaly_series(self):
        # Has clear outliers
        return pd.Series([100.0, 105.0, 95.0, 110.0, 1000.0, 90.0, 102.0, 98.0])

    def test_compute_rolling_stats_basic(self, sample_series):
        result = _compute_rolling_stats(sample_series, window=3)
        assert "mean" in result
        assert "std" in result
        assert "min" in result
        assert "max" in result
        assert "median" in result
        assert len(result["mean"]) == len(sample_series)

    def test_compute_rolling_stats_short_window(self, sample_series):
        result = _compute_rolling_stats(sample_series, window=2)
        assert len(result["mean"]) == len(sample_series)

    def test_compute_rolling_stats_flat_series(self, flat_series):
        result = _compute_rolling_stats(flat_series, window=3)
        assert (result["std"].dropna() == 0).all()

    def test_compute_ema_basic(self, sample_series):
        result = _compute_ema(sample_series, span=3)
        assert len(result) == len(sample_series)
        assert result.iloc[0] == sample_series.iloc[0]

    def test_compute_ema_flat_series(self, flat_series):
        result = _compute_ema(flat_series, span=3)
        assert (result == 100.0).all()

    def test_compute_momentum_basic(self, sample_series):
        result = _compute_momentum(sample_series, period=2)
        assert len(result) == len(sample_series)
        assert pd.isna(result.iloc[0])

    def test_compute_momentum_flat_series(self, flat_series):
        result = _compute_momentum(flat_series, period=2)
        assert (result.dropna() == 0).all()

    def test_compute_volatility_basic(self, sample_series):
        result = _compute_volatility(sample_series, window=3)
        assert len(result) == len(sample_series)
        assert pd.isna(result.iloc[0]) or result.iloc[0] >= 0

    def test_compute_volatility_flat_series(self, flat_series):
        result = _compute_volatility(flat_series, window=3)
        assert (result.dropna() == 0).all()

    def test_zscore_anomalies_no_outliers(self, sample_series):
        result = _zscore_anomalies(sample_series, threshold=2.5)
        assert isinstance(result, pd.Series)
        assert len(result) == len(sample_series)

    def test_zscore_anomalies_with_outliers(self, anomaly_series):
        result = _zscore_anomalies(anomaly_series, threshold=2.0)
        assert bool(result.iloc[4]) is True

    def test_zscore_anomalies_flat_series(self, flat_series):
        result = _zscore_anomalies(flat_series, threshold=2.0)
        assert not result.any()

    def test_iqr_anomalies_with_outliers(self, anomaly_series):
        result = _iqr_anomalies(anomaly_series, multiplier=1.5)
        assert bool(result.iloc[4]) is True

    def test_linear_forecast_basic(self, sample_series):
        result = _linear_forecast(sample_series, steps=3)
        assert len(result) == 3
        assert all(isinstance(x, (int, float)) for x in result)

    def test_linear_forecast_more_steps(self, sample_series):
        result = _linear_forecast(sample_series, steps=6)
        assert len(result) == 6

    def test_holt_winters_forecast_basic(self, sample_series):
        result = _holt_winters_forecast(sample_series, steps=3, seasonal_periods=None)
        assert result is None or len(result) == 3

    def test_holt_winters_forecast_with_seasonal(self, sample_series):
        result = _holt_winters_forecast(sample_series, steps=3, seasonal_periods=2)
        assert result is None or len(result) == 3


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
