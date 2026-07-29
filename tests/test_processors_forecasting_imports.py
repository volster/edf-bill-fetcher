"""Tests that forecasting helper functions are importable from the processors.forecasting submodule.

All tests are RED at Phase 0 because ``edf_bill_fetcher.processors.forecasting``
does not yet exist.
"""

from __future__ import annotations


def test_holt_winters_forecast_importable() -> None:
    from edf_bill_fetcher.processors.forecasting import _holt_winters_forecast

    assert _holt_winters_forecast is not None


def test_linear_forecast_importable() -> None:
    from edf_bill_fetcher.processors.forecasting import _linear_forecast

    assert _linear_forecast is not None


def test_compute_ema_importable() -> None:
    from edf_bill_fetcher.processors.forecasting import _compute_ema

    assert _compute_ema is not None


def test_compute_momentum_importable() -> None:
    from edf_bill_fetcher.processors.forecasting import _compute_momentum

    assert _compute_momentum is not None


def test_compute_rolling_stats_importable() -> None:
    from edf_bill_fetcher.processors.forecasting import _compute_rolling_stats

    assert _compute_rolling_stats is not None


def test_compute_volatility_importable() -> None:
    from edf_bill_fetcher.processors.forecasting import _compute_volatility

    assert _compute_volatility is not None


def test_zscore_anomalies_importable() -> None:
    from edf_bill_fetcher.processors.forecasting import _zscore_anomalies

    assert _zscore_anomalies is not None


def test_iqr_anomalies_importable() -> None:
    from edf_bill_fetcher.processors.forecasting import _iqr_anomalies

    assert _iqr_anomalies is not None
