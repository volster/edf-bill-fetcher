"""D3: processors/forecasting is the single home for forecast/anomaly primitives."""

from __future__ import annotations


def test_forecasting_module_does_not_import_writers() -> None:
    import edf_bill_fetcher.processors.forecasting as fp

    src = fp.__file__
    assert src is not None
    with open(src, encoding="utf-8") as fh:
        text = fh.read()
    assert "edf_bill_fetcher.writers" not in text


def test_helpers_aliases_resolve_to_canonical() -> None:
    from edf_bill_fetcher.processors import forecasting as fp
    from edf_bill_fetcher.writers import _helpers as wh

    # Identity check: the alias *is* the canonical function object.
    assert wh._linear_forecast_pair is fp._linear_forecast_pair
    assert wh._holt_winters_forecast_pair is fp._holt_winters_forecast_pair
    assert wh._linear_forecast is fp._linear_forecast
    assert wh._holt_winters_forecast is fp._holt_winters_forecast
    assert wh._compute_volatility is fp._compute_volatility
    assert wh._zscore_anomalies is fp._zscore_anomalies
    assert wh._iqr_anomalies is fp._iqr_anomalies


def test_sanitised_returns_drops_nonfinite() -> None:
    import numpy as np
    import pandas as pd

    from edf_bill_fetcher.processors.forecasting import _sanitised_returns

    # A zero amount forces pct_change to divide by zero -> inf. The finite
    # returns around it survive; the non-finite division is dropped.
    series = pd.Series([100.0, 0.0, 150.0, 120.0])
    out = _sanitised_returns(series)
    assert not np.isinf(out).any()
    assert not out.isna().any()
    assert len(out) == 2
