"""Tests for payment/tariff/data quality analysis functions."""

import sys

sys.path.insert(0, "C:/Users/matthew/edf-bill-fetcher")

import numpy as np
import pandas as pd
import pytest

from edf_collector import (
    _analyze_tariff_impact,
    _data_quality_report,
    _detect_payment_patterns,
)


class TestPaymentDetection:
    """Tests for payment pattern detection."""

    def _make_df_with_payments(self):
        return pd.DataFrame(
            {
                "Date": pd.to_datetime(
                    [
                        "2024-01-01",
                        "2024-02-01",
                        "2024-03-01",
                        "2024-04-01",
                        "2024-05-01",
                        "2024-06-01",
                    ]
                ),
                "Amount (£)": [-200.0, 150.0, -200.0, 180.0, -210.0, 200.0],
                "Entry Type": ["Payment", "Charge", "Payment", "Charge", "Payment", "Charge"],
            }
        )

    def _make_df_no_payments(self):
        return pd.DataFrame(
            {
                "Date": pd.to_datetime(["2024-01-01", "2024-02-01"]),
                "Amount (£)": [150.0, 180.0],
                "Entry Type": ["Charge", "Charge"],
            }
        )

    def test_detect_payment_patterns_with_payments(self):
        df = self._make_df_with_payments()
        result = _detect_payment_patterns(df)
        assert isinstance(result, dict)
        assert "count" in result or len(result) > 0

    def test_detect_payment_patterns_no_payments(self):
        df = self._make_df_no_payments()
        result = _detect_payment_patterns(df)
        # Just verify function runs and returns something
        assert isinstance(result, dict)


class TestTariffAnalysis:
    """Tests for tariff impact analysis."""

    def _make_df_with_tariffs(self):
        return pd.DataFrame(
            {
                "Date": pd.to_datetime(["2024-01-01", "2024-04-01", "2024-07-01", "2024-10-01"]),
                "Amount (£)": [200.0, 250.0, 300.0, 350.0],
                "Period Charge (£)": [180.0, 220.0, 260.0, 300.0],
                "Units (kWh)": [500.0, 550.0, 600.0, 650.0],
                "Tariff": ["Standard", "Standard", "Fixed", "Fixed"],
                "Standing Chg (p/day)": [25.0, 25.0, 28.0, 28.0],
            }
        )

    def test_analyze_tariff_impact_with_data(self):
        df = self._make_df_with_tariffs()
        result = _analyze_tariff_impact(df)
        assert isinstance(result, dict)

    def test_analyze_tariff_impact_empty_df(self):
        df = pd.DataFrame()
        result = _analyze_tariff_impact(df)
        assert isinstance(result, dict)


class TestDataQuality:
    """Tests for data quality reporting."""

    def _make_complete_df(self):
        return pd.DataFrame(
            {
                "Date": pd.to_datetime(["2024-01-01"] * 5),
                "Source": ["Email", "PDF", "HTM", "Email", "PDF"],
                "Amount (£)": [100.0, 200.0, 150.0, 250.0, 300.0],
                "Period From": ["01/01/2024"] * 5,
                "Period To": ["31/01/2024"] * 5,
                "Unit Rate (p/kWh)": [25.0] * 5,
            }
        )

    def _make_incomplete_df(self):
        return pd.DataFrame(
            {
                "Date": pd.to_datetime(["2024-01-01", "2024-01-01"] + [pd.NaT] * 3),
                "Source": ["Email", "PDF"] + [None] * 3,
                "Amount (£)": [100.0, 200.0, np.nan, 250.0, 300.0],
                "Period From": ["01/01/2024", "01/02/2024", "N/A", "N/A", "N/A"],
                "Period To": ["31/01/2024", "28/02/2024", "N/A", "N/A", "N/A"],
                "Unit Rate (p/kWh)": [25.0, 25.0, "N/A", "N/A", "N/A"],
            }
        )

    def test_data_quality_report_complete(self):
        df = self._make_complete_df()
        result = _data_quality_report(df)
        assert isinstance(result, dict)
        # Keys may vary, just verify it runs

    def test_data_quality_report_incomplete(self):
        df = self._make_incomplete_df()
        result = _data_quality_report(df)
        assert isinstance(result, dict)
        # Verify it runs without crashing


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
