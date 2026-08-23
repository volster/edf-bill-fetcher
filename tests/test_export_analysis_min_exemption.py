# tests/test_export_analysis_min_exemption.py
import warnings

import pandas as pd

from edf_bill_fetcher.io.writers.export import _prepare_analysis_frame
from edf_bill_fetcher.models.config import ConfigDict


def test_prepare_analysis_frame_exists():
    """The inline gate logic at export.py:807-811 must be extracted into a
    callable helper so legal-candidate handling can be tested in isolation."""
    assert callable(_prepare_analysis_frame)


def test_analysis_min_exempt_for_span_over_365_days():
    df_an = pd.DataFrame(
        [
            {
                "Invoice #": "KI-0014",
                "Date": "2022-01-02",
                "Period From": "2020-01-01",
                "Period To": "2020-01-02",
                "Amount (£)": 100.0,  # below analysis_min £500
                "Period Charge (£)": 1000.0,  # substantial period charge
                "Entry Type": "New Bill",
            }
        ]
    )
    config: ConfigDict = {"analysis_min": 500.0}
    dfc = _prepare_analysis_frame(df_an, config)

    # KI-0014 should NOT be dropped (legal-candidate handling overrides the amount gate)
    assert len(dfc) == 1
    assert dfc.iloc[0]["Invoice #"] == "KI-0014"


def test_analysis_min_still_filters_short_period_low_amount():
    """A short-period invoice below analysis_min should still be dropped."""
    df_an = pd.DataFrame(
        [
            {
                "Invoice #": "T99",
                "Date": "2021-04-01",
                "Period From": "2021-01-01",
                "Period To": "2021-03-31",
                "Amount (£)": 100.0,  # below analysis_min £500
                "Period Charge (£)": 50.0,
                "Entry Type": "New Bill",
            }
        ]
    )
    config: ConfigDict = {"analysis_min": 500.0}
    dfc = _prepare_analysis_frame(df_an, config)
    assert len(dfc) == 0


def test_unparseable_dates_warn():
    """A New Bill row with unparseable Date/Period From warns about the drop."""
    df_an = pd.DataFrame(
        [
            {
                "Invoice #": "KI-BAD",
                "Date": "not-a-date",
                "Period From": "also-not-a-date",
                "Period To": "2021-03-31",
                "Amount (£)": 100.0,  # below analysis_min £500
                "Period Charge (£)": 50.0,
                "Entry Type": "New Bill",
            }
        ]
    )
    config: ConfigDict = {"analysis_min": 500.0}
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        dfc = _prepare_analysis_frame(df_an, config)

    assert len(dfc) == 0
    assert any("unparseable dates" in str(w.message) for w in caught)


def test_payment_below_min_kept() -> None:
    df_an = pd.DataFrame(
        [{
            "Invoice #": "P1",
            "Date": "2021-04-01",
            "Period From": "2021-01-01",
            "Period To": "2021-03-31",
            "Amount (£)": 100.0,
            "Period Charge (£)": 50.0,
            "Entry Type": "Payment",
        }]
    )
    config: ConfigDict = {"analysis_min": 500.0}
    dfc = _prepare_analysis_frame(df_an, config)
    assert len(dfc) == 1


def test_credit_below_min_kept() -> None:
    df_an = pd.DataFrame(
        [{
            "Invoice #": "C1",
            "Date": "2021-04-01",
            "Period From": "2021-01-01",
            "Period To": "2021-03-31",
            "Amount (£)": 100.0,
            "Period Charge (£)": 50.0,
            "Entry Type": "Credit",
        }]
    )
    config: ConfigDict = {"analysis_min": 500.0}
    dfc = _prepare_analysis_frame(df_an, config)
    assert len(dfc) == 1


def test_payment_legal_candidate_kept() -> None:
    df_an = pd.DataFrame(
        [{
            "Invoice #": "P2",
            "Date": "2022-01-02",
            "Period From": "2020-01-01",
            "Period To": "2020-01-02",
            "Amount (£)": 100.0,
            "Period Charge (£)": 50.0,
            "Entry Type": "Payment",
        }]
    )
    config: ConfigDict = {"analysis_min": 500.0}
    dfc = _prepare_analysis_frame(df_an, config)
    assert len(dfc) == 1


def test_bill_above_min_kept() -> None:
    df_an = pd.DataFrame(
        [{
            "Invoice #": "T-ABOVE",
            "Date": "2021-04-01",
            "Period From": "2021-01-01",
            "Period To": "2021-03-31",
            "Amount (£)": 600.0,
            "Period Charge (£)": 600.0,
            "Entry Type": "New Bill",
        }]
    )
    config: ConfigDict = {"analysis_min": 500.0}
    dfc = _prepare_analysis_frame(df_an, config)
    assert len(dfc) == 1
