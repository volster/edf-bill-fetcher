import pandas as pd

from edf_bill_fetcher.models.report_models import compute_payment_analysis

RECORD_KEYS = ["Date", "Entry Type", "Period Charge (£)", "Amount (£)", "Details"]


def test_empty_frame_yields_zeroed_analysis() -> None:
    analysis = compute_payment_analysis(pd.DataFrame(columns=RECORD_KEYS))

    assert analysis.count == 0
    assert analysis.total_paid == 0.0
    assert analysis.avg_payment == 0.0
    assert analysis.median_payment == 0.0
    assert analysis.largest_payment == 0.0
    assert analysis.smallest_payment == 0.0
    assert analysis.avg_interval_days is None
    assert analysis.median_interval_days is None
    assert analysis.last_payment_date is None
    assert analysis.last_payment_amount is None
    assert analysis.chronology.empty


def test_single_payment_uses_period_charge_and_no_intervals() -> None:
    df = pd.DataFrame(
        [
            {
                "Date": "01/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": 900,
                "Amount (£)": 500,
                "Details": "customer payment",
            }
        ]
    )

    analysis = compute_payment_analysis(df)

    assert analysis.count == 1
    assert analysis.total_paid == 900.0
    assert analysis.largest_payment == 900.0
    assert analysis.smallest_payment == 900.0
    assert analysis.chronology["_amount"].iloc[0] == 900.0
    assert analysis.avg_interval_days is None
    assert analysis.median_interval_days is None


def test_two_payments_thirty_days_apart() -> None:
    df = pd.DataFrame(
        [
            {
                "Date": "31/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": 100,
                "Amount (£)": 100,
                "Details": "",
            },
            {
                "Date": "01/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": 100,
                "Amount (£)": 100,
                "Details": "",
            },
        ]
    )

    analysis = compute_payment_analysis(df)

    assert analysis.avg_interval_days == 30.0
    assert analysis.median_interval_days == 30.0


def test_amount_fallback_when_period_charge_is_na() -> None:
    df = pd.DataFrame(
        [
            {
                "Date": "01/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": "N/A",
                "Amount (£)": 100,
                "Details": "",
            }
        ]
    )

    analysis = compute_payment_analysis(df)

    assert analysis.chronology["_amount"].iloc[0] == 100.0
    assert analysis.total_paid == 100.0


def test_credit_included_and_new_bill_excluded() -> None:
    df = pd.DataFrame(
        [
            {
                "Date": "01/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": 200,
                "Amount (£)": 200,
                "Details": "",
            },
            {
                "Date": "02/01/2023",
                "Entry Type": "Credit",
                "Period Charge (£)": 50,
                "Amount (£)": 50,
                "Details": "",
            },
            {
                "Date": "03/01/2023",
                "Entry Type": "New Bill",
                "Period Charge (£)": 999,
                "Amount (£)": 999,
                "Details": "",
            },
        ]
    )

    analysis = compute_payment_analysis(df)

    assert analysis.count == 2
    assert analysis.total_paid == 250.0


def test_negative_amount_is_absoled_at_stat_level() -> None:
    df = pd.DataFrame(
        [
            {
                "Date": "01/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": -500,
                "Amount (£)": -500,
                "Details": "",
            }
        ]
    )

    analysis = compute_payment_analysis(df)

    assert analysis.total_paid == 500.0
    assert analysis.chronology["_amount"].iloc[0] == -500.0


def test_last_payment_is_chronologically_last_row() -> None:
    df = pd.DataFrame(
        [
            {
                "Date": "05/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": 100,
                "Amount (£)": 100,
                "Details": "",
            },
            {
                "Date": "20/01/2023",
                "Entry Type": "Payment",
                "Period Charge (£)": 250,
                "Amount (£)": 250,
                "Details": "",
            },
        ]
    )

    analysis = compute_payment_analysis(df)

    assert analysis.last_payment_date == "20/01/2023"
    assert analysis.last_payment_amount == 250.0


# --- DataQualityReport (Arch #3 Task 1) ----------------------------------------


def _dq_records() -> pd.DataFrame:
    """A small frame exercising each data-quality dimension.

    3 records: one fully populated, one with a parseable date but missing
    amount, one duplicate of the first (same Date + Amount) to exercise the
    duplicate_rate path. Reading is classified on two rows.
    """
    return pd.DataFrame(
        [
            {
                "Date": "01/01/2023",
                "Source": "PDF",
                "Entry Type": "New Bill",
                "Amount (£)": 120.0,
                "Period From": "01/01/2023",
                "Period To": "01/02/2023",
                "Reading": "Actual",
                "Unit Rate (p/kWh)": 30.0,
            },
            {
                "Date": "01/02/2023",
                "Source": "HTM",
                "Entry Type": "Payment",
                "Amount (£)": None,
                "Period From": "N/A",
                "Period To": "N/A",
                "Reading": "N/A",
                "Unit Rate (p/kWh)": "N/A",
            },
            {
                "Date": "01/01/2023",
                "Source": "PDF",
                "Entry Type": "New Bill",
                "Amount (£)": 120.0,
                "Period From": "01/01/2023",
                "Period To": "01/02/2023",
                "Reading": "Actual",
                "Unit Rate (p/kWh)": 30.0,
            },
        ]
    )


def test_data_quality_report_empty_frame() -> None:
    from edf_bill_fetcher.models.report_models import (
        DataQualityReport,
        compute_data_quality_report,
    )

    report = compute_data_quality_report(pd.DataFrame())

    assert isinstance(report, DataQualityReport)
    assert report.total_records == 0
    assert report.date_parsed == 0
    assert report.date_failed == 0
    assert report.date_parse_rate == 0.0
    assert report.amt_complete == 0
    assert report.amt_missing == 0
    assert report.period_complete == 0
    assert report.period_completeness_rate == 0.0
    assert report.reading_classified == 0
    assert report.reading_classify_rate == 0.0
    assert report.ur_computable == 0
    assert report.ur_computable_rate == 0.0
    assert report.duplicate_count == 0
    assert report.duplicate_rate == 0.0
    assert report.source_distribution == {}
    assert report.entry_type_distribution == {}


def test_data_quality_report_counts_and_rates() -> None:
    from edf_bill_fetcher.models.report_models import compute_data_quality_report

    report = compute_data_quality_report(_dq_records())

    assert report.total_records == 3
    assert report.date_parsed == 3
    assert report.date_failed == 0
    assert report.date_parse_rate == 1.0
    assert report.amt_complete == 2
    assert report.amt_missing == 1
    assert report.period_complete == 2
    assert report.period_completeness_rate == 2 / 3
    assert report.reading_classified == 2
    assert report.reading_classify_rate == 2 / 3
    assert report.ur_computable == 2
    assert report.ur_computable_rate == 2 / 3
    assert report.duplicate_count == 1
    assert report.duplicate_rate == 1 / 3
    assert report.source_distribution == {"PDF": 2, "HTM": 1}
    assert report.entry_type_distribution == {"New Bill": 2, "Payment": 1}


def test_data_quality_report_does_not_mutate_input() -> None:
    from edf_bill_fetcher.models.report_models import compute_data_quality_report

    df = _dq_records()
    columns_before = set(df.columns)

    compute_data_quality_report(df)

    assert set(df.columns) == columns_before  # no _dt_parsed leakage


# --- StatisticalAnalysis (Arch #3 Task 2) --------------------------------------


def _amount_frame(*amounts: float) -> pd.DataFrame:
    return pd.DataFrame([{"Date": "01/01/2023", "Amount (£)": a} for a in amounts])


def test_statistical_insufficient_data() -> None:
    from edf_bill_fetcher.models.report_models import (
        StatisticalAnalysis,
        compute_statistical_analysis,
    )

    report = compute_statistical_analysis(_amount_frame(100.0, 200.0))

    assert isinstance(report, StatisticalAnalysis)
    assert report.count == 2
    assert report.shapiro_stat is None
    assert report.shapiro_p is None


def test_statistical_descriptive_stats() -> None:
    from edf_bill_fetcher.models.report_models import compute_statistical_analysis

    report = compute_statistical_analysis(_amount_frame(100.0, 200.0, 300.0, 400.0))

    assert report.count == 4
    assert report.mean == 250.0
    assert report.median == 250.0
    assert report.minimum == 100.0
    assert report.maximum == 400.0
    assert report.range == 300.0
    assert report.std is not None


def test_statistical_rolling_matches_pandas() -> None:
    import pandas as pd

    from edf_bill_fetcher.models.report_models import compute_statistical_analysis

    amounts = [100.0, 200.0, 300.0, 400.0, 500.0, 600.0, 700.0]
    report = compute_statistical_analysis(_amount_frame(*amounts))
    series = pd.Series(amounts)

    assert report.count == 7
    assert report.rolling["mean"] == series.rolling(6, min_periods=1).mean().iloc[-1]


# --- ForecastResult (Arch #3 Task 3) -------------------------------------------


def test_forecast_insufficient_data() -> None:
    from edf_bill_fetcher.models.report_models import (
        ForecastResult,
        compute_forecast,
    )

    result = compute_forecast(_amount_frame(100.0, 200.0))

    assert isinstance(result, ForecastResult)
    assert result.n == 2
    assert result.linear_forecast == []
    assert result.ema_forecast == []
    assert result.hw_forecast is None


def test_forecast_ema_alpha_03() -> None:
    from edf_bill_fetcher.models.report_models import compute_forecast

    amounts = [100.0, 200.0, 300.0, 400.0]
    result = compute_forecast(_amount_frame(*amounts))

    # EMA with alpha=0.3, folding left-to-right from the first value.
    ema = amounts[0]
    for val in amounts[1:]:
        ema = 0.3 * val + (1 - 0.3) * ema

    assert result.n == 4
    assert result.ema_forecast == [ema] * 6


def test_forecast_linear_shape() -> None:
    from edf_bill_fetcher.models.report_models import compute_forecast

    result = compute_forecast(_amount_frame(100.0, 200.0, 300.0, 400.0, 500.0, 600.0, 700.0))

    assert result.n == 7
    assert len(result.linear_forecast) == 6
    assert len(result.ema_forecast) == 6
    assert all(isinstance(v, float) for v in result.linear_forecast)
