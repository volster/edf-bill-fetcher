import pandas as pd
import pytest

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


def test_statistical_extended_fields() -> None:
    from edf_bill_fetcher.models.report_models import compute_statistical_analysis

    report = compute_statistical_analysis(_amount_frame(100.0, 200.0, 300.0, 400.0))

    assert isinstance(report.skewness, float)
    assert isinstance(report.kurtosis, float)
    assert isinstance(report.rolling_median, float)
    assert isinstance(report.ema, float)
    assert report.momentum == 300.0  # 400 - 100 (3-period diff)
    assert isinstance(report.volatility, float)
    assert report.z_count == 0
    assert report.iqr_count == 0


def test_statistical_cv_is_percentage() -> None:
    from edf_bill_fetcher.models.report_models import compute_statistical_analysis

    # mean=250, std~129 -> cv as a percentage is ~51.6 (not the 0.52 ratio).
    report = compute_statistical_analysis(_amount_frame(100.0, 200.0, 300.0, 400.0))
    assert report.cv is not None
    assert report.cv > 1.0


def test_statistical_volatility_finite_on_zero_amount() -> None:
    import math

    from edf_bill_fetcher.models.report_models import compute_statistical_analysis

    report = compute_statistical_analysis(_amount_frame(100.0, 0.0, 150.0, 120.0, 200.0, 180.0))
    assert math.isfinite(report.volatility)


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


def test_forecast_ema_uses_span6_ewm() -> None:
    import pandas as pd

    from edf_bill_fetcher.models.report_models import compute_forecast

    amounts = [100.0, 200.0, 300.0, 400.0]
    result = compute_forecast(_amount_frame(*amounts))

    series = pd.Series(amounts)
    ema_last = float(series.ewm(span=6, adjust=False).mean().iloc[-1])

    assert result.n == 4
    assert result.ema_forecast == [ema_last] * 6
    assert len(result.ema_series) == 4


def test_forecast_linear_fitted_present() -> None:
    from edf_bill_fetcher.models.report_models import compute_forecast

    result = compute_forecast(_amount_frame(100.0, 200.0, 300.0, 400.0, 500.0, 600.0, 700.0))

    assert result.n == 7
    assert result.linear_fitted is not None
    assert len(result.linear_fitted) == 7
    assert len(result.linear_forecast) == 6


def test_forecast_accuracy_metrics() -> None:
    from edf_bill_fetcher.models.report_models import compute_forecast

    amounts = [100.0 * i for i in range(1, 11)]
    result = compute_forecast(_amount_frame(*amounts))

    assert result.n == 10
    assert result.mae is not None
    assert result.rmse is not None
    assert result.mape is not None


def test_forecast_linear_shape() -> None:
    from edf_bill_fetcher.models.report_models import compute_forecast

    result = compute_forecast(_amount_frame(100.0, 200.0, 300.0, 400.0, 500.0, 600.0, 700.0))

    assert result.n == 7
    assert len(result.linear_forecast) == 6
    assert len(result.ema_forecast) == 6
    assert all(isinstance(v, float) for v in result.linear_forecast)


# --- OfgemComparison (Arch #3 Task 4) ---------------------------------------------


def _ofgem_frame(*rows: tuple[str, float, float]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"Date": date, "Period Charge (£)": charge, "Units (kWh)": units}
            for date, charge, units in rows
        ]
    )


def _patch_caps(monkeypatch: pytest.MonkeyPatch, caps: dict, latest: dict | None) -> None:
    import edf_bill_fetcher.models.report_models as rm

    monkeypatch.setattr(rm, "load_ofgem_caps", lambda auto_carry=True: (caps, latest))


def test_ofgem_in_table_and_unavailable_rows(monkeypatch: pytest.MonkeyPatch) -> None:
    from edf_bill_fetcher.models.report_models import compute_ofgem_comparison

    _patch_caps(
        monkeypatch,
        {
            "2024-Q1": {"unit_rate": 28.62, "standing_charge": 53.35},
            "2024-Q2": {"unit_rate": 24.50, "standing_charge": 60.10},
        },
        None,
    )
    result = compute_ofgem_comparison(
        _ofgem_frame(
            ("15/02/2024", 200.0, 500.0),  # 40.0 p/kWh > 28.62 → EXCEEDS CAP
            ("15/05/2024", 300.0, 600.0),  # 50.0 p/kWh > 24.50 → EXCEEDS
            ("15/08/2024", 400.0, 700.0),  # 2024-Q3 absent → CAP DATA UNAVAILABLE
        )
    )

    assert [r.quarter for r in result.rows] == ["2024-Q1", "2024-Q2", "2024-Q3"]
    q1 = result.rows[0]
    assert q1.bill_rate == pytest.approx(40.0)
    assert q1.cap_rate == 28.62
    assert q1.status == "EXCEEDS CAP"
    q3 = result.rows[2]
    assert q3.cap_rate is None
    assert q3.diff is None
    assert q3.status == "CAP DATA UNAVAILABLE"
    assert result.exceed_count == 2
    assert result.unavailable_count == 1


def test_ofgem_carried_forward(monkeypatch: pytest.MonkeyPatch) -> None:
    from edf_bill_fetcher.models.report_models import compute_ofgem_comparison

    _patch_caps(
        monkeypatch,
        {"2026-Q3": {"unit_rate": 25.0, "standing_charge": 60.0}},
        {"unit_rate": 25.0, "standing_charge": 60.0},
    )
    result = compute_ofgem_comparison(_ofgem_frame(("15/10/2026", 200.0, 800.0)))  # 25.0 p/kWh

    assert result.rows[0].quarter == "2026-Q4"
    assert result.rows[0].status == "AT CAP (CAP CARRIED FORWARD)"
    assert result.carried_count == 1
    assert result.overall_verdict == "COMPLIANT (CARRIED)"


def test_ofgem_nan_rate_quarter_dropped(monkeypatch: pytest.MonkeyPatch) -> None:
    from edf_bill_fetcher.models.report_models import compute_ofgem_comparison

    _patch_caps(
        monkeypatch,
        {"2026-Q3": {"unit_rate": 25.0, "standing_charge": 60.0}},
        {"unit_rate": 25.0, "standing_charge": 60.0},
    )
    result = compute_ofgem_comparison(_ofgem_frame(("15/10/2026", 0.0, 0.0)))  # 0/0 → NaN

    assert result.rows == []
    assert result.overall_verdict == "COMPLIANT"


def test_ofgem_verdict_precedence() -> None:
    from edf_bill_fetcher.models.report_models import OfgemComparison

    assert OfgemComparison([], 1, 1, 1, None, None).overall_verdict == "REVIEW REQUIRED"
    assert OfgemComparison([], 0, 1, 1, None, None).overall_verdict == "INCOMPLETE"
    assert OfgemComparison([], 0, 0, 1, None, None).overall_verdict == "COMPLIANT (CARRIED)"
    assert OfgemComparison([], 0, 0, 0, None, None).overall_verdict == "COMPLIANT"


# --- TariffAnalysis (Arch #3 Task 5) ----------------------------------------------


def _tariff_frame(tariffs: list[str]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Date": f"01/{idx:02d}/2024",
                "Tariff": t,
                "Unit Rate (p/kWh)": 20.0 + idx,
                "Period Charge (£)": 100.0 + idx,
            }
            for idx, t in enumerate(tariffs, 1)
        ]
    )


def test_tariff_empty_frame() -> None:
    from edf_bill_fetcher.models.report_models import TariffAnalysis, compute_tariff_analysis

    result = compute_tariff_analysis(pd.DataFrame(columns=["Date", "Tariff", "Unit Rate (p/kWh)"]))
    assert isinstance(result, TariffAnalysis)
    assert result.empty


def test_tariff_stats_and_changes() -> None:
    from edf_bill_fetcher.models.report_models import compute_tariff_analysis

    frame = _tariff_frame(["Standard", "Standard", "Standard", "Fixed", "Fixed"])
    result = compute_tariff_analysis(frame)

    assert not result.empty
    assert result.num_tariffs == 2
    assert result.tariff_changes == 2  # two tariff segments (Standard run, Fixed run)

    stats = result.stats.set_index("Tariff")
    std = stats.loc["Standard"]
    assert int(std["count"]) == 3
    assert std["avg_unit_rate"] == pytest.approx((21.0 + 22.0 + 23.0) / 3)
    assert std["min_unit_rate"] == 21.0
    assert std["max_unit_rate"] == 23.0
    assert std["avg_charge"] == pytest.approx((101.0 + 102.0 + 103.0) / 3)


# --- DisputeAnalysis (Arch #3 Task 6) ---------------------------------------------


def test_dispute_analysis_wraps_compute_dispute_flags() -> None:
    from edf_bill_fetcher.models.report_models import DisputeAnalysis, compute_dispute_analysis
    from edf_bill_fetcher.processors.analysis import compute_dispute_flags

    df = pd.DataFrame(
        [
            {"Date": "01/01/2024", "Amount (£)": 100.0, "_dt": pd.Timestamp("2024-01-01")},
            {"Date": "15/02/2024", "Amount (£)": 500.0, "_dt": pd.Timestamp("2024-02-15")},
        ]
    )

    result = compute_dispute_analysis(df)
    flags, counts = compute_dispute_flags(df)

    assert isinstance(result, DisputeAnalysis)
    assert result.flags == flags
    assert result.counts == counts
    assert any(f[0] == "LARGE JUMP" for f in result.flags)
