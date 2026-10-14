"""Typed computed-report results shared by the Excel writer and the PDF/DOCX reporters."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime

import numpy as np
import pandas as pd

from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.helpers.ofgem_caps import load_ofgem_caps
from edf_bill_fetcher.helpers.payment_figures import payment_amounts
from edf_bill_fetcher.processors.preparation import (
    balance_frame,
    monthly_balance_frame,
    normalize_frame,
    number,
    scope,
)


@dataclass
class PaymentAnalysis:  # noqa: D101
    count: int
    total_paid: float
    avg_payment: float
    median_payment: float
    largest_payment: float
    smallest_payment: float
    avg_interval_days: float | None
    median_interval_days: float | None
    last_payment_date: str | None
    last_payment_amount: float | None
    chronology: pd.DataFrame = field(default_factory=lambda: pd.DataFrame())


def compute_payment_analysis(df: pd.DataFrame) -> PaymentAnalysis:
    """Compute payment/credit statistics for the analysis surfaces."""
    df = normalize_frame(df)
    payments = df[
        df["Entry Type"].isin(["Payment", "Credit"]) & df["Transaction Amount (£)"].notna()
    ].copy()
    if payments.empty:
        return PaymentAnalysis(
            count=0,
            total_paid=0.0,
            avg_payment=0.0,
            median_payment=0.0,
            largest_payment=0.0,
            smallest_payment=0.0,
            avg_interval_days=None,
            median_interval_days=None,
            last_payment_date=None,
            last_payment_amount=None,
            chronology=payments,
        )

    payments["_dt"] = payments["Date"].apply(parse_to_sort_date)
    payments = payments.sort_values("_dt", kind="stable").reset_index(drop=True)
    # The shared helper provides payment/credit magnitudes; account balance
    # observations are never substituted for a missing transaction.
    payments["_amount"] = payment_amounts(payments)

    dated = payments[payments["_dt"].notna()]
    ledgers = pd.Series([scope(row) for row in dated.to_dict("records")], index=dated.index)
    intervals = dated.groupby(ledgers, sort=False)["_dt"].diff().dt.days.dropna()
    latest = dated.iloc[-1] if not dated.empty else None

    return PaymentAnalysis(
        count=len(payments),
        total_paid=float(abs(payments["_amount"].sum())),
        avg_payment=float(abs(payments["_amount"].mean())),
        median_payment=float(abs(payments["_amount"].median())),
        largest_payment=float(abs(payments["_amount"].max())),
        smallest_payment=float(abs(payments["_amount"].min())),
        avg_interval_days=float(intervals.mean()) if len(intervals) > 0 else None,
        median_interval_days=float(intervals.median()) if len(intervals) > 0 else None,
        last_payment_date=latest["Date"] if latest is not None else None,
        last_payment_amount=float(latest["_amount"]) if latest is not None else None,
        chronology=payments,
    )


@dataclass
class DataQualityReport:
    """Typed computed data-quality metrics shared by Excel and the reporters."""

    total_records: int
    date_parsed: int
    date_failed: int
    date_parse_rate: float
    amt_complete: int
    amt_missing: int
    period_complete: int
    period_completeness_rate: float
    reading_classified: int
    reading_classify_rate: float
    ur_computable: int
    ur_computable_rate: float
    duplicate_count: int
    duplicate_rate: float
    source_distribution: dict[str, int]
    entry_type_distribution: dict[str, int]

    def to_dict(self) -> dict:
        """Map back to the legacy flat dict consumed by the Excel writer."""
        if self.total_records == 0:
            return {}
        return {
            "total_records": self.total_records,
            "date_parsed": self.date_parsed,
            "date_failed": self.date_failed,
            "date_parse_rate": self.date_parse_rate,
            "amt_complete": self.amt_complete,
            "amt_missing": self.amt_missing,
            "period_complete": self.period_complete,
            "period_completeness_rate": self.period_completeness_rate,
            "reading_classified": self.reading_classified,
            "reading_classify_rate": self.reading_classify_rate,
            "ur_computable": self.ur_computable,
            "ur_computable_rate": self.ur_computable_rate,
            "duplicate_count": self.duplicate_count,
            "duplicate_rate": self.duplicate_rate,
            "source_distribution": self.source_distribution,
            "entry_type_distribution": self.entry_type_distribution,
        }


def compute_data_quality_report(
    df: pd.DataFrame, *, duplicate_count: int | None = None
) -> DataQualityReport:
    """Compute the canonical data-quality report on a copy of ``df``."""
    df = normalize_frame(df)
    total_records = len(df)
    if total_records == 0:
        return DataQualityReport(
            total_records=0,
            date_parsed=0,
            date_failed=0,
            date_parse_rate=0.0,
            amt_complete=0,
            amt_missing=0,
            period_complete=0,
            period_completeness_rate=0.0,
            reading_classified=0,
            reading_classify_rate=0.0,
            ur_computable=0,
            ur_computable_rate=0.0,
            duplicate_count=0,
            duplicate_rate=0.0,
            source_distribution={},
            entry_type_distribution={},
        )

    df["_dt_parsed"] = df["Date"].apply(parse_to_sort_date)
    date_parsed = int(df["_dt_parsed"].notna().sum())
    date_failed = total_records - date_parsed

    amt_complete = int(df["Amount (£)"].map(number).notna().sum())
    amt_missing = total_records - amt_complete

    period_complete = (
        int(
            (
                (df["Period From"].apply(parse_to_sort_date).notna())
                & (df["Period To"].apply(parse_to_sort_date).notna())
                & (
                    df["Period To"].apply(parse_to_sort_date)
                    >= df["Period From"].apply(parse_to_sort_date)
                )
            ).sum()
        )
        if "Period From" in df.columns and "Period To" in df.columns
        else 0
    )

    reading_classified = (
        int(df["Reading"].isin(["Actual", "Smart", "Estimated"]).sum())
        if "Reading" in df.columns
        else 0
    )

    ur_computable = (
        int(df["Unit Rate (p/kWh)"].map(number).notna().sum())
        if "Unit Rate (p/kWh)" in df.columns
        else 0
    )

    from edf_bill_fetcher.processors.preparation import prepare_evidence

    if duplicate_count is None:
        duplicate_count = len(
            prepare_evidence(df.to_dict("records"), {"use_dedup": True}).duplicates
        )

    source_distribution = df["Source"].value_counts().to_dict()
    entry_type_distribution = (
        df["Entry Type"].value_counts().to_dict() if "Entry Type" in df.columns else {}
    )

    return DataQualityReport(
        total_records=total_records,
        date_parsed=date_parsed,
        date_failed=date_failed,
        date_parse_rate=date_parsed / total_records,
        amt_complete=amt_complete,
        amt_missing=amt_missing,
        period_complete=period_complete,
        period_completeness_rate=period_complete / total_records,
        reading_classified=reading_classified,
        reading_classify_rate=reading_classified / total_records,
        ur_computable=ur_computable,
        ur_computable_rate=ur_computable / total_records,
        duplicate_count=duplicate_count,
        duplicate_rate=duplicate_count / total_records,
        source_distribution=source_distribution,
        entry_type_distribution=entry_type_distribution,
    )


@dataclass
class StatisticalAnalysis:
    """Typed descriptive + rolling + anomaly statistics shared by all surfaces."""

    count: int
    mean: float
    median: float
    std: float
    minimum: float
    maximum: float
    range: float
    cv: float | None
    rolling: dict[str, float]
    shapiro_stat: float | None
    shapiro_p: float | None
    skewness: float
    kurtosis: float
    rolling_median: float
    ema: float
    momentum: float
    volatility: float
    jb_stat: float | None
    jb_p: float | None
    z_count: int
    iqr_count: int
    z_dates: list[str]
    iqr_dates: list[str]


def compute_statistical_analysis(df: pd.DataFrame) -> StatisticalAnalysis:
    """Compute descriptive + rolling + anomaly statistics from the amount column."""
    from edf_bill_fetcher.helpers.date_utils import compute_ema, compute_momentum
    from edf_bill_fetcher.processors.forecasting import (
        _iqr_anomalies,
        _sanitised_returns,
        _zscore_anomalies,
    )

    work = balance_frame(df)
    if len({scope(r) for r in work.to_dict("records")}) > 1:
        work = work.iloc[:0]
    if "Date" in work.columns:
        work["_dt"] = work["Date"].apply(parse_to_sort_date)
        work = work.sort_values("_dt").reset_index(drop=True)
    raw = pd.to_numeric(work["Amount (£)"], errors="coerce").astype(float)
    series = pd.Series(raw.values, index=pd.to_datetime(work["_dt"]))
    n = int(raw.notna().sum())

    empty = StatisticalAnalysis(
        count=0,
        mean=0.0,
        median=0.0,
        std=0.0,
        minimum=0.0,
        maximum=0.0,
        range=0.0,
        cv=None,
        rolling={"mean": 0.0, "std": 0.0, "min": 0.0, "max": 0.0},
        shapiro_stat=None,
        shapiro_p=None,
        skewness=0.0,
        kurtosis=0.0,
        rolling_median=0.0,
        ema=0.0,
        momentum=0.0,
        volatility=0.0,
        jb_stat=None,
        jb_p=None,
        z_count=0,
        iqr_count=0,
        z_dates=[],
        iqr_dates=[],
    )
    if n == 0:
        return empty

    amounts_series = raw.dropna().astype(float)
    mean = float(amounts_series.mean())
    median = float(amounts_series.median())
    std = float(amounts_series.std())
    minimum = float(amounts_series.min())
    maximum = float(amounts_series.max())
    skewness = float(amounts_series.skew())
    kurtosis = float(amounts_series.kurtosis())
    cv = std / mean * 100 if mean > 0 else 0.0

    roll = raw.rolling(6, min_periods=1)
    roll_std = float(roll.std().iloc[-1])
    if roll_std != roll_std:  # NaN guard, mirroring the PDF renderer
        roll_std = 0.0
    rolling = {
        "mean": float(roll.mean().iloc[-1]),
        "std": roll_std,
        "min": float(roll.min().iloc[-1]),
        "max": float(roll.max().iloc[-1]),
    }
    rolling_median = float(roll.median().iloc[-1])

    ema = float(compute_ema(amounts_series, span=6).iloc[-1])
    momentum = float(compute_momentum(amounts_series, period=3).iloc[-1]) if n >= 4 else 0.0

    returns = _sanitised_returns(series)
    volatility = (
        float(returns.rolling(6, min_periods=1).std().iloc[-1]) if len(returns) >= 2 else 0.0
    )

    shapiro_stat: float | None = None
    shapiro_p: float | None = None
    jb_stat: float | None = None
    jb_p: float | None = None
    if n >= 3:
        try:
            from scipy import stats as sp_stats

            s_stat, s_p = sp_stats.shapiro(amounts_series)
            shapiro_stat = float(s_stat)
            shapiro_p = float(s_p)
            jb_s, jb_p_ = sp_stats.jarque_bera(amounts_series)
            jb_stat = float(jb_s)
            jb_p = float(jb_p_)
        except ImportError:
            pass

    z_mask = _zscore_anomalies(series, threshold=2.5)
    iqr_mask = _iqr_anomalies(series, multiplier=1.5)

    return StatisticalAnalysis(
        count=n,
        mean=mean,
        median=median,
        std=std,
        minimum=minimum,
        maximum=maximum,
        range=maximum - minimum,
        cv=cv,
        rolling=rolling,
        shapiro_stat=shapiro_stat,
        shapiro_p=shapiro_p,
        skewness=skewness,
        kurtosis=kurtosis,
        rolling_median=rolling_median,
        ema=ema,
        momentum=momentum,
        volatility=volatility,
        jb_stat=jb_stat,
        jb_p=jb_p,
        z_count=int(z_mask.sum()),
        iqr_count=int(iqr_mask.sum()),
        z_dates=[
            f"{d.strftime('%d/%m/%Y') if hasattr(d, 'strftime') else d} ({v:,.2f})"
            for d, v in series[z_mask].items()
        ],
        iqr_dates=[
            f"{d.strftime('%d/%m/%Y') if hasattr(d, 'strftime') else d} ({v:,.2f})"
            for d, v in series[iqr_mask].items()
        ],
    )


@dataclass
class ForecastResult:
    """Typed multi-method forecast shared by the Excel writer and the reports."""

    n: int
    linear_forecast: list[float]
    ema_forecast: list[float]
    hw_forecast: list[float] | None
    model_info: list[str]
    linear_fitted: list[float] | None
    hw_fitted: list[float] | None
    ema_series: list[float]
    hist_vol: float
    mae: float | None
    rmse: float | None
    mape: float | None


def compute_forecast(df: pd.DataFrame) -> ForecastResult:
    """Compute linear / EMA / Holt-Winters projections using the canonical helpers.

    This is the same math the ``Forecast & Projection`` sheet runs, so the
    report and the workbook cannot drift apart.
    """
    from edf_bill_fetcher.helpers.date_utils import compute_ema
    from edf_bill_fetcher.processors.forecasting import (
        _holt_winters_forecast_pair,
        _linear_forecast_pair,
        _sanitised_returns,
    )

    work = monthly_balance_frame(df)
    raw = pd.to_numeric(work["Amount (£)"], errors="coerce").astype(float)
    series = pd.Series(raw.values, index=pd.to_datetime(work["_dt"]))
    n = int(raw.notna().sum()) if not series.isna().any() else 0

    def _empty() -> ForecastResult:
        return ForecastResult(
            n=n,
            linear_forecast=[],
            ema_forecast=[],
            hw_forecast=None,
            model_info=[],
            linear_fitted=None,
            hw_fitted=None,
            ema_series=[],
            hist_vol=0.05,
            mae=None,
            rmse=None,
            mape=None,
        )

    if n < 3 or series.isna().any():
        return _empty()

    forecast_steps = 6

    linear_fitted_arr, linear_fc = _linear_forecast_pair(series, forecast_steps)
    hw_fitted_arr, hw_fc = _holt_winters_forecast_pair(series, forecast_steps)

    ema_series_pd = compute_ema(series, span=6)
    ema_last = float(ema_series_pd.iloc[-1])
    ema_forecast = [ema_last] * forecast_steps

    returns = _sanitised_returns(series)
    hist_vol = float(returns.std()) if len(returns) >= 2 else 0.05

    model_info = ["Linear Trend — simple linear regression on time index"]
    if hw_fc is not None:
        model_info.append(
            "Holt-Winters — exponential smoothing with trend + seasonality (statsmodels)"
        )
    else:
        model_info.append("Holt-Winters — not available (need ≥4 usable points + statsmodels)")
    model_info.append("EMA — extends last Exponential Moving Average (span=6)")

    mae = rmse = mape = None
    if n >= 7:
        train = series.iloc[:-6]
        actual = series.values[-6:]
        lin_hist = _linear_forecast_pair(train, 6)[1]
        if lin_hist is not None:
            nonzero = actual != 0
            denom = actual[nonzero]
            mae = float(np.mean(np.abs(lin_hist - actual)))
            rmse = float(np.sqrt(np.mean((lin_hist - actual) ** 2)))
            if len(denom) > 0:
                mape = float(np.mean(np.abs((lin_hist[nonzero] - denom) / denom)))

    return ForecastResult(
        n=n,
        linear_forecast=[float(v) for v in linear_fc] if linear_fc is not None else [],
        ema_forecast=ema_forecast,
        hw_forecast=[float(v) for v in hw_fc] if hw_fc is not None else None,
        model_info=model_info,
        linear_fitted=[float(v) for v in linear_fitted_arr]
        if linear_fitted_arr is not None
        else None,
        hw_fitted=[float(v) for v in hw_fitted_arr] if hw_fitted_arr is not None else None,
        ema_series=[float(v) for v in ema_series_pd],
        hist_vol=hist_vol,
        mae=mae,
        rmse=rmse,
        mape=mape,
    )


def _period_to_ofgem_quarter(dt: datetime | None) -> str | None:
    if dt is None or pd.isna(dt):
        return None
    try:
        quarter = (dt.month - 1) // 3 + 1
        return f"{dt.year}-Q{quarter}"
    except Exception:
        return None


@dataclass
class OfgemRow:
    """One quarterly row in the OFGEM-cap comparison table."""

    quarter: str
    bill_rate: float
    cap_rate: float | None
    diff: float | None
    status: str


@dataclass
class OfgemComparison:
    """Cap-comparison result shared by the PDF/DOCX/HTML reporters."""

    rows: list[OfgemRow]
    exceed_count: int
    unavailable_count: int
    carried_count: int
    overall_avg: float | None
    overall_median: float | None

    @property
    def overall_verdict(self) -> str:
        """Highest-severity verdict: exceed > unavailable > carried > compliant."""
        if self.exceed_count > 0:
            return "REVIEW REQUIRED"
        if self.unavailable_count > 0:
            return "NOT ASSESSED"
        if self.carried_count > 0:
            return "PROVISIONAL BENCHMARK (CARRIED)"
        return "NOT ASSESSED" if not self.rows else "BELOW BENCHMARK"

    @property
    def overall_diff(self) -> str:
        """Summary-diff string for the OVERALL row's Difference column."""
        if self.exceed_count > 0:
            return f"{self.exceed_count} periods above reference benchmark"
        if self.unavailable_count > 0:
            return f"{self.unavailable_count} period(s) not benchmarked"
        if self.carried_count > 0:
            return f"{self.carried_count} period(s) used carried-forward cap"
        return "No comparable evidence" if not self.rows else "No benchmark exceedances"


def compute_ofgem_comparison(df: pd.DataFrame, config: dict | None = None) -> OfgemComparison:
    """Compare extracted energy rates during consumption with a contextual benchmark.

    The bundled data is a GB-average direct-debit benchmark, not an account's
    regional legal cap. Missing applicability or cap data remains unassessed.
    """
    cfg = config or {}
    caps, _ = load_ofgem_caps(auto_carry=False)
    work = normalize_frame(df)
    observations: dict[str, list[tuple[float, float, bool]]] = {}
    unknown = 0
    for _, bill in work.iterrows():
        if bill["Entry Type"] not in {"New Bill", "Ongoing Balance"}:
            continue
        payment = (
            str(bill.get("Payment Method") or cfg.get("payment_method", ""))
            .lower()
            .replace(" ", "")
            .replace("-", "")
        )
        fuel = str(bill.get("Fuel") or cfg.get("fuel", "")).lower()
        tariff_type = str(bill.get("Tariff Type") or cfg.get("tariff_type", "")).lower()
        region = str(cfg.get("cap_region", "GB-average"))
        applicable = (
            payment == "directdebit"
            and fuel == "electricity"
            and tariff_type in {"variable", "default"}
            and region == "GB-average"
        )
        spans = []
        subperiods = bill.get("Sub Periods")
        if isinstance(subperiods, list):
            spans = [
                (p.get("period_from"), p.get("period_to"), p.get("rate_p"), p.get("units_kwh"))
                for p in subperiods
            ]
        elif isinstance(subperiods, str) and subperiods.strip():
            for part in subperiods.split(";"):
                fields = part.strip().split("|")
                if len(fields) == 5:
                    spans.append((fields[0], fields[1], fields[3], fields[2]))
        if not spans:
            spans = [
                (
                    bill["Period From"],
                    bill["Period To"],
                    bill["Energy Unit Rate (p/kWh)"],
                    bill["Units (kWh)"],
                )
            ]
        for start_raw, end_raw, rate_raw, units_raw in spans:
            start, end = parse_to_sort_date(start_raw), parse_to_sort_date(end_raw)
            rate, units = number(rate_raw), number(units_raw)
            if pd.isna(start) or pd.isna(end) or end < start or rate is None or rate <= 0:
                unknown += 1
                continue
            if end == start:
                end += pd.Timedelta(days=1)
            length = (end - start).days
            cursor = start
            while cursor < end:
                quarter = _period_to_ofgem_quarter(cursor)
                boundary = cursor.to_period("Q").end_time.normalize() + pd.Timedelta(days=1)
                stop = min(end, boundary)
                weight = (stop - cursor).days * (
                    units / length if units is not None and units > 0 else 1
                )
                observations.setdefault(str(quarter), []).append((rate, weight, applicable))
                cursor = stop
    rows = []
    exceeded = 0
    unavailable = unknown
    carried = 0
    all_rates = []
    for quarter, values in sorted(observations.items()):
        rate = sum(v * w for v, w, _ in values) / sum(w for _, w, _ in values)
        all_rates.append(rate)
        cap = caps.get(quarter)
        if cap is None:
            unavailable += 1
            rows.append(OfgemRow(quarter, rate, None, None, "BENCHMARK DATA UNAVAILABLE"))
            continue
        cap_rate = float(cap["unit_rate"])
        difference = rate - cap_rate
        is_carried = bool(cap.get("is_carry", False))
        carried += int(is_carried)
        if not all(a for _, _, a in values):
            unavailable += 1
            status = "APPLICABILITY UNASSESSED — GB-average reference only"
        elif is_carried:
            status = "PROVISIONAL BENCHMARK (CARRIED)"
        elif difference > 0.01:
            exceeded += 1
            status = "ABOVE GB-AVERAGE BENCHMARK — REVIEW"
        elif abs(difference) <= 0.01:
            status = "AT GB-AVERAGE BENCHMARK"
        else:
            status = "BELOW GB-AVERAGE BENCHMARK"
        rows.append(OfgemRow(quarter, rate, cap_rate, difference, status))
    return OfgemComparison(
        rows,
        exceeded,
        unavailable,
        carried,
        float(np.mean(all_rates)) if all_rates else None,
        float(np.median(all_rates)) if all_rates else None,
    )


@dataclass
class TariffAnalysis:
    """Unit-rate stats + tariff-change count shared by all four surfaces."""

    stats: pd.DataFrame
    num_tariffs: int
    tariff_changes: int

    @property
    def empty(self) -> bool:
        """True when there are no tariff records with a computable unit rate."""
        return bool(self.stats.empty)


def compute_tariff_analysis(df: pd.DataFrame) -> TariffAnalysis:
    """Compute per-tariff unit-rate statistics and the tariff-change count."""
    if "Tariff" not in df.columns or "Unit Rate (p/kWh)" not in df.columns:
        return TariffAnalysis(pd.DataFrame(), 0, 0)

    tariff_data = df[df["Tariff"].notna() & (df["Tariff"] != "N/A")].copy()
    if tariff_data.empty:
        return TariffAnalysis(pd.DataFrame(), 0, 0)

    tariff_data["unit_rate_num"] = pd.to_numeric(tariff_data["Unit Rate (p/kWh)"], errors="coerce")
    tariff_data = tariff_data.dropna(subset=["unit_rate_num"])
    if tariff_data.empty:
        return TariffAnalysis(pd.DataFrame(), 0, 0)

    stats = (
        tariff_data.groupby("Tariff")
        .agg(
            count=("unit_rate_num", "count"),
            avg_unit_rate=("unit_rate_num", "mean"),
            median_unit_rate=("unit_rate_num", "median"),
            min_unit_rate=("unit_rate_num", "min"),
            max_unit_rate=("unit_rate_num", "max"),
            avg_charge=("Period Charge (£)", lambda x: pd.to_numeric(x, errors="coerce").mean()),
        )
        .reset_index()
    )

    tariff_data = tariff_data.sort_values("_dt" if "_dt" in tariff_data.columns else "Date")
    changes = tariff_data["Tariff"].ne(tariff_data["Tariff"].shift()).cumsum()
    num_tariffs = int(tariff_data["Tariff"].nunique())
    tariff_changes = int(changes.max()) if not changes.empty else 0

    return TariffAnalysis(stats, num_tariffs, tariff_changes)


@dataclass
class DisputeAnalysis:
    """Dispute flags + severity counts shared by the PDF/DOCX/HTML reporters."""

    flags: list[tuple[str, str | float | None, float | None, str, str]]
    counts: dict[str, int]


def compute_dispute_analysis(dfc: pd.DataFrame, mean_daily: float = 0.0) -> DisputeAnalysis:
    """Typed wrapper over the canonical dispute-flag detector."""
    from edf_bill_fetcher.processors.analysis import compute_dispute_flags

    flags, counts = compute_dispute_flags(dfc, mean_daily)
    return DisputeAnalysis(flags=flags, counts=counts)
