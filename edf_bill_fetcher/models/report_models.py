"""Typed computed-report results shared by the Excel writer and the PDF/DOCX reporters."""

from __future__ import annotations

import importlib.util
from dataclasses import dataclass, field
from datetime import datetime

import numpy as np
import pandas as pd

from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.helpers.ofgem_caps import load_ofgem_caps
from edf_bill_fetcher.helpers.payment_figures import payment_amounts


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
    payments = df[df["Entry Type"].isin(["Payment", "Credit"])].copy()
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
    payments = payments.sort_values("_dt").reset_index(drop=True)
    # `_amount` holds the RAW (signed) per-row transaction figure so
    # chronology rendering matches the old `payment_amount(row)[0]`
    # behavior exactly; the abs() mirroring the old dict is applied at
    # the STAT level (abs(sum), abs(mean), ...), not element-wise.
    payments["_amount"] = payment_amounts(payments)

    pay_dates = payments["_dt"].dropna()
    intervals = pay_dates.diff().dt.days.dropna()

    return PaymentAnalysis(
        count=len(payments),
        total_paid=float(abs(payments["_amount"].sum())),
        avg_payment=float(abs(payments["_amount"].mean())),
        median_payment=float(abs(payments["_amount"].median())),
        largest_payment=float(abs(payments["_amount"].max())),
        smallest_payment=float(abs(payments["_amount"].min())),
        avg_interval_days=float(intervals.mean()) if len(intervals) > 0 else None,
        median_interval_days=float(intervals.median()) if len(intervals) > 0 else None,
        last_payment_date=payments.iloc[-1]["Date"] if len(payments) > 0 else None,
        last_payment_amount=float(payments["_amount"].iloc[-1]) if len(payments) > 0 else None,
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


def compute_data_quality_report(df: pd.DataFrame) -> DataQualityReport:
    """Compute the canonical data-quality report on a copy of ``df``."""
    df = df.copy()
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

    amt_complete = int(df["Amount (£)"].notna().sum())
    amt_missing = total_records - amt_complete

    period_complete = (
        int(((df["Period From"] != "N/A") & (df["Period To"] != "N/A")).sum())
        if "Period From" in df.columns and "Period To" in df.columns
        else 0
    )

    reading_classified = int((df["Reading"] != "N/A").sum()) if "Reading" in df.columns else 0

    ur_computable = (
        int(df["Unit Rate (p/kWh)"].apply(lambda x: isinstance(x, int | float)).sum())
        if "Unit Rate (p/kWh)" in df.columns
        else 0
    )

    duplicate_count = int(df.duplicated(subset=["Date", "Amount (£)"]).sum())

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
    """Typed descriptive + rolling statistics shared by the reporters."""

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


def compute_statistical_analysis(df: pd.DataFrame) -> StatisticalAnalysis:
    """Compute descriptive + rolling statistics from the amount column."""
    work = df.copy()
    if "Date" in work.columns:
        work["_dt"] = work["Date"].apply(parse_to_sort_date)
        work = work.sort_values("_dt")
    amounts = pd.to_numeric(work["Amount (£)"], errors="coerce").dropna()
    n = len(amounts)
    if n == 0:
        return StatisticalAnalysis(
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
        )

    amounts_series = amounts.astype(float)
    mean = float(amounts_series.mean())
    median = float(amounts_series.median())
    std = float(amounts_series.std())
    minimum = float(amounts_series.min())
    maximum = float(amounts_series.max())
    cv = std / mean if mean and mean > 0 else None

    rolling = amounts_series.rolling(6, min_periods=1)
    roll_mean = float(rolling.mean().iloc[-1])
    roll_std = float(rolling.std().iloc[-1])
    roll_min = float(rolling.min().iloc[-1])
    roll_max = float(rolling.max().iloc[-1])
    if roll_std != roll_std:  # NaN guard, mirroring the PDF renderer
        roll_std = 0.0

    shapiro_stat: float | None = None
    shapiro_p: float | None = None
    if n >= 3:
        try:
            from scipy import stats as sp_stats

            s_stat, s_p = sp_stats.shapiro(amounts_series)
            shapiro_stat = float(s_stat)
            shapiro_p = float(s_p)
        except ImportError:
            pass

    return StatisticalAnalysis(
        count=n,
        mean=mean,
        median=median,
        std=std,
        minimum=minimum,
        maximum=maximum,
        range=maximum - minimum,
        cv=cv,
        rolling={"mean": roll_mean, "std": roll_std, "min": roll_min, "max": roll_max},
        shapiro_stat=shapiro_stat,
        shapiro_p=shapiro_p,
    )


@dataclass
class ForecastResult:
    """Typed multi-method forecast shared by the PDF/DOCX reporters."""

    n: int
    linear_forecast: list[float]
    ema_forecast: list[float]
    hw_forecast: list[float] | None
    model_info: list[str]


def compute_forecast(df: pd.DataFrame) -> ForecastResult:
    """Compute the linear-regression, EMA, and Holt-Winters projections."""
    work = df.copy()
    if "Date" in work.columns:
        work["_dt"] = work["Date"].apply(parse_to_sort_date)
        work = work.sort_values("_dt").reset_index(drop=True)
    amounts = pd.to_numeric(work["Amount (£)"], errors="coerce").dropna().astype(float)
    n = len(amounts)
    if n < 3:
        return ForecastResult(
            n=n, linear_forecast=[], ema_forecast=[], hw_forecast=None, model_info=[]
        )

    has_scipy = importlib.util.find_spec("scipy") is not None
    model_info: list[str] = []
    if has_scipy:
        from scipy import stats as sp_stats

        x = np.arange(n)
        slope, intercept, r_value, p_value, _std_err = sp_stats.linregress(x, amounts)
        linear_forecast = [float(intercept + slope * (n + i)) for i in range(1, 7)]
        model_info.append(
            f"Linear Regression: slope={slope:.2f}, intercept={intercept:.2f}, "
            f"R²={r_value**2:.4f}, p={p_value:.4f}"
        )
    else:
        linear_forecast = [float(amounts.mean())] * 6
        model_info.append("Linear Regression: not available (install scipy) - using mean")

    alpha = 0.3
    ema = float(amounts.iloc[0])
    for val in amounts.iloc[1:]:
        ema = alpha * float(val) + (1 - alpha) * ema
    ema_forecast = [ema] * 6
    model_info.append(f"EMA (α={alpha}): current level={ema:.2f}")

    try:
        from statsmodels.tsa.holtwinters import ExponentialSmoothing

        has_statsmodels = True
    except ImportError:
        has_statsmodels = False

    hw_forecast: list[float] | None = None
    if has_statsmodels and n >= 6:
        try:
            model = ExponentialSmoothing(amounts, trend="add", seasonal=None)
            hw_fit = model.fit(smoothing_level=alpha, smoothing_trend=0.1, optimized=True)
            hw_forecast = [float(v) for v in hw_fit.forecast(6).tolist()]
        except Exception:
            hw_forecast = None

    if hw_forecast:
        model_info.append("Holt-Winters: additive trend, no seasonality (fitted via statsmodels)")
    else:
        model_info.append("Holt-Winters: not available (install statsmodels)")

    return ForecastResult(
        n=n,
        linear_forecast=linear_forecast,
        ema_forecast=ema_forecast,
        hw_forecast=hw_forecast,
        model_info=model_info,
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
            return "INCOMPLETE"
        if self.carried_count > 0:
            return "COMPLIANT (CARRIED)"
        return "COMPLIANT"

    @property
    def overall_diff(self) -> str:
        """Summary-diff string for the OVERALL row's Difference column."""
        if self.exceed_count > 0:
            return f"{self.exceed_count} periods exceed cap"
        if self.unavailable_count > 0:
            return f"{self.unavailable_count} period(s) not benchmarked"
        if self.carried_count > 0:
            return f"{self.carried_count} period(s) used carried-forward cap"
        return "No exceedances"


def compute_ofgem_comparison(df: pd.DataFrame, config: dict | None = None) -> OfgemComparison:
    """Compute per-quarter unit-rate comparison against the OFGEM price cap."""
    del config
    ofgem_caps, latest_known_cap = load_ofgem_caps(auto_carry=True)

    work = df.copy()
    if "_dt" not in work.columns:
        work["_dt"] = work["Date"].apply(parse_to_sort_date)
    work = work.sort_values("_dt").reset_index(drop=True)

    valid_pc = work["Period Charge (£)"].notna() & (work["Period Charge (£)"] != "N/A")
    valid_units = (
        work["Units (kWh)"].notna() & (work["Units (kWh)"] != "N/A") & (work["Units (kWh)"] != "")
    )
    bills = work[valid_pc & valid_units].copy()
    if bills.empty:
        return OfgemComparison([], 0, 0, 0, None, None)

    bills["_unit_rate"] = (
        bills["Period Charge (£)"].astype(float) / bills["Units (kWh)"].astype(float) * 100
    )
    bills["_quarter"] = bills["_dt"].apply(_period_to_ofgem_quarter)

    all_rates = bills["_unit_rate"].dropna()
    overall_avg = float(all_rates.mean()) if not all_rates.empty else None
    overall_median = float(all_rates.median()) if not all_rates.empty else None

    rows: list[OfgemRow] = []
    exceed_count = 0
    unavailable_count = 0
    carried_count = 0
    for quarter in sorted(bills["_quarter"].dropna().unique()):
        avg_rate = bills[bills["_quarter"] == quarter]["_unit_rate"].mean()
        if pd.isna(avg_rate):
            continue
        if quarter not in ofgem_caps:
            if latest_known_cap:
                carried_count += 1
                cap_rate = latest_known_cap["unit_rate"]
                diff = float(avg_rate) - cap_rate
                if diff > 0:
                    status = "EXCEEDS CAP (CAP CARRIED FORWARD)"
                    exceed_count += 1
                elif abs(diff) < 0.01:
                    status = "AT CAP (CAP CARRIED FORWARD)"
                else:
                    status = "BELOW CAP (CAP CARRIED FORWARD)"
                rows.append(OfgemRow(quarter, float(avg_rate), cap_rate, diff, status))
            else:
                unavailable_count += 1
                rows.append(OfgemRow(quarter, float(avg_rate), None, None, "CAP DATA UNAVAILABLE"))
            continue
        cap_rate = ofgem_caps[quarter]["unit_rate"]
        diff = float(avg_rate) - cap_rate
        if diff > 0:
            status = "EXCEEDS CAP"
            exceed_count += 1
        elif abs(diff) < 0.01:
            status = "AT CAP"
        else:
            status = "BELOW CAP"
        rows.append(OfgemRow(quarter, float(avg_rate), cap_rate, diff, status))

    return OfgemComparison(
        rows, exceed_count, unavailable_count, carried_count, overall_avg, overall_median
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
