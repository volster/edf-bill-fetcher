"""Typed computed-report results shared by the Excel writer and the PDF/DOCX reporters."""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
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
