"""Branch coverage for edf_bill_fetcher/writers/_helpers.py error/edge paths.

Targets the module's writer-specific helpers: the statsmodels import guard,
tariff/quality analysis, evidence-index builder, forecast/volatility/anomaly
helpers, SAP back-billing detection + matching, dispute-flag computation, and
the small label/amount parse helpers.

The Wave 4 coverage gap was 112 missed statements (68% module coverage). This
file drives the module to >=95% by exercising the defensive branches, empty
inputs, and exception paths that the mainline writer tests never reach.
"""

from __future__ import annotations

import builtins
import importlib
import sys
from collections.abc import Sequence
from datetime import datetime
from typing import Any

import numpy as np
import pandas as pd
import pytest

from edf_bill_fetcher.models.events import SapBackBillingEvent
from edf_bill_fetcher.processors.detection import _assess_reason
from edf_bill_fetcher.writers import _helpers as h

# ---------------------------------------------------------------------------
# Module-level guard: HAS_STATSMODELS falls back to False on ImportError
# ---------------------------------------------------------------------------


def test_statsmodels_import_guard_falls_back(monkeypatch: pytest.MonkeyPatch) -> None:
    """HAS_STATSMODELS is False when the statsmodels import fails."""
    original = h.HAS_STATSMODELS
    real_import = builtins.__import__

    def _block_statsmodels(
        name: str,
        globals: dict[str, Any] | None = None,
        locals: dict[str, Any] | None = None,
        fromlist: Sequence[str] = (),
        level: int = 0,
    ) -> Any:
        if name == "statsmodels" or name.startswith("statsmodels."):
            raise ImportError("blocked for test")
        return real_import(name, globals, locals, fromlist, level)

    monkeypatch.setattr(builtins, "__import__", _block_statsmodels)
    for key in [k for k in sys.modules if k == "statsmodels" or k.startswith("statsmodels.")]:
        monkeypatch.delitem(sys.modules, key)

    importlib.reload(h)
    assert h.HAS_STATSMODELS is False

    # Restore the real import so the rest of the suite sees the real flag.
    monkeypatch.setattr(builtins, "__import__", real_import)
    importlib.reload(h)
    assert h.HAS_STATSMODELS == original


# ---------------------------------------------------------------------------
# _analyze_tariff_impact / _data_quality_report edge inputs
# ---------------------------------------------------------------------------


def test_analyze_tariff_impact_missing_columns() -> None:
    df = pd.DataFrame({"Invoice #": ["KI-1"], "Amount (£)": [100.0]})
    assert h._analyze_tariff_impact(df) == {}


def test_analyze_tariff_impact_missing_unit_rate_column() -> None:
    df = pd.DataFrame({"Tariff": ["A"], "Amount (£)": [100.0]})
    assert h._analyze_tariff_impact(df) == {}


def test_data_quality_report_empty_frame() -> None:
    assert h._data_quality_report(pd.DataFrame()) == {}


# ---------------------------------------------------------------------------
# _disclosed_label / _reading_type_to_aem all four quadrants
# ---------------------------------------------------------------------------


def test_disclosed_label_all_combinations() -> None:
    assert h._disclosed_label(True, True) == "Admitted + overlap"
    assert h._disclosed_label(True, False) == "Admitted phrase"
    assert h._disclosed_label(False, True) == "Period overlap"
    assert h._disclosed_label(False, False) == ""


def test_reading_type_to_aem_all_values() -> None:
    assert h._reading_type_to_aem("Actual") == "A"
    assert h._reading_type_to_aem("Estimated") == "E"
    assert h._reading_type_to_aem("Smart") == "A"
    assert h._reading_type_to_aem("Unknown") == "E"


# ---------------------------------------------------------------------------
# build_evidence_index
# ---------------------------------------------------------------------------


def test_build_evidence_index_empty_or_invalid_inputs() -> None:
    assert h.build_evidence_index(None) == {}
    assert h.build_evidence_index(pd.DataFrame()) == {}
    assert h.build_evidence_index(["not", "a", "frame"]) == {}


def test_build_evidence_index_populates_keys_and_skips_bad_rows() -> None:
    df = pd.DataFrame(
        [
            # Full row -> both inv: and amt_days: keys.
            {
                "Invoice #": "KI-100",
                "Amount (£)": "120.50",
                "Period From": "01/01/2024",
                "Period To": "31/01/2024",
            },
            # N/A invoice -> invoice key skipped, amount key still set.
            {
                "Invoice #": "N/A",
                "Amount (£)": "50.00",
                "Period From": "01/02/2024",
                "Period To": "29/02/2024",
            },
            # Non-string invoice -> invoice key skipped.
            {
                "Invoice #": 12345,
                "Amount (£)": "10.00",
                "Period From": "01/03/2024",
                "Period To": "31/03/2024",
            },
            # Bad period -> skipped entirely.
            {
                "Invoice #": "KI-101",
                "Amount (£)": "5.00",
                "Period From": "garbage",
                "Period To": "31/03/2024",
            },
            # Bad amount -> skipped entirely.
            {
                "Invoice #": "KI-102",
                "Amount (£)": "abc",
                "Period From": "01/04/2024",
                "Period To": "30/04/2024",
            },
        ]
    )
    index = h.build_evidence_index(df, header_row_offset=1)
    assert index["inv:KI-100"] == 2
    assert index["amt_days:120.50|30"] == 2
    assert "inv:N/A" not in index
    assert index["amt_days:50.00|28"] == 3  # Feb 1 -> Feb 29 = 28 days
    assert index["amt_days:10.00|30"] == 4  # Mar 1 -> Mar 31 = 30 days
    # Bad rows still get their invoice key before the skip.
    assert index["inv:KI-101"] == 5
    assert index["inv:KI-102"] == 6
    assert len([k for k in index if k.startswith("amt_days:")]) == 3


# ---------------------------------------------------------------------------
# Volatility / anomaly detection edge cases
# ---------------------------------------------------------------------------


def test_compute_volatility_rolling_std() -> None:
    series = pd.Series([10.0, 12.0, 11.0, 13.0])
    result = h._compute_volatility(series, window=2)
    assert isinstance(result, pd.Series)
    assert len(result) == 4
    assert not result.isna().all()


def test_zscore_anomalies_too_short() -> None:
    result = h._zscore_anomalies(pd.Series([1.0, 2.0]))
    assert list(result) == [False, False]


def test_zscore_anomalies_zero_std() -> None:
    result = h._zscore_anomalies(pd.Series([5.0, 5.0, 5.0]))
    assert list(result) == [False, False, False]


def test_zscore_anomalies_normal_case() -> None:
    # n=10 so the max achievable z-score exceeds the 2.5 threshold.
    result = h._zscore_anomalies(pd.Series([1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 100.0]))
    assert list(result) == [False] * 9 + [True]


def test_iqr_anomalies_too_short() -> None:
    result = h._iqr_anomalies(pd.Series([1.0, 2.0, 3.0]))
    assert list(result) == [False, False, False]


def test_iqr_anomalies_zero_iqr() -> None:
    result = h._iqr_anomalies(pd.Series([5.0, 5.0, 5.0, 5.0]))
    assert list(result) == [False, False, False, False]


def test_iqr_anomalies_normal_case() -> None:
    result = h._iqr_anomalies(pd.Series([1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 100.0]))
    assert list(result) == [False] * 9 + [True]


# ---------------------------------------------------------------------------
# Linear / Holt-Winters forecast helpers
# ---------------------------------------------------------------------------


def test_linear_forecast_pair_too_short() -> None:
    fitted, forecast = h._linear_forecast_pair(pd.Series([1.0, 2.0]))
    assert fitted is None
    assert forecast is None


def test_linear_forecast_pair_all_nan() -> None:
    fitted, forecast = h._linear_forecast_pair(pd.Series([np.nan, np.nan, np.nan]))
    assert fitted is None
    assert forecast is None


def test_linear_forecast_pair_polyfit_exception(monkeypatch: pytest.MonkeyPatch) -> None:
    def _boom(x: object, y: object, deg: int) -> object:
        raise np.linalg.LinAlgError("SVD did not converge")

    monkeypatch.setattr(np, "polyfit", _boom)
    fitted, forecast = h._linear_forecast_pair(pd.Series([1.0, 2.0, 3.0, 4.0, 5.0]))
    assert fitted is None
    assert forecast is None


def test_linear_forecast_pair_happy_path() -> None:
    fitted, forecast = h._linear_forecast_pair(pd.Series([1.0, 2.0, 3.0, 4.0, 5.0]), steps=6)
    assert fitted is not None
    assert len(fitted) == 5
    assert forecast is not None
    assert len(forecast) == 6


@pytest.mark.skipif(not h.HAS_STATSMODELS, reason="statsmodels not installed")
def test_holt_winters_forecast_pair_too_short() -> None:
    fitted, forecast = h._holt_winters_forecast_pair(pd.Series([1.0, 2.0, 3.0]))
    assert fitted is None
    assert forecast is None


@pytest.mark.skipif(not h.HAS_STATSMODELS, reason="statsmodels not installed")
def test_holt_winters_forecast_pair_clean_series_too_short() -> None:
    # 3 valid points after dropping NaN -> None, None.
    fitted, forecast = h._holt_winters_forecast_pair(pd.Series([1.0, 2.0, 3.0, np.nan, np.nan]))
    assert fitted is None
    assert forecast is None


@pytest.mark.skipif(not h.HAS_STATSMODELS, reason="statsmodels not installed")
def test_holt_winters_forecast_pair_auto_seasonal_periods() -> None:
    # len >= 8 with seasonal_periods=None -> seasonal_periods = min(12, len//2).
    series = pd.Series(
        [float(i) for i in range(16)], index=pd.date_range("2024-01-01", periods=16, freq="D")
    )
    fitted, forecast = h._holt_winters_forecast_pair(series, steps=6)
    assert fitted is not None
    assert len(fitted) == 16
    assert forecast is not None
    assert len(forecast) == 6


def test_holt_winters_forecast_pair_short_series_no_seasonal() -> None:
    # len 4-7 with seasonal_periods=None -> seasonal stays None (no seasonal component).
    series = pd.Series(
        [float(i) for i in range(6)], index=pd.date_range("2024-01-01", periods=6, freq="D")
    )
    fitted, forecast = h._holt_winters_forecast_pair(series, steps=6)
    assert fitted is not None
    assert len(fitted) == 6
    assert forecast is not None
    assert len(forecast) == 6


@pytest.mark.skipif(not h.HAS_STATSMODELS, reason="statsmodels not installed")
def test_holt_winters_forecast_pair_exception(monkeypatch: pytest.MonkeyPatch) -> None:
    import edf_bill_fetcher.processors.forecasting as fp

    class _BoomModel:
        def __init__(self, *args: object, **kwargs: object) -> None:
            pass

        def fit(self, **kwargs: object) -> object:
            raise RuntimeError("fit failed")

    monkeypatch.setattr(fp, "ExponentialSmoothing", _BoomModel)
    fitted, forecast = fp._holt_winters_forecast_pair(pd.Series([1.0, 2.0, 3.0, 4.0, 5.0]))
    assert fitted is None
    assert forecast is None


def test_linear_forecast_legacy_entry_point() -> None:
    forecast = h._linear_forecast(pd.Series([1.0, 2.0, 3.0, 4.0, 5.0]), steps=6)
    assert forecast is not None
    assert len(forecast) == 6


@pytest.mark.skipif(not h.HAS_STATSMODELS, reason="statsmodels not installed")
def test_holt_winters_forecast_legacy_entry_point() -> None:
    series = pd.Series(
        [float(i) for i in range(16)], index=pd.date_range("2024-01-01", periods=16, freq="D")
    )
    forecast = h._holt_winters_forecast(series, steps=6)
    assert forecast is not None
    assert len(forecast) == 6


# ---------------------------------------------------------------------------
# _detect_payment_patterns
# ---------------------------------------------------------------------------


def test_detect_payment_patterns_no_payments() -> None:
    df = pd.DataFrame(
        {
            "Entry Type": ["New Bill", "New Bill"],
            "Date": ["01/01/2024", "01/02/2024"],
            "Amount (£)": [100.0, 110.0],
        }
    )
    assert h._detect_payment_patterns(df) == {}


def test_detect_payment_patterns_happy_path() -> None:
    df = pd.DataFrame(
        {
            "Entry Type": ["Payment", "Payment", "New Bill"],
            "Date": ["01/01/2024", "01/02/2024", "15/02/2024"],
            "Amount (£)": [50.0, 50.0, 200.0],
            "Period Charge (£)": [None, None, 200.0],
        }
    )
    result = h._detect_payment_patterns(df)
    assert result["count"] == 2
    assert result["total_paid"] == 100.0
    assert result["avg_interval_days"] == 31.0


# ---------------------------------------------------------------------------
# _assess_reason / _parse_amount_for_event / _confidence_band
# ---------------------------------------------------------------------------


def test_assess_reason_admitted() -> None:
    narrative = _assess_reason(
        "KI-1",
        datetime(2024, 3, 1),
        35,
        True,
        datetime(2023, 1, 1),
        datetime(2024, 2, 5),
    )
    assert "KI-1" in narrative
    assert "billed on 01 Mar 2024" in narrative
    assert "35 days of consumption were supplied more than 12 months before the bill" in narrative
    assert "admits a cancellation/reversal" in narrative


def test_assess_reason_not_admitted() -> None:
    narrative = _assess_reason(
        "KI-2",
        datetime(2024, 3, 1),
        35,
        False,
        datetime(2023, 1, 1),
        datetime(2024, 2, 5),
    )
    assert "KI-2" in narrative
    assert "billed on 01 Mar 2024" in narrative
    assert "35 days of consumption were supplied more than 12 months before the bill" in narrative
    assert "No admit-phrase was found" in narrative


def test_parse_amount_for_event_variants() -> None:
    assert h._parse_amount_for_event(None) == 0.0
    assert h._parse_amount_for_event("") == 0.0
    assert h._parse_amount_for_event("£abc") == 0.0
    assert h._parse_amount_for_event("1,234.56") == 1234.56
    assert h._parse_amount_for_event(100) == 100.0
    assert h._parse_amount_for_event("£1,234.56") == 1234.56


def test_confidence_band_all_levels() -> None:
    assert h._confidence_band(75) == "High"
    assert h._confidence_band(40) == "Medium"
    assert h._confidence_band(10) == "Low"
    assert h._confidence_band(9) is None


# ---------------------------------------------------------------------------
# detect_sap_back_billing_events edge cases
# ---------------------------------------------------------------------------


def test_detect_sap_back_billing_events_empty() -> None:
    assert h.detect_sap_back_billing_events([]) == []


def test_detect_sap_back_billing_events_clearing_doc_skipped() -> None:
    for bad_cd in ("", "NA", "None", "*"):
        rows = [
            {
                "Statistical Key Flag": "",
                "Clearing Document": bad_cd,
                "Clearing Date": "01/03/2024",
                "Clearing Reason": "R",
                "Amount": "10.00",
                "Transaction Text": "T",
                "Posting Date": "01/03/2024",
            }
            for _ in range(4)
        ]
        assert h.detect_sap_back_billing_events(rows) == []


def test_detect_sap_back_billing_events_debt_mgmt_filtered() -> None:
    rows = [
        {
            "Statistical Key Flag": "Installment Plan Item",
            "Clearing Document": "CD001",
            "Clearing Date": "01/03/2024",
            "Clearing Reason": "R",
            "Amount": "10.00",
            "Transaction Text": "T",
            "Posting Date": "01/03/2024",
        }
        for _ in range(4)
    ]
    assert h.detect_sap_back_billing_events(rows) == []


def test_detect_sap_back_billing_events_cluster_below_min_size() -> None:
    rows = [
        {
            "Statistical Key Flag": "",
            "Clearing Document": "CD001",
            "Clearing Date": "01/03/2024",
            "Clearing Reason": "R",
            "Amount": "10.00",
            "Transaction Text": "T",
            "Posting Date": "01/03/2024",
        }
        for _ in range(3)
    ]
    assert h.detect_sap_back_billing_events(rows, min_cluster_size=4) == []


def test_detect_sap_back_billing_events_happy_path() -> None:
    rows = [
        {
            "Statistical Key Flag": "",
            "Clearing Document": "CD001",
            "Clearing Date": "01/03/2024",
            "Clearing Reason": "Back-bill",
            "Amount": "10.00",
            "Transaction Text": "Credit for Consum Billing",
            "Posting Date": "01/03/2024",
        }
        for _ in range(4)
    ]
    events = h.detect_sap_back_billing_events(rows)
    assert len(events) == 1
    assert events[0].clearing_doc == "CD001"
    assert events[0].net_amount == 40.0
    assert events[0].has_credit_for_consum_billing is True


# ---------------------------------------------------------------------------
# match_sap_events_to_edf
# ---------------------------------------------------------------------------


def _event(
    *,
    cdate: object = "2024-03-15",
    net: float = 100.0,
    rows: list[dict] | None = None,
    posting: str | None = None,
) -> SapBackBillingEvent:
    # Task 6 (commit d06909c) made Posting Date the preferred date axis.
    # When ``posting`` is given, the rows carry a Posting Date so the
    # matcher's Posting Date branch is exercised; otherwise the rows
    # omit it and the matcher falls back to Clearing Date.
    if rows is None:
        row: dict = {"Amount": f"{net:.2f}"}
        if posting is not None:
            row["Posting Date"] = posting
        rows = [row]
    return SapBackBillingEvent(
        clearing_doc="CD1",
        clearing_date=pd.Timestamp(cdate) if cdate is not None else pd.NaT,
        clearing_reason="Back-bill",
        rows=rows,
        net_amount=net,
        has_credit_for_consum_billing=False,
        has_account_maintenance=False,
        largest_single_posting=net,
        posting_date_range=("01/03/2024", "15/03/2024"),
        evidence_trail="trail",
    )


def _rec(
    *,
    invoice: str = "KI-100",
    pf: str = "01/03/2024",
    pt: str = "31/03/2024",
    amt: str = "100.00",
    period_charge: str | None = None,
) -> dict:
    # Task 7 (commit dcdf6eb) made Period Charge (£) the canonical amount
    # axis; Amount (£) is the fallback.  Default period_charge to amt so
    # the canonical path is exercised; pass period_charge explicitly to
    # drive the fallback branch.
    return {
        "Invoice #": invoice,
        "Period From": pf,
        "Period To": pt,
        "Period Charge (£)": amt if period_charge is None else period_charge,
        "Amount (£)": amt,
    }


def test_match_sap_events_empty_inputs() -> None:
    assert h.match_sap_events_to_edf([], [_rec()]) == []
    assert h.match_sap_events_to_edf([_event()], []) == []


def test_match_sap_events_skips_invalid_records() -> None:
    events = [_event(cdate="2024-03-15", net=100.0)]
    records = [
        # Blank invoice -> skipped.
        _rec(invoice="", amt="100.00"),
        # N/A invoice -> skipped.
        _rec(invoice="N/A", amt="100.00"),
        # Non-numeric amount -> amt coerced to 0.0, still matched on date only.
        _rec(invoice="KI-1", amt="£abc"),
        # Both periods missing -> skipped.
        _rec(invoice="KI-2", pf="N/A", pt="N/A"),
    ]
    matches = h.match_sap_events_to_edf(events, records)
    # KI-1: date in span, amount 0 -> day band 14 -> date_score 5 -> total 5 -> None band.
    # KI-2 skipped. Blank/N/A skipped. So no matches.
    assert matches == []


def test_match_sap_events_nat_clearing_date_skipped() -> None:
    events = [_event(cdate=None, net=100.0)]
    assert h.match_sap_events_to_edf(events, [_rec()]) == []


def test_match_sap_events_period_to_missing_uses_from() -> None:
    events = [_event(cdate="2024-03-10", net=100.0)]
    records = [_rec(pf="01/03/2024", pt="")]  # empty Period To -> NaT -> uses pf
    matches = h.match_sap_events_to_edf(events, records)
    assert len(matches) == 1
    # pf-only: date_delta = |Mar10 - Mar1| = 9 days.
    assert matches[0].date_delta_days == 9


def test_match_sap_events_high_in_span_amount_band() -> None:
    # net < 1.0 with a row whose amount lands in the 0.05 band -> amount_score 40.
    events = [_event(cdate="2024-03-15", net=0.5, rows=[{"Amount": "101.00"}])]
    records = [_rec(amt="100.00")]
    matches = h.match_sap_events_to_edf(events, records)
    assert len(matches) == 1
    assert matches[0].confidence_band == "High"
    assert "inside EDF period" in matches[0].notes


def test_match_sap_events_amount_band_loop_back() -> None:
    # rel_delta 0.07 -> first band (0.05) fails, second (0.25) hits -> 20.
    events = [_event(cdate="2024-03-15", net=0.5, rows=[{"Amount": "107.00"}])]
    records = [_rec(amt="100.00")]
    matches = h.match_sap_events_to_edf(events, records)
    assert len(matches) == 1
    # 20 (amount) + 50 (in-span) = 70 -> Medium.
    assert matches[0].confidence_band == "Medium"
    assert matches[0].confidence_score == 70


def test_match_sap_events_amount_band_no_match_falls_to_date() -> None:
    # rel_delta 0.60 -> no band matches -> amount_score 0 -> date-only scoring.
    events = [_event(cdate="2024-03-15", net=0.5, rows=[{"Amount": "160.00"}])]
    records = [_rec(amt="100.00")]
    matches = h.match_sap_events_to_edf(events, records)
    # in-span, amount 0 -> nearest_delta 14 -> 5 -> total 5 -> None band.
    assert matches == []


def test_match_sap_events_all_rows_below_floor_skips_bands() -> None:
    # All rows abs < 1 -> best_rel_delta stays inf -> amount band loop skipped.
    events = [_event(cdate="2024-03-15", net=0.5, rows=[{"Amount": "0.20"}, {"Amount": "0.30"}])]
    records = [_rec(amt="100.00")]
    matches = h.match_sap_events_to_edf(events, records)
    # in-span, amount 0 -> nearest_delta 14 -> 5 -> total 5 -> None band.
    assert matches == []


def test_match_sap_events_in_span_day_band_no_hit() -> None:
    # Event mid-period (15 days from both edges) -> nearest_delta 15 > 14 -> no band.
    events = [_event(cdate="2024-03-16", net=0.5, rows=[{"Amount": "160.00"}])]
    records = [_rec(amt="100.00")]
    matches = h.match_sap_events_to_edf(events, records)
    assert matches == []


def test_match_sap_events_ratio_bands() -> None:
    records = [
        _rec(invoice="KI-R1", amt="100.00"),  # ratio 1.00 -> 40
        _rec(invoice="KI-R2", amt="125.00"),  # ratio 0.80 -> 20
        _rec(invoice="KI-R3", amt="180.00"),  # ratio 0.556 -> 5
    ]
    matches = h.match_sap_events_to_edf([_event(cdate="2024-03-15", net=100.0)], records)
    scores = {m.edf_record["Invoice #"]: m.confidence_score for m in matches}
    assert scores["KI-R1"] == 90  # 40 + 50
    assert scores["KI-R2"] == 70  # 20 + 50
    assert scores["KI-R3"] == 55  # 5 + 50


def test_match_sap_events_out_of_span_day_bands() -> None:
    # Event 2 months before the period -> out of span, date_delta_days = 45 -> no band.
    events = [_event(cdate="2024-01-15", net=100.0)]
    records = [_rec(amt="100.00")]
    matches = h.match_sap_events_to_edf(events, records)
    assert len(matches) == 1
    assert matches[0].confidence_band == "Medium"
    assert "Within 76d of period-end" in matches[0].notes


def test_match_sap_events_low_out_of_span_note() -> None:
    # 3 days past period end -> date band 3 -> 25; amount far off -> 0 -> Low.
    events = [_event(cdate="2024-04-03", net=100.0)]
    records = [_rec(amt="999.00")]
    matches = h.match_sap_events_to_edf(events, records)
    assert len(matches) == 1
    assert matches[0].confidence_band == "Low"
    assert "Within 3d of period-end; may be coincidental" in matches[0].notes


def test_match_sap_events_low_in_span_coincidental_note() -> None:
    # Single-day period containing the clearing date; amount far off -> Low band
    # with the "amounts do not correspond" note (band downgrade from Medium).
    events = [_event(cdate="2024-03-15", net=100.0)]
    records = [_rec(pf="15/03/2024", pt="15/03/2024", amt="999.00")]
    matches = h.match_sap_events_to_edf(events, records)
    assert len(matches) == 1
    assert matches[0].confidence_band == "Low"
    assert "amounts do not correspond" in matches[0].notes


# ---------------------------------------------------------------------------
# handle_cluster_unmatched (Task 8 — Decision 4)
# Spec §3.3: a SAP event whose Posting Date falls inside a back-billing
# cluster window but achieves no amount-band agreement with any in-cluster
# invoice is tagged as an internal mechanism of that cluster.
# ---------------------------------------------------------------------------


def test_handle_cluster_unmatched_tags_when_no_amount_agreement() -> None:
    # Posting Date inside the cluster window; SAP net 999 vs in-cluster
    # invoices 100/200 -> no amount-band agreement -> tagged.
    ev = SapBackBillingEvent(
        clearing_doc="CD-TAG",
        clearing_date=pd.Timestamp("2024-03-15"),
        clearing_reason="Back-bill",
        rows=[{"Posting Date": "2024-03-15", "Amount": "999.00"}],
        net_amount=999.0,
        has_credit_for_consum_billing=False,
        has_account_maintenance=False,
        largest_single_posting=999.0,
        posting_date_range=("2024-03-15", "2024-03-15"),
        evidence_trail="",
    )
    clusters = [
        {
            "name": "T60/T61",
            "posting_date_start": "2024-03-01",
            "posting_date_end": "2024-03-31",
            "invoices": [
                {"Invoice #": "T60", "Period Charge (£)": 100.0},
                {"Invoice #": "T61", "Period Charge (£)": 200.0},
            ],
        }
    ]
    result = h.handle_cluster_unmatched(ev, clusters)
    assert result is not None
    assert result["Matched EDF Invoice #"] == "T60/T61 internal mechanism"
    assert result["Confidence"] == 0
    assert "Posting Date inside cluster window" in result["Notes"]
    assert "£999.00" in result["Notes"]


def test_handle_cluster_unmatched_none_when_amount_agrees() -> None:
    # SAP net 100 vs in-cluster invoice 100 -> within 50% -> not tagged.
    ev = SapBackBillingEvent(
        clearing_doc="CD-AGREE",
        clearing_date=pd.Timestamp("2024-03-15"),
        clearing_reason="Back-bill",
        rows=[{"Posting Date": "2024-03-15", "Amount": "100.00"}],
        net_amount=100.0,
        has_credit_for_consum_billing=False,
        has_account_maintenance=False,
        largest_single_posting=100.0,
        posting_date_range=("2024-03-15", "2024-03-15"),
        evidence_trail="",
    )
    clusters = [
        {
            "name": "T62",
            "posting_date_start": "2024-03-01",
            "posting_date_end": "2024-03-31",
            "invoices": [{"Invoice #": "T62", "Period Charge (£)": 100.0}],
        }
    ]
    result = h.handle_cluster_unmatched(ev, clusters)
    assert result is None


def test_handle_cluster_unmatched_none_when_window_excludes_date() -> None:
    # Posting Date outside the cluster window -> None.
    ev = SapBackBillingEvent(
        clearing_doc="CD-OUTSIDE",
        clearing_date=pd.Timestamp("2024-06-15"),
        clearing_reason="Back-bill",
        rows=[{"Posting Date": "2024-06-15", "Amount": "999.00"}],
        net_amount=999.0,
        has_credit_for_consum_billing=False,
        has_account_maintenance=False,
        largest_single_posting=999.0,
        posting_date_range=("2024-06-15", "2024-06-15"),
        evidence_trail="",
    )
    clusters = [
        {
            "name": "T63",
            "posting_date_start": "2024-03-01",
            "posting_date_end": "2024-03-31",
            "invoices": [{"Invoice #": "T63", "Period Charge (£)": 100.0}],
        }
    ]
    result = h.handle_cluster_unmatched(ev, clusters)
    assert result is None


# ---------------------------------------------------------------------------
# compute_dispute_flags
# ---------------------------------------------------------------------------


def _flag_df(
    dates: list[str],
    amounts: list[float],
    *,
    readings: list[str] | None = None,
    entry_types: list[str] | None = None,
    period_charges: list[float | None] | None = None,
    bad_dt: bool = False,
    bad_amounts: bool = False,
) -> pd.DataFrame:
    df = pd.DataFrame(
        {
            "Date": dates,
            "_dt": (
                ["not-a-date"] * len(dates) if bad_dt else pd.to_datetime(dates, dayfirst=True)
            ),
            "Amount (£)": (["£100"] * len(dates) if bad_amounts else amounts),
        }
    )
    if readings is not None:
        df["Reading"] = readings
    if entry_types is not None:
        df["Entry Type"] = entry_types
    if period_charges is not None:
        df["Period Charge (£)"] = period_charges
    return df


def test_compute_dispute_flags_too_short() -> None:
    df = _flag_df(["01/01/2024"], [100.0])
    flags, counts = h.compute_dispute_flags(df)
    assert flags == []
    assert counts == {"HIGH": 0, "MEDIUM": 0, "INFO": 0}


def test_compute_dispute_flags_large_jump_happy() -> None:
    df = _flag_df(["01/01/2024", "01/02/2024"], [100.0, 200.0])
    flags, counts = h.compute_dispute_flags(df)
    names = [f[0] for f in flags]
    assert "LARGE JUMP" in names
    assert counts["HIGH"] == 1


def test_compute_dispute_flags_billing_gap_medium_and_high() -> None:
    df_med = _flag_df(["01/01/2024", "15/03/2024"], [100.0, 100.0])  # 74 days
    flags_med, _ = h.compute_dispute_flags(df_med)
    gap_med = [f for f in flags_med if f[0] == "BILLING GAP"]
    assert len(gap_med) == 1
    assert gap_med[0][4] == "MEDIUM"

    df_high = _flag_df(["01/01/2024", "15/05/2024"], [100.0, 100.0])  # 135 days
    flags_high, _ = h.compute_dispute_flags(df_high)
    gap_high = [f for f in flags_high if f[0] == "BILLING GAP"]
    assert gap_high[0][4] == "HIGH"


def test_compute_dispute_flags_estimated_run_flush_and_ongoing() -> None:
    # 3 estimated then actual -> flush at the non-estimated row (run>=3).
    df_flush = _flag_df(
        ["01/01/2024", "01/02/2024", "01/03/2024", "01/04/2024"],
        [100.0, 100.0, 100.0, 100.0],
        readings=["Estimated", "Estimated", "Estimated", "Actual"],
    )
    flags_flush, _ = h.compute_dispute_flags(df_flush)
    runs = [f for f in flags_flush if f[0] == "ESTIMATED RUN"]
    assert len(runs) == 1
    assert "ongoing" not in runs[0][3]

    # All estimated -> ongoing flush at end.
    df_ongoing = _flag_df(
        ["01/01/2024", "01/02/2024", "01/03/2024"],
        [100.0, 100.0, 100.0],
        readings=["Estimated", "Estimated", "Estimated"],
    )
    flags_ongoing, _ = h.compute_dispute_flags(df_ongoing)
    runs_ongoing = [f for f in flags_ongoing if f[0] == "ESTIMATED RUN"]
    assert len(runs_ongoing) == 1
    assert "ongoing" in runs_ongoing[0][3]


def test_compute_dispute_flags_estimated_run_below_three() -> None:
    df = _flag_df(
        ["01/01/2024", "01/02/2024"],
        [100.0, 100.0],
        readings=["Estimated", "Actual"],
    )
    flags, _ = h.compute_dispute_flags(df)
    assert all(f[0] != "ESTIMATED RUN" for f in flags)


def test_compute_dispute_flags_high_daily_rate() -> None:
    df = _flag_df(["01/01/2024", "01/02/2024"], [100.0, 200.0])
    flags, _ = h.compute_dispute_flags(df, mean_daily=1.0)
    names = [f[0] for f in flags]
    assert "HIGH DAILY RATE" in names


def test_compute_dispute_flags_balance_reduction() -> None:
    df = _flag_df(["01/01/2024", "01/02/2024"], [1000.0, 200.0])
    flags, counts = h.compute_dispute_flags(df)
    names = [f[0] for f in flags]
    assert "BALANCE REDUCTION" in names
    assert counts["INFO"] == 1


def test_compute_dispute_flags_reconciliation_mismatch() -> None:
    # Balance delta 0 vs period charge 200 -> diff 200 > threshold -> HIGH.
    df = _flag_df(
        ["01/01/2024", "01/02/2024"],
        [100.0, 100.0],
        entry_types=["New Bill", "New Bill"],
        period_charges=[None, 200.0],
    )
    flags, _ = h.compute_dispute_flags(df)
    names = [f[0] for f in flags]
    assert "RECONCILIATION MISMATCH" in names


def test_compute_dispute_flags_reconciliation_within_threshold() -> None:
    # Balance delta 50 vs period charge 55 -> diff 5 <= threshold -> no flag.
    df = _flag_df(
        ["01/01/2024", "01/02/2024"],
        [100.0, 150.0],
        entry_types=["New Bill", "New Bill"],
        period_charges=[None, 55.0],
    )
    flags, _ = h.compute_dispute_flags(df)
    assert all(f[0] != "RECONCILIATION MISMATCH" for f in flags)


def test_compute_dispute_flags_reconciliation_non_float_pc() -> None:
    # Non-float Period Charge -> inner try continues (line 746-747).
    df = pd.DataFrame(
        {
            "Date": ["01/01/2024", "01/02/2024"],
            "_dt": pd.to_datetime(["01/01/2024", "01/02/2024"], dayfirst=True),
            "Amount (£)": [100.0, 150.0],
            "Entry Type": ["New Bill", "New Bill"],
            "Period Charge (£)": [None, "£55"],
        }
    )
    flags, _ = h.compute_dispute_flags(df)
    assert all(f[0] != "RECONCILIATION MISMATCH" for f in flags)


def test_compute_dispute_flags_exception_paths() -> None:
    # bad_dt + bad_amounts + mean_daily>0 + Period Charge present:
    # exercises every `except (ValueError, TypeError, KeyError, ...): pass`.
    df = _flag_df(
        ["01/01/2024", "01/02/2024"],
        [100.0, 200.0],
        readings=["Actual", "Actual"],
        entry_types=["New Bill", "New Bill"],
        period_charges=[55.0, 55.0],
        bad_dt=True,
        bad_amounts=True,
    )
    flags, counts = h.compute_dispute_flags(df, mean_daily=1.0)
    assert flags == []
    assert counts == {"HIGH": 0, "MEDIUM": 0, "INFO": 0}
