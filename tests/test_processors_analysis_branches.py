"""Branch coverage for edf_bill_fetcher.processors.analysis.

Targets the lines not exercised by tests/test_analysis.py: dispute-flag
warning branches, ESTIMATED RUN terminal append, BALANCE REDUCTION,
RECONCILIATION MISMATCH guard paths, payment-period-charge fallback,
tariff empty-after-dropna, and the disclosed/reversal/reading helpers.
"""

import pandas as pd
import pytest

from edf_bill_fetcher.processors.analysis import (
    _analyze_tariff_impact,
    _detect_payment_patterns,
    _disclosed_label,
    _reading_type_to_aem,
    _reversal_match,
    compute_dispute_flags,
)


def _df(rows, extra_cols=None):
    """Build a sorted dispute-flag DataFrame from (date, amount, reading) rows."""
    df = pd.DataFrame(
        {
            "Date": [r[0] for r in rows],
            "_dt": pd.to_datetime([r[0] for r in rows]),
            "Amount (£)": [r[1] for r in rows],
            "Reading": [r[2] if len(r) > 2 else "Actual" for r in rows],
        }
    )
    if extra_cols:
        for col, values in extra_cols.items():
            df[col] = values
    return df


class TestDisputeFlagWarningBranches:
    """compute_dispute_flags non-happy paths — warning emission."""

    def test_billing_gap_warning_on_bad_row(self):
        # `_dt` non-datetime on one row → subtraction raises → BILLING_GAP warns (line 95-96)
        df = pd.DataFrame(
            {
                "Date": ["2024-01-01", "2024-02-15", "2024-03-01"],
                "_dt": [pd.Timestamp("2024-01-01"), "garbage", pd.Timestamp("2024-03-01")],
                "Amount (£)": [100.0, 200.0, 300.0],
            }
        )
        with pytest.warns(UserWarning, match="compute_dispute_flags\\[BILLING_GAP\\]"):
            flags, counts = compute_dispute_flags(df)
        assert counts == {"HIGH": 0, "MEDIUM": 0, "INFO": 0}

    def test_large_jump_warning_on_bad_row(self):
        df = pd.DataFrame(
            {
                "Date": ["2024-01-01", "2024-02-01"],
                "_dt": pd.to_datetime(["2024-01-01", "2024-02-01"]),
                "Amount (£)": ["not-a-number", 100.0],
            }
        )
        with pytest.warns(UserWarning, match="LARGE_JUMP"):
            flags, counts = compute_dispute_flags(df)
        assert flags == []
        assert counts == {"HIGH": 0, "MEDIUM": 0, "INFO": 0}

    def test_high_daily_rate_warning(self):
        df = pd.DataFrame(
            {
                "Date": ["2024-01-01", "2024-02-01"],
                "_dt": pd.to_datetime(["2024-01-01", "2024-02-01"]),
                "Amount (£)": ["x", 100.0],
            }
        )
        with pytest.warns(UserWarning, match="HIGH_DAILY_RATE"):
            compute_dispute_flags(df, mean_daily=50.0)

    def test_balance_reduction_warning(self):
        df = pd.DataFrame(
            {
                "Date": ["2024-01-01", "2024-02-01"],
                "_dt": pd.to_datetime(["2024-01-01", "2024-02-01"]),
                "Amount (£)": ["bad", 100.0],
            }
        )
        with pytest.warns(UserWarning, match="BALANCE_REDUCTION"):
            compute_dispute_flags(df)

    def test_reconciliation_warning(self):
        df = _df(
            [("2024-01-01", 100.0), ("2024-02-01", "bad")],
            extra_cols={
                "Entry Type": ["Ongoing Balance", "New Bill"],
                "Period Charge (£)": [None, 100.0],
            },
        )
        with pytest.warns(UserWarning, match="RECONCILIATION_MISMATCH"):
            compute_dispute_flags(df)


class TestEstimatedRunBranches:
    """ESTIMATED RUN — mid-loop flush and ongoing terminal append."""

    def test_estimated_run_flush_on_interrupt(self):
        # 3 estimated, then Actual → mid-loop flush at line 108-109
        df = _df(
            [
                ("2024-01-01", 100.0, "Estimated"),
                ("2024-02-01", 200.0, "Estimated"),
                ("2024-03-01", 300.0, "Estimated"),
                ("2024-04-01", 400.0, "Actual"),
            ]
        )
        flags, counts = compute_dispute_flags(df)
        est = [f for f in flags if f[0] == "ESTIMATED RUN"]
        assert len(est) == 1
        assert "consecutive estimated readings from" in est[0][3]
        assert counts["HIGH"] >= 1

    def test_estimated_run_ongoing_terminal_append(self):
        # Run continues to the end → terminal append at line 120-121 "(ongoing)"
        df = _df(
            [
                ("2024-01-01", 100.0, "Estimated"),
                ("2024-02-01", 200.0, "Estimated"),
                ("2024-03-01", 300.0, "Estimated"),
                ("2024-04-01", 400.0, "Estimated"),
            ]
        )
        flags, _ = compute_dispute_flags(df)
        est = [f for f in flags if f[0] == "ESTIMATED RUN"]
        assert len(est) == 1
        assert "(ongoing)" in est[0][3]

    def test_estimated_run_too_short_no_flag(self):
        df = _df(
            [
                ("2024-01-01", 100.0, "Estimated"),
                ("2024-02-01", 200.0, "Estimated"),
                ("2024-03-01", 300.0, "Actual"),
            ]
        )
        flags, _ = compute_dispute_flags(df)
        assert all(f[0] != "ESTIMATED RUN" for f in flags)


class TestBalanceReductionAndReconciliation:
    """BALANCE REDUCTION flag and RECONCILIATION MISMATCH paths."""

    def test_balance_reduction_flag(self):
        df = _df([("2024-01-01", 1000.0), ("2024-02-01", 100.0)])
        flags, counts = compute_dispute_flags(df)
        br = [f for f in flags if f[0] == "BALANCE REDUCTION"]
        assert len(br) == 1
        assert br[0][4] == "INFO"
        assert counts["INFO"] == 1

    def test_large_jump_stores_delta_as_amount(self):
        """C-6: LARGE JUMP's amount field must be the jump delta, not the balance."""
        df = _df([("2024-01-01", 100.0), ("2024-02-01", 600.0)])
        flags, _ = compute_dispute_flags(df)
        lj = [f for f in flags if f[0] == "LARGE JUMP"]
        assert len(lj) == 1
        assert lj[0][2] == pytest.approx(500.0)  # 600 - 100, positive jump

    def test_balance_reduction_stores_delta_as_amount(self):
        """C-6: BALANCE REDUCTION's amount field must be the reduction size."""
        df = _df([("2024-01-01", 1000.0), ("2024-02-01", 300.0)])
        flags, _ = compute_dispute_flags(df)
        br = [f for f in flags if f[0] == "BALANCE REDUCTION"]
        assert len(br) == 1
        assert br[0][2] == pytest.approx(700.0)  # 1000 - 300, positive reduction

    def test_non_delta_flags_keep_balance_as_amount(self):
        """C-6: non-delta flags (e.g. RECONCILIATION MISMATCH) keep the balance."""
        df = _df(
            [("2024-01-01", 100.0), ("2024-02-01", 500.0)],
            extra_cols={
                "Entry Type": ["Ongoing Balance", "New Bill"],
                "Period Charge (£)": [None, 100.0],
            },
        )
        flags, _ = compute_dispute_flags(df)
        rm = [f for f in flags if f[0] == "RECONCILIATION MISMATCH"]
        assert len(rm) == 1
        assert rm[0][2] == pytest.approx(500.0)  # running balance, not the 400 delta

    def test_reconciliation_mismatch_flagged(self):
        df = _df(
            [("2024-01-01", 100.0), ("2024-02-01", 500.0)],
            extra_cols={
                "Entry Type": ["Ongoing Balance", "New Bill"],
                "Period Charge (£)": [None, 100.0],  # balance delta 400 vs 100 → mismatch
            },
        )
        flags, counts = compute_dispute_flags(df)
        rm = [f for f in flags if f[0] == "RECONCILIATION MISMATCH"]
        assert len(rm) == 1
        assert rm[0][4] == "HIGH"  # diff 300 > 50% of pc 100 → HIGH severity
        assert counts["HIGH"] == 2  # LARGE JUMP (100→500) + RECONCILIATION MISMATCH

    def test_reconciliation_within_threshold_no_flag(self):
        df = _df(
            [("2024-01-01", 100.0), ("2024-02-01", 160.0)],
            extra_cols={
                "Entry Type": ["Ongoing Balance", "New Bill"],
                "Period Charge (£)": [None, 60.0],  # delta 60 == pc 60 → no mismatch
            },
        )
        flags, _ = compute_dispute_flags(df)
        assert all(f[0] != "RECONCILIATION MISMATCH" for f in flags)

    def test_reconciliation_pc_parse_skip(self):
        # Period Charge non-numeric → inner try/except continue (lines 187-188)
        df = _df(
            [("2024-01-01", 100.0), ("2024-02-01", 200.0)],
            extra_cols={
                "Entry Type": ["Ongoing Balance", "New Bill"],
                "Period Charge (£)": [None, "N/A"],
            },
        )
        flags, _ = compute_dispute_flags(df)
        assert all(f[0] != "RECONCILIATION MISMATCH" for f in flags)


class TestPaymentPatternFallback:
    """_detect_payment_patterns — Period Charge fallback branch (line 246-248)."""

    def test_no_period_charge_column_fallback_to_amount(self):
        df = pd.DataFrame(
            {
                "Date": pd.to_datetime(["2024-01-01", "2024-02-01", "2024-03-01"]),
                "Amount (£)": [-200.0, -200.0, -210.0],
                "Entry Type": ["Payment", "Payment", "Payment"],
            }
        )
        result = _detect_payment_patterns(df)
        assert result["count"] == 3
        assert result["total_paid"] == pytest.approx(610.0)

    def test_period_charge_column_preferred(self):
        df = pd.DataFrame(
            {
                "Date": pd.to_datetime(["2024-01-01", "2024-02-01"]),
                "Amount (£)": [-500.0, -600.0],  # running balances — large
                "Period Charge (£)": [50.0, 60.0],  # actual payments
                "Entry Type": ["Payment", "Payment"],
            }
        )
        result = _detect_payment_patterns(df)
        assert result["total_paid"] == pytest.approx(110.0)


class TestTariffImpactEmptyPaths:
    """_analyze_tariff_impact — empty-after-dropna branch (line 280)."""

    def test_all_non_numeric_unit_rates_returns_empty(self):
        df = pd.DataFrame(
            {
                "Tariff": ["T1", "T2"],
                "Unit Rate (p/kWh)": ["N/A", "unknown"],
                "Period Charge (£)": [10.0, 20.0],
                "Date": pd.to_datetime(["2024-01-01", "2024-02-01"]),
            }
        )
        assert _analyze_tariff_impact(df) == {}

    def test_missing_tariff_column_returns_empty(self):
        df = pd.DataFrame({"Date": pd.to_datetime(["2024-01-01"])})
        assert _analyze_tariff_impact(df) == {}


class TestDisclosedLabel:
    """_disclosed_label — all four branch combinations (lines 394-400)."""

    def test_admitted_and_overlaps(self):
        assert _disclosed_label(True, True) == "Admitted + overlap"

    def test_admitted_only(self):
        assert _disclosed_label(True, False) == "Admitted phrase"

    def test_overlaps_only(self):
        assert _disclosed_label(False, True) == "Period overlap"

    def test_neither(self):
        assert _disclosed_label(False, False) == ""


class TestReversalMatch:
    """_reversal_match — guard, amount, and overlap branches (lines 417-440)."""

    def test_none_or_empty_evidence(self):
        assert (
            _reversal_match(
                None, "INV-1", 100.0, pd.Timestamp("2024-01-01"), pd.Timestamp("2024-02-01")
            )
            is False
        )
        empty = pd.DataFrame()
        assert (
            _reversal_match(
                empty, "INV-1", 100.0, pd.Timestamp("2024-01-01"), pd.Timestamp("2024-02-01")
            )
            is False
        )

    def test_missing_entry_type_column(self):
        df = pd.DataFrame({"Amount (£)": [100.0]})
        assert (
            _reversal_match(
                df, "INV-1", 100.0, pd.Timestamp("2024-01-01"), pd.Timestamp("2024-02-01")
            )
            is False
        )

    def test_unparseable_killed_amount(self):
        df = pd.DataFrame({"Entry Type": ["Credit"], "Amount (£)": [100.0]})
        assert (
            _reversal_match(
                df,
                "INV-1",
                "oops",  # type: ignore[arg-type]  # deliberate: exercise the ValueError path
                pd.Timestamp("2024-01-01"),
                pd.Timestamp("2024-02-01"),
            )
            is False
        )

    def test_row_amount_parse_error_skipped(self):
        df = pd.DataFrame({"Entry Type": ["Credit"], "Amount (£)": ["bad"]})
        assert (
            _reversal_match(
                df, "INV-1", 100.0, pd.Timestamp("2024-01-01"), pd.Timestamp("2024-02-01")
            )
            is False
        )

    def test_amount_within_50p_matches(self):
        df = pd.DataFrame({"Entry Type": ["Credit"], "Amount (£)": [100.20]})
        assert (
            _reversal_match(
                df, "INV-1", 100.0, pd.Timestamp("2024-01-01"), pd.Timestamp("2024-02-01")
            )
            is True
        )

    def test_amount_outside_50p_no_match(self):
        df = pd.DataFrame({"Entry Type": ["Credit"], "Amount (£)": [200.0]})
        assert (
            _reversal_match(
                df, "INV-1", 100.0, pd.Timestamp("2024-01-01"), pd.Timestamp("2024-02-01")
            )
            is False
        )

    def test_unparseable_period_accepts_on_amount(self):
        df = pd.DataFrame(
            {
                "Entry Type": ["Credit"],
                "Amount (£)": [100.0],
                "Period From": ["garbage"],
                "Period To": [None],
            }
        )
        assert (
            _reversal_match(
                df, "INV-1", 100.0, pd.Timestamp("2024-01-01"), pd.Timestamp("2024-02-01")
            )
            is True
        )

    def test_overlap_ge_30_days_matches(self):
        df = pd.DataFrame(
            {
                "Entry Type": ["Credit"],
                "Amount (£)": [100.0],
                "Period From": ["2024-01-01"],
                "Period To": ["2024-02-15"],
            }
        )
        killed_pf = pd.Timestamp("2024-01-01")
        killed_pt = pd.Timestamp("2024-02-01")
        assert _reversal_match(df, "INV-1", 100.0, killed_pf, killed_pt) is True

    def test_overlap_below_30_days_no_match(self):
        df = pd.DataFrame(
            {
                "Entry Type": ["Credit"],
                "Amount (£)": [100.0],
                "Period From": ["2024-01-20"],
                "Period To": ["2024-02-01"],
            }
        )
        killed_pf = pd.Timestamp("2024-01-01")
        killed_pt = pd.Timestamp("2024-02-01")
        assert _reversal_match(df, "INV-1", 100.0, killed_pf, killed_pt) is False


class TestReadingTypeToAem:
    """_reading_type_to_aem — all four mappings (lines 445-451)."""

    def test_mappings(self):
        assert _reading_type_to_aem("Actual") == "A"
        assert _reading_type_to_aem("Estimated") == "E"
        assert _reading_type_to_aem("Smart") == "A"
        assert _reading_type_to_aem("Unknown") == "E"
