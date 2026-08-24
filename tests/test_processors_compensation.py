"""TDD tests for the SLC-aware compensation estimator.

Covers ``estimate_compensation`` in ``processors/compensation.py``:
back-billing excess pro-ration, forced-credit-balance (hold) interest,
late-credit interest, refund suppression, no-op inputs, the DISCLAIMER
constant, and config plumbing (credit_hold_days / credit_interest_rate
/ as_of).
"""

from __future__ import annotations

import pandas as pd
import pytest

from edf_bill_fetcher.processors.compensation import DISCLAIMER, estimate_compensation

AS_OF = "2026-06-01"
ROW_KEYS = {
    "category",
    "invoice_ref",
    "date",
    "base_amount",
    "days",
    "rate",
    "indicative_amount",
    "legal_basis",
    "disclaimer",
}


def _record(
    invoice: str,
    date: str,
    amount: float,
    period_from: str,
    period_to: str,
    period_charge: float = 0.0,
) -> dict[str, object]:
    return {
        "Invoice #": invoice,
        "Date": date,
        "Period From": period_from,
        "Period To": period_to,
        "Amount (£)": amount,
        "Period Charge (£)": period_charge,
    }


# ---------- back-billing excess pro-ration ----------


def test_back_billing_excess_pro_ration() -> None:
    """A bill charging for consumption supplied >12 months earlier yields a
    positive indicative excess claim with the SLC 7A / s.84B legal basis.

    Fixture: Period From 01 Jan 2022 -> Period To 28 Feb 2024 (788 days),
    billed 01 Mar 2024.  The 12-month cutoff (02 Mar 2023) falls inside the
    period, so Excess Days = 425 (< Days Billed) and the claim is the
    day-ratio slice of the Period Charge, not the whole charge.
    """
    df = pd.DataFrame(
        [
            _record(
                "KI-0001",
                "01 Mar 2024",
                1200.00,
                "01 Jan 2022",
                "28 Feb 2024",
                period_charge=1200.00,
            )
        ]
    )
    rows = estimate_compensation(df, config={"as_of": AS_OF})
    excess_rows = [r for r in rows if r["category"] == "back_billing_excess"]
    assert len(excess_rows) == 1
    row = excess_rows[0]
    assert set(row.keys()) == ROW_KEYS
    assert row["invoice_ref"] == "KI-0001"
    assert row["date"] == "2024-03-01"
    assert row["base_amount"] == pytest.approx(1200.00)
    assert row["days"] == 425
    assert row["rate"] is None
    assert row["indicative_amount"] == pytest.approx(round(1200.00 * 425 / 788, 2), abs=0.01)
    assert 0 < row["indicative_amount"] < row["base_amount"]
    assert "Electricity Act 1989 s.84B" in row["legal_basis"]
    assert "SLC 7A" in row["legal_basis"]
    assert row["disclaimer"] == DISCLAIMER


# ---------- forced-credit-balance (hold) interest ----------


def test_credit_hold_interest_beyond_hold_window() -> None:
    """A negative Amount held 120 days (30 beyond the 90-day hold window)
    earns hold interest at the default 2% annual rate on those 30 days.

    The same credit is also never refunded, so a late-credit row (120 days)
    is emitted alongside it.
    """
    df = pd.DataFrame(
        [
            _record(
                "KCR-0001",
                "01 Feb 2026",
                -100.00,
                "01 Jan 2026",
                "31 Jan 2026",
            )
        ]
    )
    rows = estimate_compensation(df, config={"as_of": AS_OF})
    hold_rows = [r for r in rows if r["category"] == "credit_hold_interest"]
    assert len(hold_rows) == 1
    row = hold_rows[0]
    assert set(row.keys()) == ROW_KEYS
    assert row["invoice_ref"] == "KCR-0001"
    assert row["date"] == "2026-02-01"
    assert row["base_amount"] == pytest.approx(100.00)
    assert row["days"] == 30  # 120 elapsed - 90 hold window
    assert row["rate"] == pytest.approx(0.02)
    assert row["indicative_amount"] == pytest.approx(round(0.02 * 100.00 * 30 / 365, 2), abs=0.01)
    assert "SLC 21BA" in row["legal_basis"]
    assert row["disclaimer"] == DISCLAIMER

    late_rows = [r for r in rows if r["category"] == "late_credit_interest"]
    assert len(late_rows) == 1
    assert late_rows[0]["days"] == 120
    assert late_rows[0]["indicative_amount"] == pytest.approx(
        round(0.02 * 100.00 * 120 / 365, 2), abs=0.01
    )


# ---------- late-credit interest ----------


def test_late_credit_interest_credit_never_refunded() -> None:
    """A credit held 30 days (inside the 90-day hold window) earns NO hold
    interest, but because it was never refunded it earns late-credit
    interest for the full 30 days from Date to as_of.
    """
    df = pd.DataFrame(
        [
            _record(
                "KCR-0002",
                "02 May 2026",
                -100.00,
                "01 Apr 2026",
                "30 Apr 2026",
            )
        ]
    )
    rows = estimate_compensation(df, config={"as_of": AS_OF})
    assert [r for r in rows if r["category"] == "credit_hold_interest"] == []
    late_rows = [r for r in rows if r["category"] == "late_credit_interest"]
    assert len(late_rows) == 1
    row = late_rows[0]
    assert row["invoice_ref"] == "KCR-0002"
    assert row["days"] == 30
    assert row["rate"] == pytest.approx(0.02)
    assert row["indicative_amount"] == pytest.approx(round(0.02 * 100.00 * 30 / 365, 2), abs=0.01)
    assert row["disclaimer"] == DISCLAIMER


def test_refunded_credit_yields_no_interest_rows() -> None:
    """A credit refunded (later positive record within £0.50) inside the
    hold window produces neither hold interest nor late-credit interest.
    """
    df = pd.DataFrame(
        [
            _record(
                "KCR-0003",
                "01 Jan 2026",
                -100.00,
                "01 Dec 2025",
                "31 Dec 2025",
            ),
            _record(
                "PAY-0003",
                "10 Jan 2026",
                100.00,
                "01 Dec 2025",
                "31 Dec 2025",
            ),
        ]
    )
    rows = estimate_compensation(df, config={"as_of": AS_OF})
    assert rows == []


def test_credit_interest_config_overrides() -> None:
    """credit_hold_days=0 and credit_interest_rate=0.05 are honoured."""
    df = pd.DataFrame(
        [
            _record(
                "KCR-0004",
                "02 May 2026",
                -100.00,
                "01 Apr 2026",
                "30 Apr 2026",
            )
        ]
    )
    rows = estimate_compensation(
        df, config={"as_of": AS_OF, "credit_hold_days": 0, "credit_interest_rate": 0.05}
    )
    hold_rows = [r for r in rows if r["category"] == "credit_hold_interest"]
    assert len(hold_rows) == 1
    assert hold_rows[0]["days"] == 30
    assert hold_rows[0]["rate"] == pytest.approx(0.05)
    assert hold_rows[0]["indicative_amount"] == pytest.approx(
        round(0.05 * 100.00 * 30 / 365, 2), abs=0.01
    )


# ---------- no-op inputs ----------


def test_no_excess_no_credit_returns_empty() -> None:
    """A promptly-billed positive invoice and a zero-balance record yield
    no rows (no back-billing excess, no credit balance)."""
    df = pd.DataFrame(
        [
            _record(
                "KI-0002",
                "01 Jun 2024",
                100.00,
                "01 May 2024",
                "31 May 2024",
                period_charge=100.00,
            ),
            _record(
                "KI-0003",
                "01 Jul 2024",
                0.00,
                "01 Jun 2024",
                "30 Jun 2024",
                period_charge=0.00,
            ),
        ]
    )
    assert estimate_compensation(df, config={"as_of": AS_OF}) == []


def test_empty_df_returns_empty_list() -> None:
    """An empty dataframe returns an empty list."""
    assert estimate_compensation(pd.DataFrame(), config={"as_of": AS_OF}) == []


def test_none_df_returns_empty_list() -> None:
    """A None dataframe returns an empty list."""
    assert estimate_compensation(None, config={"as_of": AS_OF}) == []  # type: ignore[arg-type]


# ---------- disclaimer ----------


def test_disclaimer_constant_non_empty() -> None:
    """The module DISCLAIMER constant is a non-empty string."""
    assert isinstance(DISCLAIMER, str)
    assert DISCLAIMER.strip() != ""


def test_disclaimer_present_on_every_row() -> None:
    """Every emitted row carries the DISCLAIMER value verbatim."""
    df = pd.DataFrame(
        [
            _record(
                "KI-0004",
                "01 Mar 2024",
                1200.00,
                "01 Jan 2022",
                "28 Feb 2024",
                period_charge=1200.00,
            ),
            _record(
                "KCR-0005",
                "01 Feb 2026",
                -50.00,
                "01 Jan 2026",
                "31 Jan 2026",
            ),
        ]
    )
    rows = estimate_compensation(df, config={"as_of": AS_OF})
    assert len(rows) >= 3  # 1 excess + hold + late
    assert all(r["disclaimer"] == DISCLAIMER for r in rows)


# ---------- fix round 1: NaN and same-day refund handling ----------


def test_nan_amount_record_does_not_suppress_credit_rows() -> None:
    """A later record with a NaN Amount must not phantom-match as a refund.

    Regression pin for the review finding: NaN compares False to both
    ``<= 0`` and ``> 0.50`` checks, so without an explicit NaN guard the
    first later NaN record "refunded" any credit and silently dropped
    both the hold-interest and late-credit rows.
    """
    df = pd.DataFrame(
        [
            _record(
                "KCR-0006",
                "01 Feb 2026",
                -100.00,
                "01 Jan 2026",
                "31 Jan 2026",
            ),
            _record(
                "KI-0006",
                "15 Feb 2026",
                float("nan"),
                "01 Jan 2026",
                "31 Jan 2026",
            ),
        ]
    )
    rows = estimate_compensation(df, config={"as_of": AS_OF})
    assert len(rows) == 2
    hold_rows = [r for r in rows if r["category"] == "credit_hold_interest"]
    late_rows = [r for r in rows if r["category"] == "late_credit_interest"]
    assert len(hold_rows) == 1
    assert hold_rows[0]["days"] == 30
    assert len(late_rows) == 1
    assert late_rows[0]["days"] == 120
    assert {r["invoice_ref"] for r in rows} == {"KCR-0006"}


def test_nan_period_charge_emits_no_row() -> None:
    """A back-billing row whose Period Charge is NaN must not emit a row
    with NaN base_amount/indicative_amount — it is skipped like a zero
    charge (``NaN <= 0`` is False, so without a guard the NaN row leaked)."""
    df = pd.DataFrame(
        [
            _record(
                "KI-0007",
                "01 Mar 2024",
                1200.00,
                "01 Jan 2022",
                "28 Feb 2024",
                period_charge=float("nan"),
            )
        ]
    )
    rows = estimate_compensation(df, config={"as_of": AS_OF})
    assert rows == []


def test_same_day_refund_after_credit_suppresses_interest_rows() -> None:
    """A refund issued on the same statement date as the credit (listed
    after the credit row) counts as a refund: no hold row (0 days beyond
    the 90-day window) and no late-credit row."""
    df = pd.DataFrame(
        [
            _record(
                "KCR-0008",
                "10 Jan 2026",
                -100.00,
                "01 Dec 2025",
                "31 Dec 2025",
            ),
            _record(
                "PAY-0008",
                "10 Jan 2026",
                100.00,
                "01 Dec 2025",
                "31 Dec 2025",
            ),
        ]
    )
    rows = estimate_compensation(df, config={"as_of": AS_OF})
    assert rows == []


def test_same_day_refund_before_credit_is_documented_limitation() -> None:
    """Pins the documented limitation: a same-date refund listed BEFORE
    the credit row cannot be ordered by the data, so the credit is still
    treated as unrefunded and the late-credit row fires."""
    df = pd.DataFrame(
        [
            _record(
                "PAY-0009",
                "10 Jan 2026",
                100.00,
                "01 Dec 2025",
                "31 Dec 2025",
            ),
            _record(
                "KCR-0009",
                "10 Jan 2026",
                -100.00,
                "01 Dec 2025",
                "31 Dec 2025",
            ),
        ]
    )
    rows = estimate_compensation(df, config={"as_of": AS_OF})
    late_rows = [r for r in rows if r["category"] == "late_credit_interest"]
    assert len(late_rows) == 1
    assert late_rows[0]["invoice_ref"] == "KCR-0009"
