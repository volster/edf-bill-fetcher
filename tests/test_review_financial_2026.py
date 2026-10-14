"""Independent financial regressions from the October 2026 code review."""

from __future__ import annotations

import pandas as pd
import pytest

from edf_bill_fetcher.models.events import SapBackBillingEvent
from edf_bill_fetcher.models.report_models import compute_payment_analysis
from edf_bill_fetcher.processors.analysis import compute_dispute_flags
from edf_bill_fetcher.processors.compensation import estimate_compensation
from edf_bill_fetcher.processors.detection import (
    _parse_sub_periods,
    _reversal_match,
    compute_transitive_domination,
    compute_unlawful_union_total,
    detect_back_billing,
    detect_rebilling,
)
from edf_bill_fetcher.processors.matching import match_sap_events_to_edf


def bill(invoice: str, date: str, **changes: object) -> dict:
    """Separate the source account balance from the posted bill charge."""
    return {
        "Invoice #": invoice,
        "Date": date,
        "Entry Type": "New Bill",
        "Account": "A-001",
        "Fuel": "electricity",
        "Period From": "2023-01-01",
        "Period To": "2023-03-31",
        "Amount (£)": 900.0,
        "Period Charge (£)": 250.0,
        "Cancel/Rebill Admitted": False,
        **changes,
    }


def credit(**changes: object) -> dict:
    return {
        "Entry Type": "Credit",
        "Account": "A-001",
        "Fuel": "electricity",
        "Date": "2023-05-01",
        "Period From": "2023-01-01",
        "Period To": "2023-03-31",
        "Transaction Amount (£)": -250.0,
        **changes,
    }


def reversal_matches(record: dict, amount: float | None = 250.0) -> bool:
    return _reversal_match(
        pd.DataFrame([record]),
        "OLD",
        amount,
        pd.Timestamp("2023-01-01"),
        pd.Timestamp("2023-03-31"),
    )


def test_same_day_bills_do_not_establish_revision_order() -> None:
    rows = [
        bill("OLD", "2023-04-01"),
        bill("NEW", "2023-04-01", **{"Cancel/Rebill Admitted": True}),
    ]
    assert detect_rebilling(pd.DataFrame(rows)).empty


@pytest.mark.parametrize("admission", [float("nan"), None, "False"])
def test_missing_or_false_admission_cannot_create_a_revision(admission: object) -> None:
    rows = [
        bill("OLD", "2023-04-01"),
        bill("NEW", "2023-05-01", **{"Cancel/Rebill Admitted": admission}),
    ]
    assert detect_rebilling(pd.DataFrame(rows)).empty


def test_repeat_observation_of_same_invoice_cannot_supersede_itself() -> None:
    rows = [
        bill("SAME", "2023-04-01"),
        bill("SAME", "2023-05-01", **{"Cancel/Rebill Admitted": True}),
    ]
    assert detect_rebilling(pd.DataFrame(rows)).empty


def test_payment_does_not_supply_evidence_that_a_charge_was_reversed() -> None:
    assert reversal_matches(credit(**{"Entry Type": "Payment", "Amount (£)": 250.0})) is False


def test_reversal_uses_credit_transaction_even_when_account_balance_differs() -> None:
    assert reversal_matches(credit(**{"Amount Kind": "Balance", "Amount (£)": 900.0})) is True


def test_equal_account_balance_cannot_substitute_for_credit_transaction() -> None:
    assert (
        reversal_matches(
            credit(
                **{"Amount Kind": "Balance", "Amount (£)": 250.0, "Transaction Amount (£)": -20.0}
            )
        )
        is False
    )


@pytest.mark.parametrize("amount", [None, float("nan"), float("inf"), 0.0])
def test_unavailable_or_zero_bill_charge_cannot_match_credit(amount: float | None) -> None:
    assert reversal_matches(credit(), amount) is False


def test_reversal_matches_posted_bill_charge_instead_of_its_account_balance() -> None:
    rows = [bill("OLD", "2023-04-01"), bill("NEW", "2023-05-01")]
    result = detect_rebilling(pd.DataFrame(rows), evidence_df=pd.DataFrame([credit()]))
    assert list(zip(result["Killer Invoice"], result["Killed Invoice"], strict=True)) == [
        ("NEW", "OLD")
    ]


def test_other_account_credit_cannot_confirm_single_ledger_revision() -> None:
    rows = [bill("OLD", "2023-04-01", **{"Amount (£)": 250.0}), bill("NEW", "2023-05-01")]
    assert detect_rebilling(
        pd.DataFrame(rows),
        evidence_df=pd.DataFrame([credit(**{"Account": "A-002", "Amount (£)": 250.0})]),
    ).empty


@pytest.mark.parametrize("kind", ["Credit", "Payment", "Refund"])
def test_non_bill_transaction_cannot_be_the_replacement_invoice(kind: str) -> None:
    rows = [
        bill("OLD", "2023-04-01"),
        bill("TRANSACTION", "2023-05-01", **{"Entry Type": kind, "Cancel/Rebill Admitted": True}),
    ]
    assert detect_rebilling(pd.DataFrame(rows)).empty


def test_domination_reaches_latest_revision_through_non_back_billed_invoices() -> None:
    revisions = pd.DataFrame(
        [
            {"Killer Invoice": "SECOND", "Killed Invoice": "OLD"},
            {"Killer Invoice": "FINAL", "Killed Invoice": "SECOND"},
        ]
    )
    old = pd.DataFrame(
        [{"Invoice #": "OLD", "Period From": "2020-01-01", "Period To": "2020-02-01"}]
    )
    assert compute_transitive_domination(revisions, old) == {"OLD": ("FINAL", False)}


@pytest.mark.parametrize(
    ("field", "other"), [("Account", "A-002"), ("Fuel", "gas"), ("Meter", "M2"), ("Register", "R2")]
)
def test_sap_matching_rejects_conflicting_known_ledger_identity(field: str, other: str) -> None:
    identity = {"Account": "A-001", "Fuel": "electricity", "Meter": "M1", "Register": "R1"}
    event = SapBackBillingEvent(
        clearing_doc="C1",
        clearing_date=pd.Timestamp("2023-03-31"),
        clearing_reason="Statistical Item Reset",
        net_amount=250.0,
        rows=[{**identity, field: other, "Amount": 250.0, "Posting Date": "2023-03-31"}],
    )
    assert match_sap_events_to_edf([event], [bill("OLD", "2023-04-01", **identity)]) == []


def test_payment_intervals_are_measured_within_independent_ledgers() -> None:
    records = [
        {"Entry Type": "Payment", "Account": "A-001", "Date": "2024-01-01", "Amount (£)": 20.0},
        {"Entry Type": "Payment", "Account": "A-002", "Date": "2024-01-02", "Amount (£)": 100.0},
        {"Entry Type": "Payment", "Account": "A-001", "Date": "2024-02-01", "Amount (£)": 50.0},
    ]
    result = compute_payment_analysis(pd.DataFrame(records))
    assert result.count == 3
    assert result.total_paid == 170.0
    assert result.largest_payment == 100.0
    assert result.smallest_payment == 20.0
    assert result.avg_interval_days == 31.0
    assert result.median_interval_days == 31.0


def test_undated_payment_is_counted_but_cannot_claim_latest_date() -> None:
    records = [
        {"Entry Type": "Payment", "Date": "2024-02-01", "Amount (£)": 20.0},
        {"Entry Type": "Payment", "Date": "N/A", "Amount (£)": 100.0},
    ]
    result = compute_payment_analysis(pd.DataFrame(records))
    assert result.count == 2
    assert result.total_paid == 120.0
    assert result.last_payment_date == "2024-02-01"
    assert result.last_payment_amount == 20.0


def test_all_undated_payments_have_unavailable_latest_payment() -> None:
    records = [{"Entry Type": "Payment", "Date": "N/A", "Amount (£)": 100.0}]
    result = compute_payment_analysis(pd.DataFrame(records))
    assert result.count == 1
    assert result.last_payment_date is None
    assert result.last_payment_amount is None


@pytest.mark.parametrize("value", ["NaN", "inf", "-inf"])
def test_nonfinite_subperiod_units_cannot_poison_indicative_exposure(value: str) -> None:
    raw = f"2023-01-01|2023-03-31|{value}|25.0|250.0"
    assert _parse_sub_periods(raw) == []
    result = detect_back_billing(
        pd.DataFrame([bill("LATE", "2024-05-01", **{"Sub Periods": raw, "Units (kWh)": value})])
    )
    assert result.iloc[0]["Sub-Period Basis"] == "Day-ratio fallback"
    assert compute_unlawful_union_total(result) == 250.0


@pytest.mark.parametrize("kind", ["Credit", "Payment", "Refund"])
def test_non_charge_transactions_do_not_create_back_billing_exposure(kind: str) -> None:
    record = bill("TRANSFER", "2024-05-01", **{"Entry Type": kind})
    assert detect_back_billing(pd.DataFrame([record])).empty


def test_sap_match_narrative_identifies_the_posting_date_it_actually_used() -> None:
    event = SapBackBillingEvent(
        clearing_doc="C1",
        clearing_date=pd.Timestamp("2023-12-01"),
        clearing_reason="Statistical Item Reset",
        net_amount=250.0,
        rows=[{"Amount": 250.0, "Posting Date": "2023-03-01"}],
    )
    result = match_sap_events_to_edf([event], [bill("OLD", "2023-04-01")])
    assert len(result) == 1
    assert result[0].confidence_band == "High"
    assert "Posting date inside EDF period" in result[0].notes


def test_payment_cannot_reduce_indicative_back_billing_exposure_by_replacing_a_bill() -> None:
    records = [
        bill("OLD", "2024-05-01", **{"Amount (£)": 250.0}),
        bill("NEW", "2024-06-01", **{"Period Charge (£)": 50.0}),
        credit(**{"Entry Type": "Payment", "Date": "2024-06-01", "Amount (£)": 250.0}),
    ]
    result = estimate_compensation(pd.DataFrame(records), {"as_of": "2024-07-01"})
    assert sum(row["indicative_amount"] for row in result) == 250.0
    assert result[0]["invoice_ref"] == "OLD"


def reused_invoice_records() -> list[dict]:
    return [
        bill(
            "A",
            "2024-01-01",
            **{
                "Period From": "2020-01-01",
                "Period To": "2020-03-31",
                "Cancel/Rebill Admitted": True,
            },
        ),
        bill(
            "B",
            "2024-02-01",
            **{
                "Period From": "2020-01-01",
                "Period To": "2020-03-31",
                "Cancel/Rebill Admitted": True,
            },
        ),
        bill(
            "A",
            "2024-03-01",
            **{
                "Period From": "2020-01-01",
                "Period To": "2020-03-31",
                "Period Charge (£)": 300.0,
                "Cancel/Rebill Admitted": True,
            },
        ),
    ]


def test_reused_invoice_reference_cannot_create_a_revision_cycle_or_hide_exposure() -> None:
    records = pd.DataFrame(reused_invoice_records())
    assert detect_rebilling(records, evidence_df=records).empty
    result = estimate_compensation(records, {"as_of": "2024-04-01"})
    assert sum(row["indicative_amount"] for row in result) == 250.0
    assert result[0]["invoice_ref"] == "A"


def test_conflicting_invoice_observations_are_explicitly_flagged_for_review() -> None:
    flags, _ = compute_dispute_flags(pd.DataFrame(reused_invoice_records()))
    conflicts = [flag for flag in flags if flag[0] == "INVOICE IDENTITY CONFLICT"]
    assert len(conflicts) == 1
    assert "A" in conflicts[0][3]
    assert "Automatic supersession is withheld" in conflicts[0][3]


def test_identical_observations_are_not_invoice_identity_conflicts() -> None:
    old = bill("A", "2024-01-01")
    flags, _ = compute_dispute_flags(pd.DataFrame([old, dict(old)]))
    assert not any(flag[0] == "INVOICE IDENTITY CONFLICT" for flag in flags)


def test_domination_cannot_hide_evidence_through_a_cyclic_revision_chain() -> None:
    revisions = pd.DataFrame(
        [
            {"Killer Invoice": "A", "Killed Invoice": "B"},
            {"Killer Invoice": "B", "Killed Invoice": "A"},
            {"Killer Invoice": "B", "Killed Invoice": "OLD"},
        ]
    )
    billed = pd.DataFrame([{"Invoice #": invoice} for invoice in ("A", "B", "OLD")])
    assert compute_transitive_domination(revisions, billed) == {}


def test_credit_predating_the_bill_cannot_establish_its_reversal() -> None:
    rows = [bill("OLD", "2023-04-01"), bill("NEW", "2023-05-01")]
    prior = credit(**{"Date": "2023-02-01"})
    assert detect_rebilling(pd.DataFrame(rows), evidence_df=pd.DataFrame([prior])).empty


def test_unknown_credit_date_retains_the_existing_review_heuristic() -> None:
    rows = [bill("OLD", "2023-04-01"), bill("NEW", "2023-05-01")]
    result = detect_rebilling(
        pd.DataFrame(rows), evidence_df=pd.DataFrame([credit(**{"Date": "N/A"})])
    )
    assert list(zip(result["Killer Invoice"], result["Killed Invoice"], strict=True)) == [
        ("NEW", "OLD")
    ]


def test_same_invoice_reference_in_independent_accounts_is_not_a_conflict() -> None:
    records = [bill("A", "2024-01-01"), bill("A", "2024-02-01", **{"Account": "A-002"})]
    flags, _ = compute_dispute_flags(pd.DataFrame(records))
    assert not any(flag[0] == "INVOICE IDENTITY CONFLICT" for flag in flags)


def test_sap_fuel_alias_cannot_bypass_ledger_matching_boundary() -> None:
    event = SapBackBillingEvent(
        clearing_doc="C1",
        clearing_date=pd.Timestamp("2023-03-31"),
        clearing_reason="Statistical Item Reset",
        net_amount=250.0,
        rows=[{"Amount": 250.0, "Posting Date": "2023-03-31", "Fuel Type": "gas"}],
    )
    assert match_sap_events_to_edf([event], [bill("OLD", "2023-04-01")]) == []


def test_sap_unknown_account_placeholder_does_not_hide_known_account_alias() -> None:
    event = SapBackBillingEvent(
        clearing_doc="C1",
        clearing_date=pd.Timestamp("2023-03-31"),
        clearing_reason="Statistical Item Reset",
        net_amount=250.0,
        rows=[
            {"Amount": 250.0, "Posting Date": "2023-03-31", "Account": "N/A", "Kraken ID": "A-001"}
        ],
    )
    assert len(match_sap_events_to_edf([event], [bill("OLD", "2023-04-01")])) == 1
