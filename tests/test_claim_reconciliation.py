"""Independent economic examples for consolidated claims and already applied credits."""

import copy

import pytest

from edf_bill_fetcher.processors.claim_reconciliation import (
    build_reconciliation,
    empty_state,
    group_digest,
    separation_digest,
    validate_state,
)
from edf_bill_fetcher.processors.rebilling_ledger import build_ledger, schedule_digest
from edf_bill_fetcher.processors.revisions import build_revisions


def evidence(active=("a", "b")):
    base = {
        "Account": "A-EXAMPLE",
        "Fuel": "Electricity",
        "Meter": "EXAMPLE-METER",
        "Entry Type": "New Bill",
        "Period From": "01/01/2024",
        "Period To": "31/01/2024",
    }
    items = [
        {
            "id": "a",
            "included": True,
            "record": {
                **base,
                "Invoice #": "OLD",
                "Date": "01/02/2024",
                "Period Charge (£)": 100,
                "Transaction Amount (£)": 100,
                "Balance (£)": 250,
            },
        },
        {
            "id": "b",
            "included": True,
            "record": {
                **base,
                "Invoice #": "REPLACEMENT",
                "Date": "02/02/2024",
                "Period Charge (£)": 120,
                "Transaction Amount (£)": 120,
                "Balance (£)": 270,
            },
        },
        {
            "id": "c",
            "included": True,
            "record": {
                **base,
                "Invoice #": "CREDIT",
                "Date": "02/02/2024",
                "Entry Type": "Credit",
                "Transaction Amount (£)": -100,
                "Details": "Cancels OLD",
            },
        },
        {
            "id": "d",
            "included": True,
            "record": {
                **base,
                "Invoice #": "NEXT",
                "Date": "01/03/2024",
                "Period From": "01/02/2024",
                "Period To": "29/02/2024",
                "Period Charge (£)": 200,
                "Transaction Amount (£)": 200,
            },
        },
    ]
    return items, make_ledger(items, active)


def make_ledger(items, active=("a", "b")):
    positions = [
        {
            "row_id": key,
            "source_digest": schedule_digest(items, [], key),
            "status": "challenged",
            "challenged_amount": "200.00",
            "proposed_reduction": "200.00",
            "first_demand_date": None,
            "basis": "Full challenge pending the earlier billing history",
            "edf_request": "Earliest demand and allocation of extra charges",
        }
        for key in active
    ]
    return build_ledger(items, build_revisions(items), positions)


def group(items, ledger, key="C1", members=("a", "b"), **changes):
    g = {
        "id": key,
        "title": "Historic replacement charges",
        "status": "reviewed",
        "members": list(members),
        "challenged": "120.00",
        "reduction": "120.00",
        "credits": [],
        "reason": "These invoices demand the same charges; this is the consolidated challenge, not their sum.",
        "account_link_reason": "",
        "source_digest": "0" * 64,
        **changes,
    }
    state = empty_state()
    state["groups"] = [g]
    normalized = validate_state(items, state)["groups"][0]
    normalized["source_digest"] = group_digest(items, ledger, normalized)
    return normalized


def credit(amount, treatment="cancellation_offset", row_id="c"):
    return {
        "row_id": row_id,
        "amount": amount,
        "treatment": treatment,
        "reason": "Verified the original credit and its allocation against the source history",
    }


def reconcile(items, ledger, groups, separations=()):
    state = empty_state()
    state.update(groups=groups, separations=list(separations))
    return build_reconciliation(items, ledger, state)


def test_replacement_history_is_consolidated_once_without_deducting_cancellation_twice():
    items, ledger = evidence()
    original = copy.deepcopy(items)
    claim = group(items, ledger, credits=[credit("100")])
    model = reconcile(items, ledger, [claim])
    assert model["ready"]
    assert model["proposed_reduction"] == "120.00"
    assert model["groups"][0]["cancellation_offsets"] == "100.00"
    assert model["groups"][0]["relief_received"] == "0.00"
    assert model["ledgers"][0]["observed_balance"] == "270.00"
    assert model["ledgers"][0]["illustrative_balance"] == "150.00"
    assert items == original


def test_actual_relief_is_subtracted_but_cancellation_offsets_are_not():
    items, ledger = evidence()
    claim = group(items, ledger, credits=[credit("70"), credit("30", "already_applied_relief")])
    model = reconcile(items, ledger, [claim])
    assert model["ready"]
    assert model["proposed_reduction"] == "90.00"
    assert model["groups"][0]["relief_received"] == "30.00"
    assert model["groups"][0]["cancellation_offsets"] == "70.00"


def test_overlaps_block_a_total_until_separate_charges_are_explicitly_substantiated():
    items, ledger = evidence()
    a, b = group(items, ledger, members=("a",)), group(items, ledger, "C2", ("b",))
    blocked = reconcile(items, ledger, [a, b])
    assert not blocked["ready"]
    assert blocked["proposed_reduction"] is None
    assert blocked["overlaps"][0]["kind"] == "overlapping_periods"
    decision = {
        "groups": ["C1", "C2"],
        "reason": "Distinct specifically identified charge lines; their source allocations are separate",
        "source_digest": separation_digest(items, ledger, a, b),
    }
    resolved = reconcile(items, ledger, [a, b], [decision])
    assert resolved["ready"]
    assert resolved["proposed_reduction"] == "240.00"
    b["reduction"] = "100.00"
    b["source_digest"] = group_digest(items, ledger, b)
    stale = reconcile(items, ledger, [a, b], [decision])
    assert not stale["ready"]
    assert stale["overlaps"][0]["needs_review"]


def test_same_document_in_two_groups_cannot_be_waived():
    items, ledger = evidence(("a",))
    a, b = group(items, ledger, members=("a",)), group(items, ledger, "C2", ("a",))
    decision = {
        "groups": ["C1", "C2"],
        "reason": "Attempt to treat duplicates separately",
        "source_digest": separation_digest(items, ledger, a, b),
    }
    model = reconcile(items, ledger, [a, b], [decision])
    assert not model["ready"]
    assert model["overlaps"][0]["kind"] == "shared_documents"
    assert not model["overlaps"][0]["resolved"]


def test_distinct_periods_can_total_but_an_ungrouped_position_still_blocks_completion():
    items, ledger = evidence(("a", "d"))
    a, b = group(items, ledger, members=("a",)), group(items, ledger, "C2", ("d",))
    model = reconcile(items, ledger, [a, b])
    assert model["ready"]
    assert not model["overlaps"]
    assert model["proposed_reduction"] == "240.00"
    partial = reconcile(items, ledger, [a])
    assert partial["reviewed_subtotal"] == "120.00"
    assert partial["proposed_reduction"] is None
    assert partial["unresolved_positions"] == ["d"]


def test_a_credit_cannot_be_spent_twice_across_reviewed_groups():
    items, ledger = evidence(("a", "d"))
    a = group(items, ledger, members=("a",), credits=[credit("70", "already_applied_relief")])
    b = group(items, ledger, "C2", ("d",), credits=[credit("50", "already_applied_relief")])
    model = reconcile(items, ledger, [a, b])
    assert not model["ready"]
    assert model["proposed_reduction"] is None
    assert all(any("more than once" in issue for issue in g["issues"]) for g in model["groups"])


@pytest.mark.parametrize(
    "change",
    [
        "excluded",
        "amount",
        "draft",
        "conflicting_meter",
        "missing_account",
        "unknown_credit",
        "excess_credit",
        "excess_relief",
    ],
)
def test_invalid_or_changed_evidence_never_becomes_a_publishable_total(change):
    items, ledger = evidence()
    claim = group(items, ledger)
    if change == "excluded":
        items[1]["included"] = False
    elif change == "amount":
        items[0]["record"]["Period Charge (£)"] = 500
    elif change == "draft":
        claim["status"] = "draft"
    elif change == "conflicting_meter":
        items[1]["record"]["Meter"] = "OTHER"
    elif change == "missing_account":
        items[1]["record"]["Account"] = ""
    elif change == "unknown_credit":
        items[2]["record"]["Transaction Amount (£)"] = None
        claim["credits"] = [credit("10")]
    elif change == "excess_credit":
        claim["credits"] = [credit("101")]
    else:
        claim.update(reduction="20.00", credits=[credit("30", "already_applied_relief")])
    updated = make_ledger(items)
    if change != "amount":
        claim["source_digest"] = group_digest(items, updated, claim)
    model = reconcile(items, updated, [claim])
    assert not model["ready"]
    assert model["proposed_reduction"] is None


def test_linking_accounts_needs_a_shared_supply_and_explicit_evidence_reason():
    items, _ = evidence()
    items[1]["record"]["Account"] = "A-MIGRATED"
    ledger = make_ledger(items)
    claim = group(items, ledger)
    assert not reconcile(items, ledger, [claim])["ready"]
    claim = group(
        items,
        ledger,
        account_link_reason="Original account transfer document links the same physical meter and electricity supply",
    )
    model = reconcile(items, ledger, [claim])
    assert model["ready"]
    assert len(model["groups"][0]["scope"]["accounts"]) == 2


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("reduction", -1),
        ("challenged", True),
        ("reduction", "nan"),
        ("reduction", "1.001"),
        ("reduction", "121"),
        ("status", "legal entitlement"),
        ("members", ["missing"]),
        ("members", ["a", "a"]),
        ("reason", ""),
    ],
)
def test_bad_claim_inputs_are_rejected(field, value):
    items, ledger = evidence()
    with pytest.raises(ValueError):
        group(items, ledger, **{field: value})


def test_missing_or_conflicting_balance_observations_do_not_turn_into_zero():
    items, ledger = evidence()
    items[0]["record"]["Balance (£)"] = None
    items[1]["record"]["Balance (£)"] = None
    ledger = make_ledger(items)
    claim = group(items, ledger)
    entry = reconcile(items, ledger, [claim])["ledgers"][0]
    assert entry["observed_balance"] is None
    assert entry["illustrative_balance"] is None
