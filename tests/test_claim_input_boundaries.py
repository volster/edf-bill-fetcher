"""Human proposals reject corrupt allocations and keep independent charges separate."""

import copy

import pytest

from edf_bill_fetcher.processors.claim_reconciliation import (
    build_reconciliation,
    empty_state,
    validate_state,
)
from edf_bill_fetcher.processors.rebilling_ledger import build_ledger, schedule_digest
from edf_bill_fetcher.processors.revisions import build_revisions
from tests.test_claim_reconciliation import credit, evidence, group, make_ledger, reconcile
from tests.test_rebilling_ledger import position


@pytest.fixture
def case():
    items, ledger = evidence()
    state = empty_state()
    state["groups"] = [group(items, ledger, credits=[credit("10")])]
    return items, ledger, state


def replace(state, path, value):
    for field in path[:-1]:
        state = state[field]
    state[path[-1]] = value


@pytest.mark.parametrize(
    ("path", "value", "message"),
    [
        (("version",), True, "Unsupported"),
        (("groups",), {}, "Too many"),
        (("separations",), {}, "Too many"),
        (("context",), {}, "complaint context"),
        (("context", "include_pack"), 1, "complaint context"),
        (("groups", 0), {}, "group fields"),
        (("groups", 0, "id"), "bad id", "Claim IDs"),
        (("groups", 0, "source_digest"), "bad", "evidence digest"),
        (("groups", 0, "credits"), {}, "credit allocations"),
        (("groups", 0, "credits", 0), {}, "allocation fields"),
        (("groups", 0, "credits", 0, "row_id"), "missing", "known source"),
        (("groups", 0, "credits", 0, "treatment"), [], "unique treatment"),
        (("groups", 0, "credits", 0, "amount"), None, "positive amounts"),
        (("groups", 0, "credits", 0, "amount"), "", "positive amounts"),
        (("groups", 0, "credits", 0, "amount"), 0, "positive amounts"),
        (("groups", 0, "members"), [], "requires bills"),
        (("groups", 0, "challenged"), None, "needs a challenged amount"),
        (("groups", 0, "reduction"), None, "requires bills"),
        (("separations",), [{}], "overlap review fields"),
        (
            ("separations",),
            [{"groups": ["C1", "missing"], "reason": "Checked", "source_digest": "a" * 64}],
            "two distinct",
        ),
    ],
)
def test_invalid_human_claim_input_is_atomic(case, path, value, message):
    items, _, state = case
    replace(state, path, value)
    before = copy.deepcopy((items, state))
    with pytest.raises(ValueError, match=message):
        validate_state(items, state)
    assert (items, state) == before


@pytest.mark.parametrize("value", [True, [], "NaN", "Infinity", "-1", "0.001", "1000000000001"])
def test_non_money_cannot_become_a_claim_amount(case, value):
    items, _, state = case
    state["groups"][0]["challenged"] = value
    with pytest.raises(ValueError, match="non-negative money"):
        validate_state(items, state)


@pytest.mark.parametrize(
    ("path", "limit"),
    [(("groups",), 1000), (("separations",), 10000), (("groups", 0, "credits"), 10000)],
)
def test_claim_collection_limits_apply_before_processing_items(case, path, limit):
    items, _, state = case
    replace(state, path, [None] * (limit + 1))
    with pytest.raises(ValueError, match="Too many|credit allocations"):
        validate_state(items, state)


def test_duplicate_credit_treatment_and_case_insensitive_claim_id_are_rejected(case):
    items, _, state = case
    state["groups"][0]["credits"] *= 2
    with pytest.raises(ValueError, match="unique treatment"):
        validate_state(items, state)
    state["groups"][0]["credits"] = []
    state["groups"].append({**state["groups"][0], "id": "c1"})
    with pytest.raises(ValueError, match="Claim IDs must be unique"):
        validate_state(items, state)


def test_separation_requires_distinct_ordered_known_groups(case):
    items, ledger, state = case
    state["groups"].append(group(items, ledger, "C2", ("b",)))
    decision = {
        "groups": ["C1", "C2"],
        "reason": "Checked distinct lines",
        "source_digest": "a" * 64,
    }
    for pair in (["C1", "C1"], ["C2", "C1"], ["C1"], [1, "C2"]):
        state["separations"] = [{**decision, "groups": pair}]
        with pytest.raises(ValueError, match="two distinct"):
            validate_state(items, state)
    state["separations"] = [decision, decision]
    with pytest.raises(ValueError, match="two distinct"):
        validate_state(items, state)


@pytest.mark.parametrize("missing", ["", None])
def test_incomplete_draft_remains_unready_without_inventing_zero(case, missing):
    items, ledger, state = case
    state["groups"][0].update(status="draft", challenged=missing, reduction=missing)
    model = build_reconciliation(items, ledger, state)
    assert not model["ready"]
    assert model["proposed_reduction"] is None
    assert model["input"]["groups"][0]["challenged"] is None


def test_independent_periods_total_once_and_ignore_incompatible_or_undated_balances():
    items, _ = evidence()
    base = items[0]["record"]
    items.extend(
        [
            {
                "id": "gas",
                "included": True,
                "record": {**base, "Fuel": "Gas", "Date": "01/03/2025", "Balance (£)": 9999},
            },
            {
                "id": "undated",
                "included": True,
                "record": {**base, "Date": "unreadable", "Balance (£)": 8888},
            },
            {
                "id": "excluded",
                "included": False,
                "record": {**base, "Date": "01/03/2025", "Balance (£)": 7777},
            },
        ]
    )
    ledger = make_ledger(items, ("a", "d"))
    january = group(items, ledger, "JAN", ("a",), challenged="100", reduction="100")
    february = group(items, ledger, "FEB", ("d",), challenged="200", reduction="200")
    model = reconcile(items, ledger, [january, february])
    assert model["ready"]
    assert model["overlaps"] == []
    assert model["proposed_reduction"] == "300.00"
    assert model["ledgers"][0]["observed_balance"] == "270.00"
    assert model["ledgers"][0]["balance_date"] == "2024-02-02"
    assert model["ledgers"][0]["illustrative_balance"] == "-30.00"


@pytest.mark.parametrize("period", ["unreadable", "01/03/2024"])
def test_unknown_or_reversed_period_keeps_overlap_unresolved(period):
    items, _ = evidence()
    items[0]["record"]["Period From"] = period
    ledger = make_ledger(items)
    model = reconcile(
        items, ledger, [group(items, ledger, "C1", ("a",)), group(items, ledger, "C2", ("b",))]
    )
    assert model["overlaps"][0]["kind"] == "unknown_periods"
    assert model["proposed_reduction"] is None


def test_draft_with_credit_but_no_bill_does_not_establish_a_separate_period():
    items, ledger = evidence()
    draft = group(
        items,
        ledger,
        "DRAFT",
        (),
        status="draft",
        challenged=None,
        reduction=None,
        credits=[credit("10")],
    )
    model = reconcile(items, ledger, [draft, group(items, ledger, "BILL", ("a", "b"))])
    assert model["overlaps"][0]["kind"] == "unknown_periods"
    assert model["proposed_reduction"] is None


def test_confirmed_revision_relationship_invalidates_an_earlier_claim_digest():
    items, ledger = evidence()
    claim = group(items, ledger)
    explicit = [
        {
            "older": "a",
            "newer": "b",
            "kind": "replacement",
            "status": "confirmed",
            "reason": "Compared the original invoices",
        }
    ]
    positions = [
        position(key, source_digest=schedule_digest(items, explicit, key)) for key in ("a", "b")
    ]
    revised = build_ledger(items, build_revisions(items, explicit), positions)
    model = reconcile(items, revised, [claim])
    assert model["proposed_reduction"] is None
    assert model["groups"][0]["needs_review"]
    assert any("evidence changed" in issue for issue in model["groups"][0]["issues"])


def test_independent_fuels_need_no_overlap_waiver_and_keep_separate_balances():
    items, _ = evidence()
    items[1]["record"].update(Fuel="Gas", Meter="GAS-METER")
    ledger = make_ledger(items)
    model = reconcile(
        items,
        ledger,
        [
            group(items, ledger, "ELEC", ("a",), challenged="100", reduction="100"),
            group(items, ledger, "GAS", ("b",)),
        ],
    )
    assert model["ready"]
    assert model["overlaps"] == []
    assert model["proposed_reduction"] == "220.00"
    assert len(model["ledgers"]) == 2
    assert {entry["observed_balance"] for entry in model["ledgers"]} == {"250.00", "270.00"}


def test_unknown_account_relationship_blocks_total_instead_of_guessing_a_link():
    items, _ = evidence()
    items[0]["record"]["Meter"] = ""
    items[1]["record"].update(Account="OTHER", Meter="")
    ledger = make_ledger(items)
    model = reconcile(
        items, ledger, [group(items, ledger, "C1", ("a",)), group(items, ledger, "C2", ("b",))]
    )
    assert model["proposed_reduction"] is None
    assert model["overlaps"][0]["kind"] == "unverified_account_relationship"


def test_conflicting_latest_balances_withhold_illustrative_balance_without_changing_proposal():
    items, _ = evidence()
    for key, amount in (("first", 300), ("second", 350)):
        items.append(
            {
                "id": key,
                "included": True,
                "record": {
                    **items[0]["record"],
                    "Date": "03/02/2024",
                    "Entry Type": "Ongoing Balance",
                    "Balance (£)": amount,
                },
            }
        )
    ledger = make_ledger(items)
    model = reconcile(items, ledger, [group(items, ledger)])
    assert model["ready"]
    assert model["proposed_reduction"] == "120.00"
    assert model["ledgers"][0]["observed_balance"] is None
    assert model["ledgers"][0]["balance_date"] is None
    assert model["ledgers"][0]["illustrative_balance"] is None
