"""Invalid draft commands are atomic; clearing claims and relationships is reversible."""

import copy

import pytest

from edf_bill_fetcher.application.record_review import apply_review, review_inputs
from edf_bill_fetcher.application.review_session import ReviewSession
from tests.synthetic_analysis import snapshot
from tests.test_claim_reconciliation import evidence
from tests.test_rebilling_ledger import set_position
from tests.test_record_review_validation import context


@pytest.fixture
def session():
    records = [item["record"] for item in evidence()[0]]
    return ReviewSession(snapshot(records))


def key(session, invoice):
    return next(item["id"] for item in session.rows if item["record"]["Invoice #"] == invoice)


def fields(members, **changes):
    return {
        "title": "Consolidated challenge",
        "status": "draft",
        "members": list(members),
        "challenged": "120.00",
        "reduction": "120.00",
        "credits": [],
        "reason": "Verified original evidence",
        "account_link_reason": "",
        **changes,
    }


@pytest.mark.parametrize(
    "command",
    ["decision", "position", "clear_position", "clear_claim", "claim_fields", "edit_fields"],
)
def test_invalid_command_preserves_draft_and_undo_history(session, command):
    before = session.payload()
    with pytest.raises(ValueError):
        if command == "decision":
            session.decision("missing")
        elif command == "position":
            session.set_position("missing")
        elif command == "clear_position":
            session.clear_position("missing")
        elif command == "clear_claim":
            session.clear_claim("missing")
        elif command == "claim_fields":
            session.set_claim("C1", title="Incomplete")
        else:
            session.update(
                key(session, "OLD"), {"Source ID": "forged"}, True, "Checked", "verified"
            )
    assert session.payload() == before
    assert session.undo_stack == []


def test_existing_malformed_date_can_be_repaired_and_reverted_without_mutating_source():
    records = [item["record"] for item in evidence()[0]]
    records[0]["Date"] = "Unreadable date"
    with pytest.warns(UserWarning, match="unparseable dates"):
        original = snapshot(records)
    session = ReviewSession(original)
    row_id = key(session, "OLD")
    session.update(row_id, {"Date": "01/02/2024"}, True, "Read the original invoice", "verified")
    assert session.record(row_id)["Date"] == "01/02/2024"
    assert session.index[row_id]["record"]["Date"] == "Unreadable date"
    session.undo()
    assert session.record(row_id)["Date"] == "Unreadable date"
    session.redo()
    assert session.record(row_id)["Date"] == "01/02/2024"
    assert (
        original.prepared.extraction.observations[0].raw
        == session.snapshot.prepared.extraction.observations[0].raw
    )


def test_reverting_a_field_removes_its_correction_and_noop_history_is_not_added(session):
    row_id = key(session, "OLD")
    original = session.record(row_id)["Period Charge (£)"]
    session.update(row_id, {"Period Charge (£)": 125}, True, "Read source", "unreviewed")
    session.update(row_id, {"Period Charge (£)": original}, True, "Restored original", "unreviewed")
    assert session.decision(row_id)["changes"] == {}
    count = len(session.undo_stack)
    session.update(row_id, {"Period Charge (£)": original}, True, "Restored original", "unreviewed")
    assert len(session.undo_stack) == count
    assert session.record(row_id)["Transaction Amount (£)"] == original


def test_clearing_revision_and_dispute_position_remains_undoable(session):
    older, newer = key(session, "OLD"), key(session, "REPLACEMENT")
    session.set_revision(older, newer, "replacement", "confirmed", "Compared originals")
    session.set_revision(older, newer, "replacement", "suggested", "")
    assert session.revisions == {}
    session.undo()
    assert session.revisions[older, newer]["status"] == "confirmed"
    session.redo()
    assert session.revisions == {}
    set_position(session, older)
    session.clear_position(older)
    assert not session.schedule
    session.undo()
    assert older in session.schedule


def test_unreviewable_group_rejection_keeps_previous_draft(session):
    row_id = key(session, "OLD")
    before = session.payload()
    with pytest.raises(ValueError, match="Cannot review this group"):
        session.set_claim("C1", **fields([row_id], status="reviewed"))
    assert session.payload() == before
    assert not session.undo_stack


@pytest.mark.parametrize("pair", [("C1", "C1"), ("C1", "missing")])
def test_invalid_claim_pair_cannot_create_a_separation(session, pair):
    session.set_claim("C1", **fields([key(session, "OLD")]))
    before = session.payload()
    with pytest.raises(ValueError, match="two different"):
        session.separate_claims(*pair, "Checked originals")
    assert session.payload() == before


def test_shared_document_cannot_be_declared_two_independent_charges(session):
    for name in ("C1", "C2"):
        session.set_claim(name, **fields([key(session, "OLD")]))
    before = session.payload()
    with pytest.raises(ValueError, match="share a bill"):
        session.separate_claims("C1", "C2", "Attempt to waive the duplicate")
    assert session.payload() == before


def test_separation_clear_and_group_removal_restore_all_dependent_decisions(session):
    session.set_claim("C1", **fields([key(session, "OLD")]))
    session.set_claim("C2", **fields([key(session, "REPLACEMENT")]))
    session.set_claim("C3", **fields([key(session, "NEXT")]))
    session.separate_claims("C2", "C1", "Verified distinct charge lines")
    session.separate_claims("C2", "C3", "Independent periods")
    before = session.payload()
    session.clear_separation("C2", "C1")
    assert ("C1", "C2") not in session.claim_separations
    session.undo()
    assert session.payload() == before
    session.clear_claim("C1")
    assert set(session.claims) == {"C2", "C3"}
    assert set(session.claim_separations) == {("C2", "C3")}
    session.undo()
    assert session.payload() == before
    session.redo()
    assert set(session.claims) == {"C2", "C3"}


@pytest.mark.parametrize("decisions", [{}, [1], [None]])
def test_invalid_draft_decision_collection_is_rejected(session, decisions):
    payload = session.payload()
    payload["decisions"] = decisions
    with pytest.raises(ValueError, match="Invalid draft decisions"):
        ReviewSession.restore(session.snapshot, payload)
    assert not session.decisions


def test_optional_relationship_and_schedule_inputs_preserve_raw_evidence(session, tmp_path):
    originals, state, _ = review_inputs(
        session.snapshot,
        [session.decision(item["id"]) for item in session.rows],
        input_path=str(tmp_path / "review.json"),
    )
    assert len(originals) == len(session.snapshot.prepared.extraction.observations)
    assert "bill_revision_decisions" not in state
    assert "dispute_schedule" not in state
    assert state["review_history"] == []


def test_undo_and_redo_on_a_fresh_draft_are_noops(session):
    before = session.payload()
    session.undo()
    session.redo()
    assert session.payload() == before
    assert not session.undo_stack
    assert not session.redo_stack


@pytest.mark.parametrize(
    "decision",
    [None, {}, {"row_id": "b" * 64, "include": True, "changes": {}, "reason": "", "extra": True}],
)
def test_malformed_complete_review_decision_is_rejected_without_changing_rows(decision):
    value = context()
    value["decisions"] = [decision]
    before = copy.deepcopy(value)
    with pytest.raises(ValueError, match="Invalid review decision"):
        apply_review(value)
    assert value == before
