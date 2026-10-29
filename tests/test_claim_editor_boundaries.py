"""Real claim-editor actions retain pending edits and withhold unresolved totals."""

import copy

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QMessageBox

from tests.test_claim_ui import dialog as dialog
from tests.test_claim_ui import fill_group
from tests.test_claim_ui import session as session
from tests.test_evidence_review import application as application


def two_groups(dialog):
    fill_group(dialog, ("a",))
    assert dialog.save_group()
    first = dialog.group_id
    dialog.new_group()
    fill_group(dialog, ("b",))
    assert dialog.save_group()
    return first, dialog.group_id


def test_overlap_review_can_be_saved_cleared_and_group_removal_undone(dialog, session, application):
    original_summary = session.snapshot.summary
    first, second = two_groups(dialog)
    assert session.claim_model()["proposed_reduction"] is None
    assert dialog.overlaps.topLevelItemCount() == 1
    dialog.overlaps.setCurrentItem(dialog.overlaps.topLevelItem(0))
    application.processEvents()
    dialog.overlap_reason.setPlainText(
        "Distinct original charge lines verified against the two invoices"
    )
    assert dialog.save_overlap()
    assert session.claim_model()["proposed_reduction"] == "240.00"
    dialog.refresh()
    assert dialog.overlaps.currentItem() is not None
    dialog.clear_overlap()
    assert session.claim_model()["proposed_reduction"] is None
    assert not session.claim_separations
    dialog.remove_group()
    assert set(session.claims) == {first}
    session.undo()
    assert set(session.claims) == {first, second}
    assert session.snapshot.summary == original_summary


@pytest.mark.parametrize(
    "choice",
    [
        QMessageBox.StandardButton.Save,
        QMessageBox.StandardButton.Discard,
        QMessageBox.StandardButton.Cancel,
    ],
)
def test_switching_groups_preserves_or_applies_pending_edits(
    dialog, session, application, monkeypatch, choice
):
    first, second = two_groups(dialog)
    old_title = session.claims[second]["title"]
    dialog.title.setText("Pending revised scope")
    monkeypatch.setattr(QMessageBox, "question", lambda *args: choice)
    item = next(
        dialog.groups.topLevelItem(n)
        for n in range(dialog.groups.topLevelItemCount())
        if dialog.groups.topLevelItem(n).data(0, Qt.ItemDataRole.UserRole) == first
    )
    dialog.groups.setCurrentItem(item)
    application.processEvents()
    if choice == QMessageBox.StandardButton.Cancel:
        assert dialog.group_id == second
        assert dialog.title.text() == "Pending revised scope"
    else:
        assert dialog.group_id == first
    assert session.claims[second]["title"] == (
        "Pending revised scope" if choice == QMessageBox.StandardButton.Save else old_title
    )


@pytest.mark.parametrize("invalid", ["group", "overlap", "context"])
def test_invalid_pending_input_keeps_editor_open_and_blocks_publication(
    dialog, session, application, invalid
):
    if invalid == "group":
        fill_group(dialog)
        dialog.title.clear()
    elif invalid == "overlap":
        two_groups(dialog)
        dialog.overlaps.setCurrentItem(dialog.overlaps.topLevelItem(0))
        application.processEvents()
        dialog.overlap_reason.setPlainText("   ")
    else:
        dialog.complaint_summary.setPlainText("x" * 4001)
    before = session.payload()
    assert not dialog.save_all()
    dialog.publish_pack()
    assert not dialog.publish_requested
    assert dialog.isVisible()
    assert dialog.feedback.text()
    assert session.payload() == before


def test_overlap_save_requires_a_selected_pair(dialog, session):
    before = session.payload()
    assert not dialog.save_overlap()
    assert "Select a pair" in dialog.feedback.text()
    assert session.payload() == before


def test_unused_credit_row_is_ignored_and_selected_row_can_be_removed(dialog, session):
    dialog.remove_credit()
    fill_group(dialog)
    dialog.add_credit()
    assert dialog.save_group()
    assert session.claims[dialog.group_id]["credits"] == []
    dialog.add_credit()
    assert dialog.credits.rowCount() == 1
    dialog.remove_credit()
    assert dialog.credits.rowCount() == 0


@pytest.mark.parametrize(
    "choice",
    [
        QMessageBox.StandardButton.Save,
        QMessageBox.StandardButton.Discard,
        QMessageBox.StandardButton.Cancel,
    ],
)
def test_window_close_protects_pending_complaint_wording(
    dialog, session, application, monkeypatch, choice
):
    before = copy.deepcopy(session.complaint_context)
    dialog.complaint_summary.setPlainText("Pending complaint wording")
    monkeypatch.setattr(QMessageBox, "question", lambda *args: choice)
    dialog.close()
    application.processEvents()
    assert dialog.isVisible() == (choice == QMessageBox.StandardButton.Cancel)
    if choice == QMessageBox.StandardButton.Save:
        assert session.complaint_context["complaint_summary"] == "Pending complaint wording"
    else:
        assert session.complaint_context == before
    if choice == QMessageBox.StandardButton.Discard:
        assert dialog.complaint_summary.toPlainText() == before["complaint_summary"]


def test_publish_requests_owner_job_without_creating_a_blank_claim(dialog, session):
    dialog.publish_pack()
    assert dialog.publish_requested
    assert not dialog.isVisible()
    assert not session.claims
