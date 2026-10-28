"""Real Qt grouping and credit edits must preserve drafts, figures and source choices."""

import pytest
from PySide6.QtCore import QCoreApplication, QEvent, Qt
from PySide6.QtWidgets import QComboBox, QDialogButtonBox, QLineEdit, QMessageBox

from edf_bill_fetcher.application.analysis_service import analyse_prepared
from edf_bill_fetcher.application.legacy_adapter import capture_legacy
from edf_bill_fetcher.application.preparation_service import prepare_snapshot
from edf_bill_fetcher.application.review_session import ReviewSession
from edf_bill_fetcher.ui.qt_claims import ClaimDialog
from tests.test_claim_reconciliation import evidence
from tests.test_evidence_review import application as application
from tests.test_qt_desktop import result_snapshot
from tests.test_rebilling_ledger import set_position


@pytest.fixture
def session():
    items, _ = evidence()
    records = [{**item["record"], "Review Row ID": item["id"]} for item in items]
    snapshot = analyse_prepared(
        prepare_snapshot(capture_legacy(records, {}, result_snapshot().prepared.extraction.request))
    )
    review = ReviewSession(snapshot)
    for key in ("a", "b"):
        set_position(review, key)
    return review


@pytest.fixture
def dialog(application, session):
    widget = ClaimDialog(session)
    widget.show()
    application.processEvents()
    yield widget
    widget.hide()
    widget.deleteLater()
    QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
    application.processEvents()


def fill_group(dialog, members=("a", "b")):
    dialog.title.setText("One consolidated challenge to repeated demands")
    dialog.challenged.setText("120")
    dialog.reduction.setText("120")
    dialog.reason.setPlainText(
        "Original bills and replacement checked: the consolidated claim is entered once, not summed."
    )
    for n in range(dialog.members.count()):
        item = dialog.members.item(n)
        item.setCheckState(
            Qt.CheckState.Checked
            if item.data(Qt.ItemDataRole.UserRole) in members
            else Qt.CheckState.Unchecked
        )
    dialog.status.setCurrentIndex(dialog.status.findData("reviewed"))


def add_allocation(dialog, amount, treatment):
    dialog.add_credit()
    row = dialog.credits.rowCount() - 1
    source, field, kind, reason = (dialog.credits.cellWidget(row, n) for n in range(4))
    assert isinstance(source, QComboBox)
    assert isinstance(field, QLineEdit)
    assert isinstance(kind, QComboBox)
    assert isinstance(reason, QLineEdit)
    source.setCurrentIndex(source.findData("c"))
    field.setText(amount)
    kind.setCurrentIndex(kind.findData(treatment))
    reason.setText("Verified allocation on the original credit source")


def test_group_editor_distinguishes_cancellation_offsets_from_actual_relief(dialog, session):
    original = session.snapshot.summary
    fill_group(dialog)
    add_allocation(dialog, "70", "cancellation_offset")
    add_allocation(dialog, "30", "already_applied_relief")
    assert dialog.save_group()
    model = session.claim_model()
    assert model["ready"]
    assert model["proposed_reduction"] == "90.00"
    assert "90.00" in dialog.summary.text()
    assert session.snapshot.summary == original
    restored = ReviewSession.restore(session.snapshot, session.payload())
    assert restored.claim_state() == session.claim_state()
    session.undo()
    assert not session.claims
    session.redo()
    assert session.claim_model()["proposed_reduction"] == "90.00"


def test_unknown_typed_credit_sources_and_excess_allocations_are_not_published(dialog, session):
    fill_group(dialog)
    add_allocation(dialog, "101", "already_applied_relief")
    assert not dialog.save_group()
    assert "exceeds" in dialog.feedback.text()
    assert not session.claims
    source = dialog.credits.cellWidget(0, 0)
    assert isinstance(source, QComboBox)
    source.setEditText("Unselected source text")
    assert not dialog.save_group()
    assert "listed credit source" in dialog.feedback.text()
    assert source.currentText() == "Unselected source text"


@pytest.mark.parametrize(
    "choice",
    [
        QMessageBox.StandardButton.Save,
        QMessageBox.StandardButton.Discard,
        QMessageBox.StandardButton.Cancel,
    ],
)
def test_pending_claim_edits_are_saved_discarded_or_retained_explicitly(
    dialog, session, monkeypatch, choice
):
    fill_group(dialog)
    monkeypatch.setattr(QMessageBox, "question", lambda *args: choice)
    dialog.reject()
    assert dialog.isVisible() == (choice == QMessageBox.StandardButton.Cancel)
    assert bool(session.claims) == (choice == QMessageBox.StandardButton.Save)
    if choice == QMessageBox.StandardButton.Cancel:
        assert dialog.title.text().startswith("One consolidated")


def test_complaint_wording_saves_without_creating_an_empty_claim(dialog, session):
    dialog.complaint_summary.setPlainText(
        "Historic additional charges are disputed pending their first-demand evidence."
    )
    dialog.requested_outcome.setPlainText(
        "Cancel protected charges and provide a substantiated balance."
    )
    dialog.include_pack.setChecked(False)
    assert dialog.save_all()
    assert not session.claims
    assert session.complaint_context["include_pack"] is False
    assert session.complaint_context["requested_outcome"].startswith("Cancel protected")


def test_source_navigation_uses_the_selected_canonical_record(dialog):
    fill_group(dialog)
    assert dialog.save_group()
    selected: list[str] = []
    dialog.show_record.connect(selected.append)
    n = next(
        n
        for n in range(dialog.members.count())
        if dialog.members.item(n).data(Qt.ItemDataRole.UserRole) == "b"
    )
    dialog.members.setCurrentRow(n)
    dialog.reveal_bill()
    assert selected == ["b"]
    assert not dialog.isVisible()


def test_dialog_actions_remain_visible_at_small_size_and_large_text(dialog, application):
    previous = application.styleSheet()
    application.setStyleSheet(previous + "\nQWidget {font-size:20px;}")
    try:
        dialog.resize(760, 600)
        application.processEvents()
        controls = dialog.findChild(QDialogButtonBox)
        assert controls is not None
        for control in (
            dialog.publish,
            controls.button(QDialogButtonBox.StandardButton.Save),
            controls.button(QDialogButtonBox.StandardButton.Close),
        ):
            assert dialog.rect().contains(control.mapTo(dialog, control.rect().bottomRight()))
        assert dialog.width() == 760
    finally:
        application.setStyleSheet(previous)
