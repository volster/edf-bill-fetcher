"""Real Qt dispute edits, source navigation and draft recovery."""

import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtWidgets import QMessageBox, QScrollArea  # noqa: E402

from edf_bill_fetcher.application.review_session import ReviewSession  # noqa: E402
from tests.test_evidence_review import application as application  # noqa: E402
from tests.test_evidence_review import open_pack, wait_until  # noqa: E402
from tests.test_evidence_review import pack as pack  # noqa: E402
from tests.test_evidence_review import window as window  # noqa: E402


def select_bill(panel, key):
    for group in range(panel.tree.topLevelItemCount()):
        parent = panel.tree.topLevelItem(group)
        for n in range(parent.childCount()):
            item = parent.child(n)
            if item.data(0, Qt.ItemDataRole.UserRole) == key:
                panel.tree.setCurrentItem(item)
                assert panel.row_id == key
                return
    raise AssertionError("Missing canonical bill in rebilling history")


def test_rebilling_form_validates_autosaves_reopens_and_undoes_without_changing_balance(
    window, application, pack
):
    _, path, _, snapshot, _ = pack
    open_pack(window, application, path)
    page, panel = window.review_page, window.review_page.rebilling
    key = next(row["id"] for row in page.session.rows if row["record"].get("Invoice #") == "B2")
    select_bill(panel, key)
    panel.use_adjustment.click()
    assert panel.challenged.text() == panel.reduction.text() == "70.00"
    assert panel.save_position() is False
    assert "scope" in page.message.text().lower()
    panel.basis.setPlainText("Challenge full adjustment pending earlier bills")
    panel.request.setPlainText("Earlier demands and meter read allocation")
    panel.first_demand.setText("01/07/2026")
    assert panel.save_position()
    wait_until(application, lambda: not page.draft_dirty)
    draft = ReviewSession.load(snapshot, path.with_suffix(".review-draft.json"))
    assert draft.schedule[key]["proposed_reduction"] == "70.00"
    assert draft.schedule[key]["first_demand_date"] == "2026-07-01"
    assert page.review.snapshot.summary == snapshot.summary
    page.undo_button.click()
    assert not page.session.schedule
    page.redo_button.click()
    assert key in page.session.schedule
    panel.clear_position()
    assert not page.session.schedule
    page.undo_button.click()
    assert key in page.session.schedule
    panel._show()
    wait_until(application, lambda: not page.has_workers())
    assert page.editor.row_id == key
    page.load_analysis(str(path))
    wait_until(application, lambda: not page.has_workers())
    assert page.session.schedule[key]["proposed_reduction"] == "70.00"


@pytest.mark.parametrize(
    "choice",
    [
        QMessageBox.StandardButton.Cancel,
        QMessageBox.StandardButton.Discard,
        QMessageBox.StandardButton.Save,
    ],
)
def test_unapplied_dispute_position_survives_refresh_and_navigation(
    window, application, pack, monkeypatch, choice
):
    _, path, _, _, _ = pack
    open_pack(window, application, path)
    page, panel = window.review_page, window.review_page.rebilling
    key = panel.row_id
    panel.challenged.setText("40")
    panel.reduction.setText("40")
    panel.status.setCurrentIndex(panel.status.findData("challenged"))
    panel.basis.setPlainText("Compare original charge with full revision history")
    panel.refresh(page.session)
    assert panel.dirty()
    assert panel.challenged.text() == "40"
    monkeypatch.setattr(QMessageBox, "question", lambda *args: choice)
    if choice == QMessageBox.StandardButton.Cancel:
        assert page.allow_leave() is False
        assert panel.dirty()
        monkeypatch.setattr(
            QMessageBox, "question", lambda *args: QMessageBox.StandardButton.Discard
        )
        assert page.allow_leave()
    else:
        assert page.allow_leave()
        assert bool(page.session.schedule) == (choice == QMessageBox.StandardButton.Save)
        if page.session.schedule:
            assert page.session.schedule[key]["basis"].startswith("Compare original")


def test_rebilling_controls_scroll_at_small_window_and_large_fonts(window, application, pack):
    _, path, _, _, _ = pack
    open_pack(window, application, path)
    window.resize(820, 600)
    page, panel = window.review_page, window.review_page.rebilling
    page.details_tabs.setCurrentWidget(panel)
    previous = application.styleSheet()
    application.setStyleSheet(previous + "\nQWidget {font-size: 20px;}")
    try:
        application.processEvents()
        area = panel.findChild(QScrollArea)
        assert area is not None
        area.ensureWidgetVisible(panel.save_button)
        application.processEvents()
        assert panel.save_button.isVisible()
        assert panel.request.isEnabled()
        assert window.width() == 820
    finally:
        application.setStyleSheet(previous)


def test_saving_during_document_navigation_keeps_tree_and_form_on_the_same_bill(
    window, application, pack, monkeypatch
):
    _, path, _, _, _ = pack
    open_pack(window, application, path)
    page, panel = window.review_page, window.review_page.rebilling
    keys = [row["id"] for row in page.session.rows if row["record"].get("Entry Type") == "New Bill"]
    first = panel.row_id
    other = next(key for key in keys if key != first)
    panel.use_adjustment.click()
    panel.basis.setPlainText("Full adjustment pending earlier demands")
    monkeypatch.setattr(QMessageBox, "question", lambda *args: QMessageBox.StandardButton.Save)
    select_bill(panel, other)
    assert first in page.session.schedule
    assert panel.tree.currentItem().data(0, Qt.ItemDataRole.UserRole) == other
    assert panel.row_id == other
    assert not panel.dirty()
