"""Human correction workflows using real Qt controls, threads and source previews."""

import os
from decimal import Decimal

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QMessageBox, QScrollArea  # noqa: E402

from edf_bill_fetcher.application.review_session import ReviewSession  # noqa: E402
from edf_bill_fetcher.io.persistence.result_store import load_result  # noqa: E402
from edf_bill_fetcher.ui.qt_review import ReviewPreview  # noqa: E402
from tests.test_evidence_review import application as application  # noqa: E402
from tests.test_evidence_review import open_pack, wait_until  # noqa: E402
from tests.test_evidence_review import pack as pack  # noqa: E402
from tests.test_evidence_review import window as window  # noqa: E402


def choose(page, invoice):
    key = next(r["id"] for r in page.session.rows if r["record"].get("Invoice #") == invoice)
    page.select_record(key)
    page.details_tabs.setCurrentWidget(page.editor)
    return key


def test_editor_validates_autosaves_recovers_and_supports_undo_redo(window, application, pack):
    root, path, _, snapshot, _ = pack
    open_pack(window, application, path)
    page = window.review_page
    key = choose(page, "B1")
    assert page.regenerate.isEnabled()
    page.editor.fields["Period Charge (£)"].setText("52.50")
    assert page.editor.apply() is False
    assert "reason" in page.message.text().lower()
    assert page.session.record(key)["Period Charge (£)"] == 40
    page.editor.reason.setPlainText("Checked amount against source")
    page.editor.status.setCurrentIndex(1)
    assert page.editor.apply()
    assert page.session.record(key)["Period Charge (£)"] == 52.50
    wait_until(application, lambda: not page.draft_dirty)
    draft = path.with_suffix(".review-draft.json")
    assert ReviewSession.load(snapshot, draft).record(key)["Period Charge (£)"] == 52.50
    page.undo_button.click()
    assert page.session.record(key)["Period Charge (£)"] == 40
    page.redo_button.click()
    assert page.session.record(key)["Period Charge (£)"] == 52.50
    page.preview_changes.click()
    wait_until(application, lambda: page.impact_worker is None)
    assert "Preview only" in page.impact.text()
    assert "122.50" in page.impact.text()
    assert load_result(path).summary == snapshot.summary
    page.load_analysis(str(path))
    wait_until(application, lambda: page.loader is None)
    assert page.session.record(key)["Period Charge (£)"] == 52.50
    assert page.session.statuses[key]["status"] == "verified"
    assert sorted(p.name for p in root.iterdir()) == [
        "evidence_files",
        "result.json",
        "result.review-draft.json",
    ]


@pytest.mark.parametrize(
    "choice",
    [
        QMessageBox.StandardButton.Cancel,
        QMessageBox.StandardButton.Discard,
        QMessageBox.StandardButton.Save,
    ],
)
def test_unapplied_form_has_no_silent_loss(window, application, pack, monkeypatch, choice):
    _, path, _, _, _ = pack
    open_pack(window, application, path)
    page = window.review_page
    key = choose(page, "B1")
    page.editor.fields["Period Charge (£)"].setText("45")
    page.editor.reason.setPlainText("Verified original")
    monkeypatch.setattr(QMessageBox, "question", lambda *args: choice)
    if choice == QMessageBox.StandardButton.Cancel:
        assert page.allow_leave() is False
        assert page.editor.dirty()
        page.editor.reset()
    else:
        assert page.allow_leave()
        assert page.editor.dirty() is False
        assert page.session.record(key)["Period Charge (£)"] == (
            45 if choice == QMessageBox.StandardButton.Save else 40
        )


def test_save_failure_allows_cancel_or_discard_without_losing_saved_draft(
    window, application, pack, monkeypatch
):
    _, path, _, _, _ = pack
    open_pack(window, application, path)
    page = window.review_page
    key = choose(page, "B1")
    page.session.update(key, {"Period Charge (£)": 55}, True, "Original checked", "verified")
    page.editor.reset()
    page._draft_changed()
    monkeypatch.setattr(
        ReviewSession, "save", lambda *args: (_ for _ in ()).throw(OSError("disk full"))
    )
    monkeypatch.setattr(QMessageBox, "question", lambda *args: QMessageBox.StandardButton.Cancel)
    assert page.allow_leave() is False
    assert page.session.record(key)["Period Charge (£)"] == 55
    monkeypatch.setattr(QMessageBox, "question", lambda *args: QMessageBox.StandardButton.Discard)
    assert page.allow_leave()
    assert page.session.record(key)["Period Charge (£)"] == 40
    assert not page.draft_dirty


def test_revision_comparison_protects_pending_reason_and_rejects_unselected_text(
    window, application, pack, monkeypatch
):
    _, path, _, _, _ = pack
    open_pack(window, application, path)
    page = window.review_page
    panel = page.revisions
    keys = {r["record"].get("Invoice #"): r["id"] for r in page.session.rows}
    panel.older.setCurrentIndex(panel.older.findData(keys["B1"]))
    panel.newer.setCurrentIndex(panel.newer.findData(keys["B2"]))
    panel.decision.setCurrentIndex(1)
    panel.reason.setPlainText("Revised bill checked against original")
    assert panel.dirty()
    panel.refresh(page.session)
    assert panel.reason.toPlainText() == "Revised bill checked against original"
    panel.newer.setEditText("Something not selected")
    assert panel.save_relationship() is False
    panel.newer.setCurrentIndex(panel.newer.findData(keys["B2"]))
    assert panel.save_relationship()
    assert page.session.revisions[(keys["B1"], keys["B2"])]["status"] == "confirmed"
    panel.reason.setPlainText("Unapplied reason")
    monkeypatch.setattr(QMessageBox, "question", lambda *args: QMessageBox.StandardButton.Cancel)
    panel.newer.setCurrentIndex(0)
    assert panel.newer.currentData() == keys["B2"]
    assert panel.reason.toPlainText() == "Unapplied reason"
    monkeypatch.setattr(QMessageBox, "question", lambda *args: QMessageBox.StandardButton.Discard)
    assert panel.allow_leave()


def test_reviewed_pack_button_uses_native_worker_and_loads_new_analysis(
    window, application, pack, tmp_path
):
    _, path, original, snapshot, _ = pack
    source_bytes = original.read_bytes()
    open_pack(window, application, path)
    page = window.review_page
    choose(page, "B1")
    page.editor.fields["Period Charge (£)"].setText("50")
    page.editor.reason.setPlainText("Confirmed source amount")
    assert page.editor.apply()
    window.fields["output_folder"].setText(str(tmp_path))
    window.report_options.update(format="html", sections=["detailed_findings"])
    page.regenerate.click()
    assert window._busy
    assert window.current_job.mode == "review"
    assert window.current_job.review_draft is not None
    wait_until(application, lambda: not window._busy and page.loader is None)
    assert window.last_outcome.status in {"complete", "review_required"}, window.last_outcome.error
    assert window.last_outcome.snapshot.summary.charges == Decimal("120.00")
    wait_until(
        application,
        lambda: (
            page.review.snapshot.summary.charges == Decimal("120.00") and not page.has_workers()
        ),
    )
    assert page.analysis_path != str(path)
    assert not page.session.decisions
    assert load_result(path).summary == snapshot.summary
    assert original.read_bytes() == source_bytes


def test_cancelled_preview_writes_nothing(pack):
    root, path, _, snapshot, _ = pack
    before = sorted(p.name for p in root.iterdir())
    worker = ReviewPreview(ReviewSession(snapshot), str(path))
    worker.cancel_event.set()
    results: list[object] = []
    worker.ready.connect(results.append)
    worker.run()
    assert results == []
    assert sorted(p.name for p in root.iterdir()) == before


@pytest.mark.parametrize(
    ("width", "height", "font"), [(1160, 800, 13), (820, 600, 13), (920, 640, 20)]
)
def test_correction_and_revision_controls_remain_reachable(
    window, application, pack, width, height, font
):
    _, path, _, _, _ = pack
    open_pack(window, application, path)
    window.resize(width, height)
    stylesheet = application.styleSheet()
    application.setStyleSheet(stylesheet + f"\nQWidget {{font-size: {font}px;}}")
    try:
        page = window.review_page
        choose(page, "B1")
        for panel in (page.editor, page.revisions):
            page.details_tabs.setCurrentWidget(panel)
            application.processEvents()
            assert panel.isVisible()
            area = panel.findChild(QScrollArea)
            assert area is not None
            area.verticalScrollBar().setValue(area.verticalScrollBar().maximum())
            application.processEvents()
            assert page.regenerate.isVisible()
            assert window.width() == width
        assert page.details_tabs.height() >= 100
    finally:
        application.setStyleSheet(stylesheet)
