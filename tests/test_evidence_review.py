"""Evidence review contracts using real saved results, colored PDFs and Qt workers."""

import hashlib
import os
import shutil
from dataclasses import replace
from datetime import date, datetime, timezone
from pathlib import Path
from threading import Event
from time import monotonic, sleep

import pytest
from reportlab.pdfgen.canvas import Canvas

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QCoreApplication, QEvent, Qt  # noqa: E402
from PySide6.QtWidgets import QApplication, QFileDialog  # noqa: E402

from edf_bill_fetcher.application.analysis_service import analyse_prepared  # noqa: E402
from edf_bill_fetcher.application.evidence_review import (  # noqa: E402
    ReviewSource,
    build_evidence_review,
    page_range,
)
from edf_bill_fetcher.application.legacy_adapter import capture_legacy  # noqa: E402
from edf_bill_fetcher.application.preparation_service import prepare_snapshot  # noqa: E402
from edf_bill_fetcher.application.run_service import RunCancelled  # noqa: E402
from edf_bill_fetcher.domain.analysis import analytical_fingerprint  # noqa: E402
from edf_bill_fetcher.domain.json_values import freeze_json, object_value, rows_value  # noqa: E402
from edf_bill_fetcher.domain.requests import RunRequest  # noqa: E402
from edf_bill_fetcher.io.persistence.result_store import load_result, save_result  # noqa: E402
from edf_bill_fetcher.ui import qt_review  # noqa: E402
from edf_bill_fetcher.ui.qt_app import EvidenceWindow  # noqa: E402
from edf_bill_fetcher.ui.qt_review import (  # noqa: E402
    EvidenceTable,
    PdfPreview,
    PreviewRequest,
    row_value,
)
from edf_bill_fetcher.ui.qt_theme import apply_theme  # noqa: E402


@pytest.fixture(scope="session")
def application():
    app = QApplication.instance() or QApplication([])
    assert isinstance(app, QApplication)
    app.setQuitOnLastWindowClosed(False)
    apply_theme(app)
    yield app
    QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
    app.processEvents()


def wait_until(application, predicate):
    deadline = monotonic() + 15
    while not predicate() and monotonic() < deadline:
        application.processEvents()
        sleep(0.01)
    application.processEvents()
    assert predicate(), "Review operation did not finish"


@pytest.fixture
def window(application, tmp_path):
    widget = EvidenceWindow(tmp_path / "settings.json")
    widget.show()
    widget._navigate(3)
    application.processEvents()
    yield widget
    widget.close()
    wait_until(application, lambda: not widget.review_page.has_workers() and not widget.isVisible())
    widget.deleteLater()
    QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
    application.processEvents()


def make_pdf(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    pdf = Canvas(str(path), pagesize=(200, 200))
    for page, color in enumerate(((1, 0, 0), (0, 1, 0), (0, 0, 1)), start=1):
        pdf.setFillColorRGB(*color)
        pdf.rect(0, 0, 200, 200, fill=1, stroke=0)
        pdf.setFillColorRGB(0, 0, 0)
        pdf.drawString(30, 20, f"Original page {page}")
        pdf.showPage()
    pdf.save()


@pytest.fixture
def pack(tmp_path):
    root = tmp_path / "pack"
    original = root / "evidence_files" / "merged.pdf"
    make_pdf(original)
    digest = hashlib.sha256(original.read_bytes()).hexdigest()
    source_id = hashlib.sha256(b"merged occurrence").hexdigest()
    base = {
        "Source ID": source_id,
        "Content Hash": digest,
        "Original Filename": "original merged.pdf",
        "Attachment Name": "original merged.pdf",
        "Account": "ACCOUNT-1",
        "Fuel": "Electricity",
        "Entry Type": "New Bill",
        "Amount Kind": "Balance",
        "Reading": "Actual",
        "Source PDF Text": "Evidence with [literal] search & characters",
        "Page From": 1,
        "Page To": 1,
    }
    records = [
        {
            **base,
            "Invoice #": "B1",
            "Date": "01/01/2026",
            "Amount (£)": 100,
            "Period Charge (£)": 40,
        },
        {
            **base,
            "Invoice #": "B2",
            "Date": "01/07/2026",
            "Amount (£)": 160,
            "Period Charge (£)": 70,
            "Page From": 3,
            "Page To": 3,
        },
        {
            "Account": "ACCOUNT-1",
            "Fuel": "Electricity",
            "Entry Type": "Payment",
            "Date": "02/07/2026",
            "Amount (£)": -25,
            "Amount Kind": "Transaction",
        },
        {
            **base,
            "Invoice #": "B1",
            "Date": "01/01/2026",
            "Amount (£)": 100,
            "Period Charge (£)": 40,
        },
    ]
    state = {
        "source_manifest": {
            source_id: {
                "Source ID": source_id,
                "Attachment Name": "original merged.pdf",
                "Original Filename": "original merged.pdf",
                "Content Hash": digest,
                "Original Path": str(tmp_path / "gone" / "merged.pdf"),
                "Saved Path": "evidence_files/merged.pdf",
            }
        },
        "filtered_records": [{"Entry Type": "New Bill", "Invoice #": "SMALL", "Amount (£)": 1}],
        "error_log": ["A preserved extraction review note"],
    }
    request = RunRequest(
        "review-test",
        datetime(2026, 7, 31, tzinfo=timezone.utc),
        date(2026, 7, 31),
        (),
        freeze_json({"as_of": "2026-07-31", "use_dedup": True}),
    )
    snapshot = analyse_prepared(prepare_snapshot(capture_legacy(records, state, request)))
    path = root / "result.json"
    save_result(snapshot, path, result_id="review-test")
    return root, path, original, snapshot, source_id


def open_pack(window, application, path):
    window.review_page.load_analysis(str(path))
    wait_until(
        application,
        lambda: window.review_page.review is not None and not window.review_page.has_workers(),
    )
    return window.review_page


def select_invoice(page, invoice):
    for index in range(page.proxy.rowCount()):
        mapped = page.proxy.mapToSource(page.proxy.index(index, 0))
        if page.model.rows[mapped.row()].get("Invoice #") == invoice:
            page.table.selectRow(index)
            return
    pytest.fail(f"Invoice {invoice} not shown")


def test_review_preserves_result_originals_fingerprint_and_accepted_models(pack, monkeypatch):
    root, path, original, snapshot, _ = pack
    before = {p.relative_to(root): p.read_bytes() for p in root.rglob("*") if p.is_file()}
    monkeypatch.setattr(
        "edf_bill_fetcher.application.analysis_service.analyse_prepared",
        lambda *_: pytest.fail("Review must not compute again"),
    )
    review = build_evidence_review(load_result(path), str(root))
    assert review.snapshot == snapshot
    assert review.snapshot.input_fingerprint == snapshot.input_fingerprint
    assert review.snapshot.summary == snapshot.summary
    assert {p.relative_to(root): p.read_bytes() for p in root.rglob("*") if p.is_file()} == before
    categories = [entry.category for entry in review.entries]
    assert categories.count("Bills & charges") == 2
    assert categories.count("Payments & credits") == 1
    assert categories.count("Duplicates") == 1
    assert categories.count("Filtered out") == 1
    assert categories.count("Findings") >= 1
    entry = next(
        entry for entry in review.entries if object_value(entry.fields).get("Invoice #") == "B2"
    )
    assert entry.sources[0].pages == (3, 3)
    assert entry.sources[0].path == str(original)


@pytest.mark.parametrize(
    "mode",
    ["moved", "missing", "changed", "escaped", "original_only", "original_changed", "legacy"],
)
def test_source_resolution_and_honest_availability(pack, tmp_path, mode):
    root, _, original, snapshot, source_id = pack
    state = object_value(snapshot.prepared.extraction.legacy_state)
    expected: str | None
    if mode == "moved":
        moved = tmp_path / "moved"
        shutil.move(root, moved)
        review = build_evidence_review(load_result(moved / "result.json"), str(moved))
        expected = str(moved / "evidence_files/merged.pdf")
    elif mode in {"original_only", "original_changed", "legacy", "escaped"}:
        metadata = state["source_manifest"][source_id]
        if mode == "legacy":
            state["source_manifest"] = {}
            state["saved_evidence_paths"] = {source_id: "evidence_files/merged.pdf"}
        elif mode == "escaped":
            metadata["Saved Path"] = "../../outside.pdf"
            metadata["Original Path"] = str(original)
        else:
            metadata.pop("Saved Path")
            metadata["Original Path"] = str(original)
            if mode == "original_changed":
                original.write_bytes(b"changed")
        # Capture the changed metadata, keeping the original financial input.
        request = snapshot.prepared.extraction.request
        originals = [object_value(obs.raw) for obs in snapshot.prepared.extraction.observations]
        altered = analyse_prepared(prepare_snapshot(capture_legacy(originals, state, request)))
        review = build_evidence_review(altered, str(root))
        expected = str(original) if mode in {"original_only", "legacy"} else None
    else:
        if mode == "missing":
            original.unlink()
        else:
            original.write_bytes(b"changed")
        review = build_evidence_review(snapshot, str(root))
        expected = None
    assert review.entries[0].sources[0].path == expected
    if expected is None:
        assert review.notes


@pytest.mark.parametrize(
    ("start", "end", "expected"),
    [
        (1, 3, (1, 3)),
        (3.0, 3.0, (3, 3)),
        (True, 2, None),
        (0, 1, None),
        (3, 2, None),
        (1.5, 2, None),
        ("1", "2", None),
        (None, 2, None),
        (1, float("inf"), None),
        (1, 10**1000, None),
        (1, None, None),
    ],
)
def test_invalid_page_metadata_never_becomes_a_fake_page(start, end, expected):
    assert page_range({"Page From": start, "Page To": end}) == expected


def test_cancellation_before_loading_source_files(pack):
    root, _, _, snapshot, _ = pack
    with pytest.raises(RunCancelled):
        build_evidence_review(snapshot, str(root), lambda: True)


def test_finding_sources_require_explicit_date_match_selection(window, application, pack):
    _, path, _, _, _ = pack
    page = open_pack(window, application, path)
    page.category.setCurrentText("Findings")
    application.processEvents()
    assert page.proxy.rowCount() >= 1
    assert page._selected_source is None
    assert page.source_choices.currentData() is None
    assert "date only" in page.source_status.text()
    assert page._image is None
    assert not page.open_source.isEnabled()
    assert page.source_choices.count() >= 2
    page.source_choices.setCurrentIndex(1)
    wait_until(application, lambda: not page.has_workers())
    assert page._image is not None


def test_search_categories_availability_numeric_sort_and_literal_text(window, application, pack):
    _, path, _, _, _ = pack
    page = open_pack(window, application, path)
    assert page.proxy.rowCount() == 3
    page.search.setText("[literal]")
    page._filter()
    assert page.proxy.rowCount() == 2
    page.search.setText("no matches")
    page._filter()
    assert page.proxy.rowCount() == 0
    assert page._image is None
    assert not page.source_choices.isEnabled()
    assert "No matching" in page.count.text()
    page.search.clear()
    page.category.setCurrentText("Payments & credits")
    assert page.proxy.rowCount() == 1
    assert "-25" in page.details.toPlainText()
    assert "No original source" in page.source_status.text()
    page.category.setCurrentText("All records")
    page.availability.setCurrentIndex(1)
    assert page.proxy.rowCount() == 2
    page.availability.setCurrentIndex(2)
    assert page.proxy.rowCount() == 1
    page.availability.setCurrentIndex(0)
    page.table.sortByColumn(3, Qt.SortOrder.AscendingOrder)
    assert page.proxy.data(page.proxy.index(0, 3)) == "-25.00"
    assert not page.details.isUndoRedoEnabled() or page.details.isReadOnly()
    assert page.extracted_text.isReadOnly()
    assert not page.model.flags(page.model.index(0, 0)) & Qt.ItemFlag.ItemIsEditable


def test_merged_pdf_page_accuracy_navigation_zoom_and_rapid_selection(window, application, pack):
    _, path, _, _, _ = pack
    page = open_pack(window, application, path)
    assert page._image.pixelColor(100, 100).red() > 240
    select_invoice(page, "B2")
    wait_until(application, lambda: not page.has_workers())
    assert page.page.value() == 3
    assert page._image.pixelColor(100, 100).blue() > 240
    page.previous.click()
    wait_until(application, lambda: not page.has_workers())
    assert page.page.value() == 2
    assert page._image.pixelColor(100, 100).green() > 240
    for invoice in ("B1", "B2", "B1", "B2"):
        select_invoice(page, invoice)
    wait_until(application, lambda: not page.has_workers())
    assert page.page.value() == 3
    assert page._image.pixelColor(100, 100).blue() > 240
    page.zoom.setCurrentText("200%")
    wait_until(application, lambda: not page.has_workers())
    assert page._image.width() == 400
    assert "[literal]" in page.extracted_text.toPlainText()


def test_missing_source_retains_rows_notes_and_no_stale_preview(window, application, pack):
    _, path, original, _, _ = pack
    original.unlink()
    page = open_pack(window, application, path)
    assert page.proxy.rowCount() == 3
    assert page._image is None
    assert not page.open_source.isEnabled()
    assert "unavailable" in page.source_status.text()
    assert "Original evidence unavailable" in page.run_notes.toPlainText()
    assert "preserved extraction" in page.run_notes.toPlainText()


def test_failed_load_preserves_previous_analysis(window, application, pack, tmp_path):
    _, path, _, _, _ = pack
    page = open_pack(window, application, path)
    accepted = page.review
    malformed = tmp_path / "bad.json"
    malformed.write_text('{"schema_version":3,"schema_version":3}')
    page.load_analysis(str(malformed))
    wait_until(application, lambda: page.loader is None)
    assert page.review is accepted
    assert page.analysis_path == str(path)
    assert "could not be opened" in page.message.text()
    assert "Duplicate JSON" in page.message.text()
    assert page.open_button.isEnabled()


@pytest.mark.parametrize("width", [820, 1160])
def test_small_and_normal_windows_keep_preview_and_actions_accessible(
    window, application, pack, width
):
    _, path, _, _, _ = pack
    page = open_pack(window, application, path)
    window.resize(width, 720 if width == 1160 else 600)
    application.processEvents()
    assert window.width() == width
    assert window.height() == (720 if width == 1160 else 600)
    for action in (window.run_button, page.open_button, page.open_source, page.page):
        assert window.rect().contains(action.mapTo(window, action.rect().bottomRight()))


@pytest.mark.parametrize("failure", ["changed", "invalid_pages", "corrupt_pdf"])
def test_preview_rechecks_bytes_and_rejects_unreliable_ranges(application, pack, failure):
    _, _, original, _, source_id = pack
    digest = hashlib.sha256(original.read_bytes()).hexdigest()
    pages = (1, 99) if failure == "invalid_pages" else (1, 1)
    if failure in {"changed", "corrupt_pdf"}:
        original.write_bytes(b"not a PDF")
        if failure == "corrupt_pdf":
            digest = hashlib.sha256(original.read_bytes()).hexdigest()
    source = ReviewSource("Original", str(original), digest, pages, "Verified", source_id)
    worker = PdfPreview(PreviewRequest(source, 1, 1))
    errors: list[str] = []
    images = []
    worker.failed.connect(errors.append)
    worker.rendered.connect(lambda image, count: images.append(image))
    worker.start()
    wait_until(application, lambda: not worker.isRunning() and bool(errors))
    assert worker.wait(1000)
    assert not images
    assert "changed" in errors[0] if failure == "changed" else bool(errors[0])


def test_close_during_preview_waits_for_native_cleanup(window, application, pack, monkeypatch):
    _, path, _, _, _ = pack
    page = open_pack(window, application, path)
    entered, release = Event(), Event()
    original_run = PdfPreview.run

    def held_run(worker):
        entered.set()
        assert release.wait(10)
        original_run(worker)

    monkeypatch.setattr(PdfPreview, "run", held_run)
    select_invoice(page, "B2")
    assert entered.wait(2)
    window.close()
    assert window.isVisible()
    assert window._closing
    release.set()
    wait_until(application, lambda: not page.has_workers() and not window.isVisible())


def test_latest_run_review_reuses_snapshot_and_saved_sources(
    window, application, pack, monkeypatch
):
    root, path, _, snapshot, _ = pack
    monkeypatch.setattr(
        qt_review,
        "load_result",
        lambda *_: pytest.fail("In-memory accepted snapshot should be reused"),
    )
    window._navigate(0)
    window._review_pending = snapshot, str(root), str(path)
    window.nav_buttons[3].click()
    wait_until(application, lambda: not window.review_page.has_workers())
    assert window.review_page.review.snapshot is snapshot
    assert window.review_page._image is not None
    assert window.pages.currentIndex() == 3


def test_browse_and_explicit_open_original_callback(window, application, pack, monkeypatch):
    _, path, original, _, _ = pack
    monkeypatch.setattr(QFileDialog, "getOpenFileName", lambda *_: (str(path), ""))
    window.run_button.click()
    wait_until(
        application,
        lambda: window.review_page.review is not None and not window.review_page.has_workers(),
    )
    urls = []

    def open_url(url):
        urls.append(url)
        return True

    monkeypatch.setattr(qt_review.QDesktopServices, "openUrl", open_url)
    assert not urls
    window.review_page.open_source.click()
    wait_until(application, lambda: not window.review_page.has_workers())
    assert Path(urls[0].toLocalFile()) == original
    original.write_bytes(b"changed after preview")
    window.review_page.open_source.click()
    wait_until(application, lambda: not window.review_page.has_workers())
    assert len(urls) == 1
    assert "changed" in window.review_page.message.text()


def test_null_values_stay_unavailable_and_money_basis_is_explicit():
    assert row_value({"Balance (£)": None}) == (None, "Unavailable")
    assert row_value({"Amount (£)": -30, "Amount Kind": "Transaction"}) == (-30, "Transaction")
    assert row_value({"Finding": "Issue", "Amount (£)": 20}) == (20, "Finding amount")
    model = EvidenceTable()
    assert model.rowCount() == 0
    assert model.columnCount() == 6


def altered_snapshot(snapshot, state, records=None):
    originals = (
        records
        if records is not None
        else [object_value(obs.raw) for obs in snapshot.prepared.extraction.observations]
    )
    return analyse_prepared(
        prepare_snapshot(capture_legacy(originals, state, snapshot.prepared.extraction.request))
    )


def test_authoritative_source_id_never_falls_back_to_wrong_filename(pack):
    root, _, _, snapshot, source_id = pack
    state = object_value(snapshot.prepared.extraction.legacy_state)
    metadata = state["source_manifest"].pop(source_id)
    state["source_manifest"]["wrong-source-id"] = metadata
    review = build_evidence_review(altered_snapshot(snapshot, state), str(root))
    assert review.entries[0].sources[0].path is None
    assert review.entries[0].sources[0].source_id == source_id


def test_ambiguous_display_filename_does_not_choose_an_original(pack):
    root, _, _, snapshot, source_id = pack
    state = object_value(snapshot.prepared.extraction.legacy_state)
    state["source_manifest"]["second-occurrence"] = state["source_manifest"][source_id].copy()
    records = [object_value(obs.raw) for obs in snapshot.prepared.extraction.observations]
    for row in records:
        row.pop("Source ID", None)
    review = build_evidence_review(altered_snapshot(snapshot, state, records), str(root))
    assert review.entries[0].sources[0].path is None
    assert review.entries[0].sources[0].status == "Ambiguous filename mapping"


def test_duplicate_source_occurrences_preserve_all_original_choices(pack):
    root, _, _, snapshot, source_id = pack
    state = object_value(snapshot.prepared.extraction.legacy_state)
    second = "second-occurrence"
    state["source_manifest"][second] = {
        **state["source_manifest"][source_id],
        "Original Filename": "email attachment.pdf",
    }
    records = [object_value(obs.raw) for obs in snapshot.prepared.extraction.observations]
    records[0]["Source IDs"] = [source_id, second, second]
    review = build_evidence_review(altered_snapshot(snapshot, state, records), str(root))
    choices = review.entries[0].sources
    assert len(choices) == 2
    assert {choice.source_id for choice in choices} == {source_id, second}
    assert {choice.title for choice in choices} == {"original merged.pdf", "email attachment.pdf"}
    assert all(choice.path for choice in choices)


def test_same_date_findings_do_not_invent_an_account_or_pdf_link(pack):
    root, _, _, snapshot, _ = pack
    state = object_value(snapshot.prepared.extraction.legacy_state)
    records = [object_value(obs.raw) for obs in snapshot.prepared.extraction.observations]
    records.append({**records[1], "Account": "ACCOUNT-2", "Invoice #": "OTHER-ACCOUNT"})
    review = build_evidence_review(altered_snapshot(snapshot, state, records), str(root))
    findings = [entry for entry in review.entries if entry.category == "Findings"]
    related = next(entry for entry in findings if len(entry.sources) >= 2)
    assert "date only" in related.relation_note
    assert "Account" not in object_value(related.fields)
    assert {"ACCOUNT-1", "ACCOUNT-2"} <= {
        account
        for source in related.sources
        for account in ("ACCOUNT-1", "ACCOUNT-2")
        if account in source.title
    }


def test_non_pdf_original_has_explicit_open_without_false_pdf_preview(window, application, pack):
    root, _, original, snapshot, source_id = pack
    text_file = original.with_suffix(".html")
    text_file.write_text("<p>Original history</p>")
    state = object_value(snapshot.prepared.extraction.legacy_state)
    state["source_manifest"][source_id].update(
        {
            "Saved Path": "evidence_files/merged.html",
            "Content Hash": hashlib.sha256(text_file.read_bytes()).hexdigest(),
            "Original Filename": "history.html",
        }
    )
    page = window.review_page
    page.load_analysis(snapshot=altered_snapshot(snapshot, state), directory=str(root))
    wait_until(application, lambda: page.review is not None and not page.has_workers())
    assert page._image is None
    assert page.open_source.isEnabled()
    assert not page.page.isEnabled()
    assert "PDF originals" in page.canvas.text()


def test_close_during_result_loading_keeps_worker_alive_until_cancelled(
    window, application, pack, monkeypatch
):
    _, path, _, _, _ = pack
    entered, release = Event(), Event()
    real_build = qt_review.build_evidence_review

    def held_build(*args):
        entered.set()
        assert release.wait(10)
        return real_build(*args)

    monkeypatch.setattr(qt_review, "build_evidence_review", held_build)
    window.review_page.load_analysis(str(path))
    assert entered.wait(2)
    window.close()
    assert window.isVisible()
    release.set()
    wait_until(application, lambda: not window.review_page.has_workers() and not window.isVisible())
    assert window.review_page.review is None


def test_actual_extraction_and_saved_report_run_offer_correct_review_sources(
    window, application, tmp_path
):
    sources = tmp_path / "pdfs"
    sources.mkdir()
    shutil.copy2(Path(__file__).parent / "fixtures/sample_bill.pdf", sources / "bill.pdf")
    window._navigate(0)
    window.sources["pdf_dir"].entry.setText(str(sources))
    window.fields["output_folder"].setText(str(tmp_path / "outputs"))
    window.run_button.click()
    wait_until(application, lambda: not window._busy and window.last_outcome is not None)
    extraction = window.last_outcome.snapshot
    original_directory = Path(window.current_job.directory)
    original_result = original_directory / "result.json"
    result_bytes = original_result.read_bytes()
    window._navigate(3)
    wait_until(application, lambda: not window.review_page.has_workers())
    assert window.review_page.review.snapshot is extraction
    assert window.review_page._image is not None
    assert Path(window.review_page._selected_source.path).is_relative_to(original_directory)
    window._navigate(1)
    window.report_input.entry.setText(str(original_result))
    window.report_options = {"format": "html", "sections": ["cover"]}
    window.run_button.click()
    wait_until(application, lambda: not window._busy and window.last_outcome is not None)
    assert window.current_job.directory != str(original_directory)
    window._navigate(3)
    wait_until(application, lambda: not window.review_page.has_workers())
    assert window.review_page.review.directory == str(original_directory)
    assert window.review_page._image is not None
    assert original_result.read_bytes() == result_bytes


@pytest.mark.parametrize("end", [10**1000, "invalid"], ids=["oversized", "nonnumeric"])
def test_invalid_recorded_page_range_keeps_valid_source_for_manual_inspection(
    window, application, pack, end
):
    root, _, _, snapshot, _ = pack
    records = rows_value(snapshot.prepared.records)
    for row in records:
        if row.get("Source ID"):
            row["Page To"] = end
    page = window.review_page
    prepared = replace(snapshot.prepared, records=freeze_json(records))
    preserved = replace(
        snapshot,
        prepared=prepared,
        input_fingerprint=analytical_fingerprint(prepared, snapshot.methods),
    )
    path = root / "invalid-pages.json"
    save_result(preserved, path, result_id="invalid-pages")
    page.load_analysis(str(path))
    wait_until(application, lambda: page.review is not None and not page.has_workers())
    assert page._image is None
    assert "recorded page range is invalid" in page.canvas.text()
    assert page.open_source.isEnabled()
    assert not page.page.isEnabled()


def test_close_during_explicit_file_open_cancels_external_launch(
    window, application, pack, monkeypatch
):
    _, path, _, _, _ = pack
    page = open_pack(window, application, path)
    entered, release = Event(), Event()
    real_run = PdfPreview.run

    def held_run(worker):
        entered.set()
        assert release.wait(10)
        real_run(worker)

    monkeypatch.setattr(PdfPreview, "run", held_run)
    monkeypatch.setattr(
        qt_review.QDesktopServices,
        "openUrl",
        lambda *_: pytest.fail("Closing the window cancels a pending file launch"),
    )
    page.open_source.click()
    assert entered.wait(2)
    window.close()
    assert window.isVisible()
    release.set()
    wait_until(application, lambda: not page.has_workers() and not window.isVisible())


def test_source_inventory_is_separate_from_financial_records(window, application, pack):
    _, path, original, _, _ = pack
    page = open_pack(window, application, path)
    assert page.proxy.rowCount() == 3
    page.category.setCurrentText("Source files")
    wait_until(application, lambda: not page.has_workers())
    assert page.proxy.rowCount() == 1
    assert page._selected_source.path == str(original)
    assert page._image is not None
    assert page.page.value() == 1
    assert "No page range recorded" in page.source_status.text()
    assert "Content Hash" in page.details.toPlainText()
    assert page.proxy.data(page.proxy.index(0, 3)) == ""


def test_sap_only_result_exposes_its_original_pdf_without_inventing_bill_rows(
    window, application, pack
):
    root, _, original, snapshot, source_id = pack
    state = object_value(snapshot.prepared.extraction.legacy_state)
    state["filtered_records"] = []
    state["sap_contract_rows"] = [
        {"Account": "ACCOUNT-1", "Contract": "SAP-CONTRACT", "Source ID": source_id}
    ]
    sap_only = altered_snapshot(snapshot, state, records=[])
    path = root / "sap-result.json"
    save_result(sap_only, path, result_id="sap-only")
    page = open_pack(window, application, path)
    assert page.category.currentText() == "Source files"
    assert page.proxy.rowCount() == 1
    assert rows_value(page.review.snapshot.prepared.records) == []
    page.category.setCurrentText("Source files")
    wait_until(application, lambda: not page.has_workers())
    assert page.proxy.rowCount() == 1
    assert page._selected_source.path == str(original)
    assert page._image is not None
    assert page.review.snapshot.summary == sap_only.summary
