"""Searchable saved evidence beside a verified, paginated original PDF preview."""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from threading import Event

from PySide6.QtCore import (
    QAbstractTableModel,
    QModelIndex,
    QObject,
    QPersistentModelIndex,
    QSortFilterProxyModel,
    Qt,
    QThread,
    QTimer,
    QUrl,
    Signal,
    Slot,
)
from PySide6.QtGui import QDesktopServices, QImage, QPixmap, QResizeEvent
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QScrollArea,
    QSpinBox,
    QSplitter,
    QTableView,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from edf_bill_fetcher.application.evidence_review import (
    EvidenceReview,
    ReviewEntry,
    ReviewSource,
    build_evidence_review,
    text,
)
from edf_bill_fetcher.application.run_service import RunCancelled, check_cancelled
from edf_bill_fetcher.domain.analysis import AnalysisSnapshot
from edf_bill_fetcher.domain.json_values import object_value, rows_value
from edf_bill_fetcher.helpers.pdf_utils import PDF_LOCK
from edf_bill_fetcher.io.persistence.result_store import load_result
from edf_bill_fetcher.ui.qt_widgets import button, label

CATEGORIES = (
    "All records",
    "Bills & charges",
    "Payments & credits",
    "Findings",
    "Duplicates",
    "Filtered out",
    "Other records",
    "Source files",
)


def row_value(row: dict) -> tuple[object, str]:
    """Describe the preserved money field without changing its sign or basis."""
    for key, basis in (
        ("Transaction Amount (£)", "Transaction"),
        ("Period Charge (£)", "Period charge"),
        ("Balance (£)", "Balance"),
        (
            "Amount (£)",
            "Finding amount"
            if "Finding" in row
            else text(row.get("Amount Kind")) or "Recorded amount",
        ),
    ):
        if row.get(key) is not None:
            return row[key], basis
    return None, "Unavailable"


class EvidenceTable(QAbstractTableModel):
    """Detached rows: the table cannot edit saved records or calculations."""

    headers = ("Date", "Record / finding", "Invoice", "Value (£)", "Basis", "Source")
    sort_role = Qt.ItemDataRole.UserRole + 1

    def __init__(
        self, entries: tuple[ReviewEntry, ...] = (), parent: QObject | None = None
    ) -> None:
        """Cache presentation fields once rather than decoding on every repaint."""
        super().__init__(parent)
        self.entries = entries
        self.rows = [object_value(entry.fields) for entry in entries]
        self.searches = [
            " ".join(
                (entry.category, text(row), " ".join(source.title for source in entry.sources))
            ).casefold()
            for row, entry in zip(self.rows, entries, strict=True)
        ]

    def rowCount(self, parent: QModelIndex | QPersistentModelIndex = QModelIndex()) -> int:
        """Return flat row membership."""
        return 0 if parent.isValid() else len(self.entries)

    def columnCount(self, parent: QModelIndex | QPersistentModelIndex = QModelIndex()) -> int:
        """Return the fixed review columns."""
        return 0 if parent.isValid() else len(self.headers)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        """Label columns with the meaning of their displayed values."""
        if role == Qt.ItemDataRole.DisplayRole and orientation == Qt.Orientation.Horizontal:
            return self.headers[section]
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        """Supply literal display text and numeric/date-aware sorting."""
        if not index.isValid() or not 0 <= index.row() < len(self.rows):
            return None
        row, entry = self.rows[index.row()], self.entries[index.row()]
        amount, basis = row_value(row)
        values = (
            text(row.get("Date")),
            " · ".join(text(row.get(key)) for key in ("Severity", "Finding") if row.get(key))
            if row.get("Finding")
            else text(row.get("Entry Type")) or entry.category,
            text(row.get("Invoice #")),
            text(amount),
            basis,
            text(row.get("Original Filename") or row.get("Attachment Name"))
            or ("Related records" if entry.sources else "Unavailable"),
        )
        if role == Qt.ItemDataRole.DisplayRole:
            if (
                index.column() == 3
                and isinstance(amount, (float, int))
                and not isinstance(amount, bool)
            ):
                return f"{amount:,.2f}"
            return values[index.column()]
        if role == self.sort_role:
            if index.column() == 3:
                try:
                    return float(text(amount)) if amount is not None else float("-inf")
                except (TypeError, ValueError, OverflowError):
                    return float("-inf")
            if index.column() == 0:
                for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d %B %Y"):
                    try:
                        return datetime.strptime(values[0], fmt).date().isoformat()
                    except ValueError:
                        pass
            return values[index.column()].casefold()
        if role == Qt.ItemDataRole.ToolTipRole:
            return values[index.column()]
        return None


class EvidenceFilter(QSortFilterProxyModel):
    """Literal search with record categories and evidence availability filters."""

    def __init__(self, parent: QObject | None = None) -> None:
        """Configure sorting independently of the saved row order."""
        super().__init__(parent)
        self.query = ""
        self.category = "All records"
        self.availability = 0
        self.setSortRole(EvidenceTable.sort_role)

    def filterAcceptsRow(self, row, parent):
        """Treat search characters as text, including brackets and ampersands."""
        model = self.sourceModel()
        if not isinstance(model, EvidenceTable):
            return False
        entry = model.entries[row]
        if self.category == "All records":
            if entry.category in {"Findings", "Duplicates", "Filtered out", "Source files"}:
                return False
        elif self.category != entry.category:
            return False
        available = any(source.path for source in entry.sources)
        if (self.availability == 1 and not available) or (self.availability == 2 and available):
            return False
        return self.query in model.searches[row]


class ReviewLoader(QThread):
    """Strict saved-result loading and file verification away from the GUI thread."""

    loaded = Signal(object)
    failed = Signal(str)

    def __init__(
        self,
        directory: str,
        path: str = "",
        snapshot: AnalysisSnapshot | None = None,
        parent: QObject | None = None,
    ) -> None:
        """Capture immutable input and cooperative cancellation state."""
        super().__init__(parent)
        self.directory, self.path, self.snapshot = directory, path, snapshot
        self.cancel_event = Event()

    def run(self) -> None:
        """Keep the previously displayed analysis intact when a new load fails."""
        try:
            check_cancelled(self.cancel_event.is_set)
            snapshot = self.snapshot if self.snapshot is not None else load_result(self.path)
            result = build_evidence_review(snapshot, self.directory, self.cancel_event.is_set)
            self.loaded.emit(result)
        except RunCancelled:
            return
        except Exception as error:
            self.failed.emit(str(error))


@dataclass(frozen=True)
class PreviewRequest:
    """The source and one-based page selected before starting background rendering."""

    source: ReviewSource
    page: int
    scale: float
    open_only: bool = False


class PdfPreview(QThread):
    """Render one verified original page and close every native PDF handle."""

    rendered = Signal(object, int)
    verified = Signal(str)
    failed = Signal(str)

    def __init__(self, request: PreviewRequest, parent: QObject | None = None) -> None:
        """Accept a frozen request; never read widgets in the worker."""
        super().__init__(parent)
        self.request = request
        self.cancel_event = Event()

    def run(self) -> None:
        """Recheck the bytes actually displayed, including changes since pack opening."""
        try:
            check_cancelled(self.cancel_event.is_set)
            source = self.request.source
            if not source.path:
                raise ValueError("The original source is unavailable.")
            data = Path(source.path).read_bytes()
            check_cancelled(self.cancel_event.is_set)
            if source.content_hash and hashlib.sha256(data).hexdigest() != source.content_hash:
                raise ValueError(
                    "This original source has changed since extraction. Preview withheld."
                )
            check_cancelled(self.cancel_event.is_set)
            if self.request.open_only:
                self.verified.emit(source.path)
                return
            import pypdfium2 as pdfium

            with PDF_LOCK:
                check_cancelled(self.cancel_event.is_set)
                document = pdfium.PdfDocument(data)
                try:
                    count = len(document)
                    if not 1 <= self.request.page <= count or (
                        source.pages and source.pages[1] > count
                    ):
                        raise ValueError(
                            "The saved page range is outside this PDF. Preview withheld because the source link is unreliable."
                        )
                    page = document[self.request.page - 1]
                    try:
                        width, height = page.get_size()
                        # Bound raster memory for unusually large PDF page sizes.
                        scale = min(self.request.scale, 4000 / max(width, height))
                        bitmap = page.render(scale=scale)
                        try:
                            image = bitmap.to_pil().convert("RGB")
                            pixels = image.tobytes()
                            result = QImage(
                                pixels,
                                image.width,
                                image.height,
                                image.width * 3,
                                QImage.Format.Format_RGB888,
                            ).copy()
                        finally:
                            bitmap.close()
                    finally:
                        page.close()
                finally:
                    document.close()
            check_cancelled(self.cancel_event.is_set)
            self.rendered.emit(result, count)
        except RunCancelled:
            return
        except Exception as error:
            self.failed.emit(str(error))


class EvidenceReviewPage(QWidget):
    """Browse records, findings and traceable sources without changing the result."""

    changed = Signal()
    idle = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        """Build a responsive split view with keyboard-accessible record selection."""
        super().__init__(parent)
        self.setObjectName("page")
        self.review: EvidenceReview | None = None
        self.analysis_path = ""
        self.loader: ReviewLoader | None = None
        self.preview_worker: PdfPreview | None = None
        self.opener: PdfPreview | None = None
        self._pending_preview: PreviewRequest | None = None
        self._active_preview: PreviewRequest | None = None
        self._selected_source: ReviewSource | None = None
        self._image: QImage | None = None
        self._open_failed = False
        self._shutdown = False
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 16)
        layout.setSpacing(10)
        header = QHBoxLayout()
        header.addWidget(label("Review evidence", "title"), 1)
        self.open_button = button("Open analysis…", self.pick_analysis)
        header.addWidget(self.open_button)
        layout.addLayout(header)
        self.summary = label(
            "Open result.json from an evidence pack, or create a new pack to review it here.",
            "hint",
            wrap=True,
        )
        layout.addWidget(self.summary)
        self.message = label("", "feedback", wrap=True)
        self.message.hide()
        layout.addWidget(self.message)
        filters = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search invoice, account, text or filename…")
        self.search.setClearButtonEnabled(True)
        self.search.setAccessibleName("Search saved evidence")
        filters.addWidget(self.search, 1)
        self.category = QComboBox()
        self.category.addItems(CATEGORIES)
        self.category.setAccessibleName("Evidence category")
        filters.addWidget(self.category)
        self.availability = QComboBox()
        self.availability.addItems(("All sources", "Available sources", "Unavailable sources"))
        self.availability.setAccessibleName("Source availability filter")
        filters.addWidget(self.availability)
        layout.addLayout(filters)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        left = QWidget()
        records = QVBoxLayout(left)
        records.setContentsMargins(0, 0, 0, 0)
        self.count = label("No analysis open", "hint")
        records.addWidget(self.count)
        self.model = EvidenceTable(parent=self)
        self.proxy = EvidenceFilter(self)
        self.proxy.setSourceModel(self.model)
        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setAccessibleName("Saved evidence records")
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.sortByColumn(0, Qt.SortOrder.AscendingOrder)
        self.table.verticalHeader().hide()
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        for column, width in enumerate((95, 140, 105, 85, 105, 130)):
            self.table.setColumnWidth(column, width)
        records.addWidget(self.table, 3)
        self.details_tabs = QTabWidget()
        self.details = QPlainTextEdit()
        self.details.setReadOnly(True)
        self.details.setPlaceholderText("Select a record to inspect all preserved fields.")
        self.details.setAccessibleName("Record details")
        self.extracted_text = QPlainTextEdit()
        self.extracted_text.setReadOnly(True)
        self.extracted_text.setAccessibleName("Original extracted PDF text")
        self.details_tabs.addTab(self.details, "Record details")
        self.details_tabs.addTab(self.extracted_text, "Extracted text")
        self.run_notes = QPlainTextEdit()
        self.run_notes.setReadOnly(True)
        self.run_notes.setAccessibleName("Saved analysis review notes")
        self.details_tabs.addTab(self.run_notes, "Review notes")
        records.addWidget(self.details_tabs, 2)
        self.splitter.addWidget(left)
        right = QWidget()
        preview = QVBoxLayout(right)
        preview.setContentsMargins(0, 0, 0, 0)
        preview.addWidget(label("Original source", "sectionTitle"))
        self.source_choices = QComboBox()
        self.source_choices.setAccessibleName("Original source occurrence")
        self.source_choices.setMinimumContentsLength(12)
        self.source_choices.setSizeAdjustPolicy(
            QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon
        )
        preview.addWidget(self.source_choices)
        self.source_status = label("Select a record to inspect its source.", "hint", wrap=True)
        self.source_status.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        preview.addWidget(self.source_status)
        self.canvas = QLabel("PDF originals appear here.")
        self.canvas.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.canvas.setMinimumSize(1, 1)
        self.canvas.setTextFormat(Qt.TextFormat.PlainText)
        self.canvas.setWordWrap(True)
        self.canvas.setObjectName("pdfCanvas")
        self.preview_scroll = QScrollArea()
        self.preview_scroll.setWidget(self.canvas)
        self.preview_scroll.setWidgetResizable(True)
        self.preview_scroll.setMinimumSize(180, 100)
        preview.addWidget(self.preview_scroll, 1)
        controls = QHBoxLayout()
        self.previous = button("‹", lambda: self.page.setValue(self.page.value() - 1))
        self.previous.setAccessibleName("Previous PDF page")
        controls.addWidget(self.previous)
        self.page = QSpinBox()
        self.page.setRange(1, 1)
        self.page.setPrefix("Page ")
        self.page.setAccessibleName("PDF page number")
        controls.addWidget(self.page)
        self.next = button("›", lambda: self.page.setValue(self.page.value() + 1))
        self.next.setAccessibleName("Next PDF page")
        controls.addWidget(self.next)
        self.zoom = QComboBox()
        self.zoom.addItems(("Fit width", "100%", "150%", "200%"))
        self.zoom.setAccessibleName("PDF preview zoom")
        controls.addWidget(self.zoom, 1)
        preview.addLayout(controls)
        self.open_source = button("Open original file", self._open_source)
        preview.addWidget(self.open_source)
        self.splitter.addWidget(right)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setSizes([540, 440])
        layout.addWidget(self.splitter, 1)
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(150)
        self._search_timer.timeout.connect(self._filter)
        self.search.textChanged.connect(lambda: self._search_timer.start())
        self.category.currentTextChanged.connect(self._filter)
        self.availability.currentIndexChanged.connect(self._filter)
        self.table.selectionModel().selectionChanged.connect(self._selection_changed)
        self.source_choices.currentIndexChanged.connect(self._source_changed)
        self.page.valueChanged.connect(self._queue_preview)
        self.zoom.currentIndexChanged.connect(self._queue_preview)
        self._preview_controls(False)

    def resizeEvent(self, event: QResizeEvent) -> None:
        """Stack source and records on smaller windows and refit the raster."""
        super().resizeEvent(event)
        self.splitter.setOrientation(
            Qt.Orientation.Vertical
            if self.width() < 820 and self.height() >= 700
            else Qt.Orientation.Horizontal
        )
        self._display_image()

    def pick_analysis(self) -> None:
        """Open an existing saved result; browsing never publishes a new analysis."""
        if self.loader is not None:
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Open saved evidence analysis", "", "Saved analysis (*.json)"
        )
        if path:
            self.load_analysis(path)

    def load_analysis(
        self, path: str = "", *, snapshot: AnalysisSnapshot | None = None, directory: str = ""
    ) -> None:
        """Verify sources in the background and commit the view only after success."""
        if self.loader is not None or self._shutdown:
            return
        self.open_button.setEnabled(False)
        self._message("Opening analysis and verifying original files…")
        worker = ReviewLoader(directory or str(Path(path).parent), path, snapshot, self)
        self.loader = worker
        worker.loaded.connect(self._loaded)
        worker.failed.connect(self._load_failed)
        worker.finished.connect(self._load_finished)
        worker.start()
        self.changed.emit()

    @Slot(object)
    def _loaded(self, review: EvidenceReview) -> None:
        if self._shutdown:
            return
        assert self.loader is not None
        self.analysis_path = self.loader.path or str(Path(review.directory) / "result.json")
        self.review = review
        self.model.beginResetModel()
        self.model.entries = review.entries
        self.model.rows = [object_value(entry.fields) for entry in review.entries]
        self.model.searches = [
            " ".join(
                (entry.category, text(row), " ".join(source.title for source in entry.sources))
            ).casefold()
            for row, entry in zip(self.model.rows, review.entries, strict=True)
        ]
        self.model.endResetModel()
        self._search_timer.stop()
        for control in (self.search, self.category, self.availability):
            control.blockSignals(True)
        self.search.clear()
        self.category.setCurrentText(
            "Source files"
            if not rows_value(review.snapshot.prepared.records)
            and any(entry.category == "Source files" for entry in review.entries)
            else "All records"
        )
        self.availability.setCurrentIndex(0)
        for control in (self.search, self.category, self.availability):
            control.blockSignals(False)
        summary = review.snapshot.summary
        self.summary.setText(
            f"Analysis as of {review.snapshot.prepared.extraction.request.as_of.isoformat()} · Charges £{summary.charges:,.2f} · Net credits £{summary.net_credits:,.2f}\n{self.analysis_path}"
        )
        self.summary.setToolTip(self.analysis_path)
        self.run_notes.setPlainText("\n\n".join(review.notes) or "No review notes were recorded.")
        self.details_tabs.setTabText(2, f"Review notes ({len(review.notes)})")
        self.message.hide()
        self._filter()
        self.changed.emit()

    @Slot(str)
    def _load_failed(self, error: str) -> None:
        if not self._shutdown:
            self._message("Analysis could not be opened: " + error)

    @Slot()
    def _load_finished(self) -> None:
        worker, self.loader = self.loader, None
        self.open_button.setEnabled(True)
        if worker is not None:
            worker.deleteLater()
        self.changed.emit()
        self._emit_idle()

    def _message(self, message: str) -> None:
        self.message.setText(message)
        self.message.show()

    def _filter(self) -> None:
        self.proxy.query = self.search.text().strip().casefold()
        self.proxy.category = self.category.currentText()
        self.proxy.availability = self.availability.currentIndex()
        self.proxy.invalidate()
        count = self.proxy.rowCount()
        self.count.setText(
            f"{count} matching {'row' if count == 1 else 'rows'}"
            + (" · No matching evidence" if not count and self.review else "")
        )
        self.table.clearSelection()
        self._clear_source()
        self.details.clear()
        self.extracted_text.clear()
        if count:
            self.table.selectRow(0)

    def _selection_changed(self) -> None:
        selected = self.table.selectionModel().selectedRows()
        self._clear_source()
        self.details.clear()
        self.extracted_text.clear()
        if not selected:
            return
        index = self.proxy.mapToSource(selected[0]).row()
        entry, row = self.model.entries[index], self.model.rows[index]
        primary_fields = (
            "Finding",
            "Severity",
            "Details",
            "Invoice #",
            "Date",
            "Account",
            "Fuel",
            "Entry Type",
            "Amount Kind",
            "Amount (£)",
            "Balance (£)",
            "Period Charge (£)",
            "Transaction Amount (£)",
            "Period From",
            "Period To",
            "Reading",
            "Original Filename",
            "Attachment Name",
            "Page From",
            "Page To",
            "Text Extraction",
            "OCR Pages",
            "OCR Minimum Confidence (%)",
            "OCR Review Required",
        )
        ordered_keys = list(dict.fromkeys((*primary_fields, *row)))
        self.details.setPlainText(
            "\n".join(
                f"{key}: {text(row[key]) or 'Unavailable'}"
                for key in ordered_keys
                if key in row and key != "Source PDF Text"
            )
            + ("\n\n" + entry.relation_note if entry.relation_note else "")
        )
        self.extracted_text.setPlainText(
            text(row.get("Source PDF Text"))
            or "No extracted PDF text was preserved for this record."
        )
        self.source_choices.blockSignals(True)
        self.source_choices.clear()
        if entry.category == "Findings":
            self.source_choices.addItem("Choose a related record (date match)", None)
        for source in entry.sources:
            self.source_choices.addItem(source.title, source)
        self.source_choices.blockSignals(False)
        self.source_choices.setEnabled(bool(entry.sources))
        if entry.category == "Findings":
            self.source_status.setText(entry.relation_note)
            self.canvas.setText(
                "Choose a related record to inspect its source.\nA date match alone does not establish the finding’s source."
            )
        elif entry.sources:
            self._source_changed()
        else:
            self.source_status.setText("No original source mapping was preserved for this record.")

    def _clear_source(self) -> None:
        self._pending_preview = None
        self._selected_source = None
        if self.preview_worker is not None:
            self.preview_worker.cancel_event.set()
        if self.opener is not None:
            self.opener.cancel_event.set()
        self._open_failed = False
        self._image = None
        self.canvas.setPixmap(QPixmap())
        self.canvas.setMinimumSize(1, 1)
        self.canvas.setText("PDF originals appear here.")
        self.source_status.setText("Select a record to inspect its source.")
        self.source_status.setToolTip("")
        self.source_choices.blockSignals(True)
        self.source_choices.clear()
        self.source_choices.blockSignals(False)
        self.source_choices.setEnabled(False)
        self._preview_controls(False)

    def _preview_controls(self, enabled: bool) -> None:
        for control in (self.previous, self.page, self.next, self.zoom):
            control.setEnabled(enabled)
        self.open_source.setEnabled(
            bool(
                self._selected_source
                and self._selected_source.path
                and self.opener is None
                and not self._open_failed
            )
        )

    def _source_changed(self) -> None:
        source = self.source_choices.currentData()
        self._pending_preview = None
        self._selected_source = source if isinstance(source, ReviewSource) else None
        if self.preview_worker is not None:
            self.preview_worker.cancel_event.set()
        if self.opener is not None:
            self.opener.cancel_event.set()
        self._open_failed = False
        self._image = None
        self.canvas.setPixmap(QPixmap())
        self.canvas.setMinimumSize(1, 1)
        self._preview_controls(False)
        if self._selected_source is None:
            self.canvas.setText("Choose a related record to inspect its source.")
            return
        source = self._selected_source
        pages = (
            " · Recorded page range is invalid"
            if source.invalid_pages
            else f" · Recorded pages {source.pages[0]}–{source.pages[1]}"
            if source.pages
            else " · No page range recorded"
        )
        self.source_status.setText(
            source.status
            + pages
            + ("\nSaved file: " + Path(source.path).name if source.path else "")
        )
        self.source_status.setToolTip(source.path or "")
        if not source.path:
            self.canvas.setText(
                "This source is missing, changed or has no reliable mapping.\nThe saved record remains available on the left."
            )
        elif source.invalid_pages:
            self.canvas.setText(
                "The recorded page range is invalid.\nUse Open original file to inspect the document manually."
            )
        elif Path(source.path).suffix.lower() != ".pdf":
            self.canvas.setText(
                "Preview is available for PDF originals.\nUse Open original file to inspect this source."
            )
        else:
            self.page.blockSignals(True)
            self.page.setRange(1, max(1, source.pages[1] if source.pages else 1))
            self.page.setValue(source.pages[0] if source.pages else 1)
            self.page.blockSignals(False)
            self._queue_preview()

    def _queue_preview(self) -> None:
        source = self._selected_source
        if (
            source is None
            or not source.path
            or Path(source.path).suffix.lower() != ".pdf"
            or self._shutdown
        ):
            return
        self._image = None
        self.canvas.setPixmap(QPixmap())
        self.canvas.setMinimumSize(1, 1)
        self.canvas.setText("Rendering original page…")
        self._pending_preview = PreviewRequest(
            source, self.page.value(), (1.5, 1.0, 1.5, 2.0)[self.zoom.currentIndex()]
        )
        if self.preview_worker is not None:
            self.preview_worker.cancel_event.set()
        else:
            self._start_preview()

    def _start_preview(self) -> None:
        request, self._pending_preview = self._pending_preview, None
        if request is None or self._shutdown:
            self._emit_idle()
            return
        self._active_preview = request
        worker = PdfPreview(request, self)
        self.preview_worker = worker
        worker.rendered.connect(self._rendered)
        worker.failed.connect(self._preview_failed)
        worker.finished.connect(self._preview_finished)
        worker.start()

    def _current_preview(self) -> bool:
        return bool(
            not self._shutdown
            and self.preview_worker is not None
            and not self.preview_worker.cancel_event.is_set()
            and self._active_preview is not None
            and self._active_preview.source == self._selected_source
            and self._pending_preview is None
        )

    @Slot(object, int)
    def _rendered(self, image: QImage, count: int) -> None:
        if not self._current_preview():
            return
        self._image = image
        self.page.blockSignals(True)
        self.page.setRange(1, count)
        self.page.blockSignals(False)
        self.page.setSuffix(f" / {count}")
        self._preview_controls(True)
        self.previous.setEnabled(self.page.value() > 1)
        self.next.setEnabled(self.page.value() < count)
        self._display_image()

    def _display_image(self) -> None:
        if self._image is None:
            return
        pixmap = QPixmap.fromImage(self._image)
        if self.zoom.currentIndex() == 0:
            pixmap = pixmap.scaledToWidth(
                max(100, self.preview_scroll.viewport().width() - 12),
                Qt.TransformationMode.SmoothTransformation,
            )
        self.canvas.setMinimumSize(pixmap.size())
        self.canvas.setPixmap(pixmap)

    @Slot(str)
    def _preview_failed(self, error: str) -> None:
        if self._current_preview():
            self.canvas.setText("Preview could not be displayed.\n" + error)
            self._open_failed = True
            self.open_source.setEnabled(False)

    @Slot()
    def _preview_finished(self) -> None:
        worker, self.preview_worker = self.preview_worker, None
        if worker is not None:
            worker.deleteLater()
        self._start_preview()

    def _open_source(self) -> None:
        source = self._selected_source
        if source is None or not source.path or self.opener is not None or self._shutdown:
            return
        worker = PdfPreview(PreviewRequest(source, 1, 1, open_only=True), self)
        self.opener = worker
        self.open_source.setEnabled(False)
        worker.verified.connect(self._open_verified)
        worker.failed.connect(self._open_failed_message)
        worker.finished.connect(self._open_finished)
        worker.start()

    def _current_opener(self) -> bool:
        return bool(
            self.opener is not None
            and not self._shutdown
            and not self.opener.cancel_event.is_set()
            and self.opener.request.source == self._selected_source
        )

    @Slot(str)
    def _open_verified(self, path: str) -> None:
        if self._current_opener() and not QDesktopServices.openUrl(QUrl.fromLocalFile(path)):
            self._message("The desktop could not open this file.")

    @Slot(str)
    def _open_failed_message(self, error: str) -> None:
        if self._current_opener():
            self._open_failed = True
            self._message(error)

    @Slot()
    def _open_finished(self) -> None:
        worker, self.opener = self.opener, None
        if worker is not None:
            worker.deleteLater()
        self.open_source.setEnabled(
            bool(
                self._selected_source
                and self._selected_source.path
                and not self._open_failed
                and not self._shutdown
            )
        )
        self._emit_idle()

    def has_workers(self) -> bool:
        """Keep owners alive until all native handles and worker signals are finished."""
        return self.loader is not None or self.preview_worker is not None or self.opener is not None

    def shutdown(self) -> None:
        """Cancel loading/rendering and discard queued requests on application close."""
        self._shutdown = True
        self._pending_preview = None
        self._search_timer.stop()
        for worker in (self.loader, self.preview_worker, self.opener):
            if worker is not None:
                worker.cancel_event.set()

    def _emit_idle(self) -> None:
        if not self.has_workers():
            self.idle.emit()
