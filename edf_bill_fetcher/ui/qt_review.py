"""Searchable saved evidence beside a verified, paginated original PDF preview."""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path
from threading import Event

from PySide6.QtCore import (
    QAbstractTableModel,
    QModelIndex,
    QObject,
    QPersistentModelIndex,
    QSize,
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
    QMessageBox,
    QPlainTextEdit,
    QScrollArea,
    QSizePolicy,
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
from edf_bill_fetcher.application.review_session import ReviewSession
from edf_bill_fetcher.application.run_service import RunCancelled, check_cancelled
from edf_bill_fetcher.domain.analysis import AnalysisSnapshot
from edf_bill_fetcher.domain.json_values import (
    freeze_json,
    object_value,
    rows_value,
)
from edf_bill_fetcher.helpers.pdf_utils import PDF_LOCK
from edf_bill_fetcher.io.persistence.result_store import load_result
from edf_bill_fetcher.ui.qt_rebilling import RebillingPanel
from edf_bill_fetcher.ui.qt_review_tools import RecordEditor, RevisionPanel
from edf_bill_fetcher.ui.qt_widgets import button, label

CATEGORIES = (
    "All records",
    "Bills & charges",
    "Payments & credits",
    "Findings",
    "Duplicates",
    "Filtered out",
    "Excluded",
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
            if entry.category in {
                "Findings",
                "Duplicates",
                "Filtered out",
                "Source files",
                "Excluded",
            }:
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
        self.session: ReviewSession | None = None

    def run(self) -> None:
        """Keep the previously displayed analysis intact when a new load fails."""
        try:
            check_cancelled(self.cancel_event.is_set)
            snapshot = self.snapshot if self.snapshot is not None else load_result(self.path)
            draft_path = Path(self.path or str(Path(self.directory) / "result.json")).with_suffix(
                ".review-draft.json"
            )
            self.session = (
                ReviewSession.load(snapshot, draft_path)
                if draft_path.is_file()
                else ReviewSession(snapshot)
            )
            result = build_evidence_review(snapshot, self.directory, self.cancel_event.is_set)
            self.loaded.emit(result)
        except RunCancelled:
            return
        except Exception as error:
            self.failed.emit(str(error))


class ReviewPreview(QThread):
    """Calculate an explicit draft preview without writing files or changing accepted results."""

    ready = Signal(object)
    failed = Signal(str)

    def __init__(self, session: ReviewSession, source: str, parent: QObject | None = None):
        """Capture an immutable snapshot and detached validated draft."""
        super().__init__(parent)
        self.snapshot, self.payload, self.source = session.snapshot, session.payload(), source
        self.cancel_event = Event()

    def run(self) -> None:
        """Use the same preparation and analysis path as a reviewed export."""
        from datetime import datetime, timezone
        from uuid import uuid4

        from edf_bill_fetcher.application.analysis_service import analyse_prepared
        from edf_bill_fetcher.application.legacy_adapter import capture_legacy
        from edf_bill_fetcher.application.preparation_service import prepare_snapshot

        try:
            session = ReviewSession.restore(self.snapshot, self.payload)
            records, state, _ = session.inputs(self.source)
            check_cancelled(self.cancel_event.is_set)
            request = replace(
                self.snapshot.prepared.extraction.request,
                run_id=uuid4().hex,
                created_at=datetime.now(timezone.utc),
                config=freeze_json(state["review_config"]),
            )
            result = analyse_prepared(prepare_snapshot(capture_legacy(records, state, request)))
            check_cancelled(self.cancel_event.is_set)
            self.ready.emit(result)
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


class ReviewSplitter(QSplitter):
    """Expand normally without inheriting hidden editors' full form height."""

    def sizeHint(self) -> QSize:
        """Give the scroll layout a real height, bounded by the current review mode."""
        hint = super().sizeHint()
        hint.setHeight(self.minimumHeight())
        return hint


class EvidenceReviewPage(QWidget):
    """Browse records, findings and traceable sources without changing the result."""

    changed = Signal()
    idle = Signal()
    regenerate_requested = Signal(object)

    def __init__(self, parent: QWidget | None = None) -> None:
        """Build a responsive split view with keyboard-accessible record selection."""
        super().__init__(parent)
        self.setObjectName("page")
        self.review: EvidenceReview | None = None
        self.session: ReviewSession | None = None
        self.draft_path = ""
        self.draft_dirty = False
        self._saved_draft: dict | None = None
        self._guarding = False
        self._draft_generation = 0
        self.impact_worker: ReviewPreview | None = None
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
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        self.workspace_scroll = QScrollArea()
        self.workspace_scroll.setWidgetResizable(True)
        content = QWidget()
        layout = QVBoxLayout(content)
        self.workspace_scroll.setWidget(content)
        outer.addWidget(self.workspace_scroll)
        layout.setContentsMargins(24, 20, 24, 40)
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
        review_actions = QHBoxLayout()
        self.edit_record = button("Edit selected record", self._edit_selected)
        self.undo_button = button("Undo", lambda: self._history(False))
        self.redo_button = button("Redo", lambda: self._history(True))
        for action in (self.edit_record, self.undo_button, self.redo_button):
            review_actions.addWidget(action)
        layout.addLayout(review_actions)
        export_actions = QHBoxLayout()
        self.save_draft = button("Save draft", self.save_progress)
        self.preview_changes = button("Preview changes", self.preview_draft)
        self.regenerate = button("Create reviewed pack", self.regenerate_draft)
        for action in (self.save_draft, self.preview_changes, self.regenerate):
            export_actions.addWidget(action)
        layout.addLayout(export_actions)
        self.draft_status = label(
            "Edits are saved separately from the accepted analysis.", "hint", wrap=True
        )
        layout.addWidget(self.draft_status)
        self.impact = label("", "hint", wrap=True)
        self.impact.hide()
        layout.addWidget(self.impact)
        self.splitter = ReviewSplitter(Qt.Orientation.Horizontal)
        self.splitter.setMinimumHeight(420)
        self.splitter.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
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
        self.details_tabs.setMinimumHeight(210)
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
        self.editor = RecordEditor()
        self.revisions = RevisionPanel()
        self.rebilling = RebillingPanel()
        self.details_tabs.addTab(self.editor, "Correct record")
        self.details_tabs.addTab(self.revisions, "Bill revisions")
        self.details_tabs.addTab(self.rebilling, "Dispute ledger")
        self.editor.applied.connect(self._draft_changed)
        self.editor.failed.connect(self._message)
        self.revisions.changed.connect(self._draft_changed)
        self.revisions.failed.connect(self._message)
        self.revisions.show_record.connect(self.select_record)
        self.rebilling.changed.connect(self._draft_changed)
        self.rebilling.failed.connect(self._message)
        self.rebilling.show_record.connect(self.select_record)
        records.addWidget(self.details_tabs, 2)
        self.splitter.addWidget(left)
        right = QWidget()
        self.preview_pane = right
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
        self.preview_scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Ignored)
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
        preview.addWidget(self.open_source, 0, Qt.AlignmentFlag.AlignLeft)
        self.splitter.addWidget(right)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setSizes([540, 440])
        layout.addWidget(self.splitter, 1)
        self._draft_timer = QTimer(self)
        self._draft_timer.setSingleShot(True)
        self._draft_timer.setInterval(400)
        self._draft_timer.timeout.connect(self.save_progress)
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._focus_timer = QTimer(self)
        self._focus_timer.setSingleShot(True)
        self._focus_timer.timeout.connect(self._reveal_tools)
        self.details_tabs.currentChanged.connect(self._focus_tools)
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
        self._review_controls()

    def resizeEvent(self, event: QResizeEvent) -> None:
        """Stack source and records on smaller windows and refit the raster."""
        super().resizeEvent(event)
        self._arrange_panes()
        self._display_image()
        self._focus_timer.start(0)

    def _arrange_panes(self) -> None:
        """Respect native font/control widths, including Windows display scaling."""
        panes = (self.splitter.widget(0), self.preview_pane)
        assert panes[0] is not None
        minimum_width = (
            panes[0].minimumSizeHint().width()
            + panes[1].minimumSizeHint().width()
            + self.splitter.handleWidth()
        )
        self.splitter.setOrientation(
            Qt.Orientation.Vertical
            if self.workspace_scroll.viewport().width() < max(820, minimum_width + 48)
            else Qt.Orientation.Horizontal
        )

    def _focus_tools(self, index: int) -> None:
        editing = self.details_tabs.widget(index) in {self.editor, self.revisions, self.rebilling}
        self.table.setMaximumHeight(110 if editing else 16777215)
        self.details_tabs.setMinimumHeight(450 if editing else 210)
        self.splitter.setMinimumHeight(600 if editing else 420)
        if editing:
            self._focus_timer.start(0)

    def _reveal_tools(self) -> None:
        editing = self.details_tabs.currentWidget() in {self.editor, self.revisions, self.rebilling}
        self._arrange_panes()
        source_layout = self.preview_pane.layout()
        assert source_layout is not None
        needed = source_layout.totalHeightForWidth(max(1, self.preview_pane.width()))
        source_height = max(
            self.preview_pane.minimumSizeHint().height(),
            needed + self.preview_scroll.minimumHeight(),
        )
        if self.splitter.orientation() == Qt.Orientation.Vertical:
            records_pane = self.splitter.widget(0)
            assert records_pane is not None
            height = (
                max(600 if editing else 420, records_pane.minimumSizeHint().height())
                + source_height
                + self.splitter.handleWidth()
            )
        else:
            height = max(
                600 if editing else 420,
                source_height,
            )
        self.splitter.setMinimumHeight(height)
        if editing:
            self.workspace_scroll.ensureWidgetVisible(self.details_tabs)

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
        if (
            self.loader is not None
            or self.impact_worker is not None
            or self._shutdown
            or not self.allow_leave()
        ):
            return
        self.open_button.setEnabled(False)
        self._message("Opening analysis and verifying original files…")
        worker = ReviewLoader(directory or str(Path(path).parent), path, snapshot, self)
        self.loader = worker
        worker.loaded.connect(self._loaded)
        worker.failed.connect(self._load_failed)
        worker.finished.connect(self._load_finished)
        self._review_controls()
        worker.start()
        self.changed.emit()

    @Slot(object)
    def _loaded(self, review: EvidenceReview) -> None:
        if self._shutdown:
            return
        assert self.loader is not None
        self.analysis_path = self.loader.path or str(Path(review.directory) / "result.json")
        self.review = review
        self.session = self.loader.session
        self._saved_draft = self.session.payload() if self.session else None
        self.draft_path = str(Path(self.analysis_path).with_suffix(".review-draft.json"))
        self.draft_dirty = False
        self._draft_timer.stop()
        self.impact.clear()
        self.impact.hide()
        self.revisions.refresh(self.session)
        self.rebilling.refresh(self.session)
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
            f"Accepted analysis as of {review.snapshot.prepared.extraction.request.as_of.isoformat()} · Charges £{summary.charges:,.2f} · Net credits £{summary.net_credits:,.2f}"
        )
        self.summary.setToolTip(self.analysis_path)
        self.run_notes.setPlainText("\n\n".join(review.notes) or "No review notes were recorded.")
        self.details_tabs.setTabText(2, f"Review notes ({len(review.notes)})")
        self.message.hide()
        self.editor.set_record(None)
        self._filter()
        self._review_controls()
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
        self._review_controls()
        self.changed.emit()
        self._emit_idle()

    def _message(self, message: str) -> None:
        self.message.setText(message)
        self.message.show()

    def _filter(self) -> None:
        if not self._guarding and not self.editor.allow_leave():
            return
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
        if self._guarding:
            return
        if not self.editor.allow_leave():
            self._guarding = True
            self.select_record(self.editor.row_id)
            self._guarding = False
            return
        selected = self.table.selectionModel().selectedRows()
        self._clear_source()
        self.details.clear()
        self.extracted_text.clear()
        if not selected:
            self.editor.set_record(None)
            self._review_controls()
            return
        index = self.proxy.mapToSource(selected[0]).row()
        entry, row = self.model.entries[index], self.model.rows[index]
        self.editor.set_record(self.session if entry.row_id else None, entry.row_id)
        self._review_controls()
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
        self._focus_timer.start(0)
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

    def _review_controls(self) -> None:
        ready = self.session is not None and self.loader is None and self.impact_worker is None
        self.editor.setEnabled(ready and bool(self.editor.row_id))
        self.revisions.setEnabled(ready)
        self.rebilling.setEnabled(ready)
        self.edit_record.setEnabled(ready and bool(self.editor.row_id))
        self.undo_button.setEnabled(ready and bool(self.session and self.session.undo_stack))
        self.redo_button.setEnabled(ready and bool(self.session and self.session.redo_stack))
        for control in (self.save_draft, self.preview_changes, self.regenerate):
            control.setEnabled(ready)
        if self.session:
            self.draft_status.setText(
                f"Draft: {len(self.session.decisions)} record decisions · {len(self.session.revisions)} revision decisions · {len(self.session.schedule)} dispute positions. "
                + (
                    "Unsaved progress."
                    if self.draft_dirty
                    else "Accepted figures above update when a reviewed pack is created."
                )
            )
            self.draft_status.setToolTip(self.draft_path)
        self.changed.emit()

    def _edit_selected(self) -> None:
        if self.editor.row_id:
            self.details_tabs.setCurrentWidget(self.editor)
            self.workspace_scroll.ensureWidgetVisible(self.details_tabs)

    def _draft_changed(self) -> None:
        self.draft_dirty = True
        self._draft_generation += 1
        self.impact.setText(
            "Draft changed. Preview or create a reviewed pack to update calculated figures."
        )
        self.impact.show()
        self.revisions.refresh(self.session)
        self.rebilling.refresh(self.session)
        self._review_controls()
        self._draft_timer.start()

    def save_progress(self) -> bool:
        """Save applied draft decisions atomically, without rewriting the accepted analysis."""
        if self.session is None or not self.draft_dirty:
            return True
        try:
            self.session.save(self.draft_path)
        except (OSError, ValueError) as error:
            self._message("Draft could not be saved: " + str(error))
            return False
        self.draft_dirty = False
        self._saved_draft = self.session.payload()
        self._review_controls()
        return True

    def allow_leave(self) -> bool:
        """Preserve unapplied edits and failed draft saves when changing analyses or closing."""
        if (
            not self.editor.allow_leave()
            or not self.revisions.allow_leave()
            or not self.rebilling.allow_leave()
        ):
            return False
        if self.save_progress():
            return True
        choice = QMessageBox.question(
            self,
            "Draft save failed",
            "Your draft could not be saved. Discard the unsaved progress and continue? Choose Cancel to keep it open and resolve the save problem.",
            QMessageBox.StandardButton.Discard | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel,
        )
        if choice == QMessageBox.StandardButton.Discard:
            if self.session is not None and self._saved_draft is not None:
                row_id = self.editor.row_id
                self.session = ReviewSession.restore(self.session.snapshot, self._saved_draft)
                self.editor.set_record(self.session if row_id else None, row_id)
                self.revisions.refresh(self.session)
                self.rebilling.refresh(self.session)
            self.draft_dirty = False
            self._draft_timer.stop()
            return True
        return False

    def _history(self, redo: bool) -> None:
        if (
            self.session is None
            or not self.editor.allow_leave()
            or not self.revisions.allow_leave()
            or not self.rebilling.allow_leave()
        ):
            return
        row_id = self.editor.row_id
        (self.session.redo if redo else self.session.undo)()
        self.editor.set_record(self.session if row_id else None, row_id)
        self._draft_changed()

    def select_record(self, row_id: str) -> None:
        """Reveal a stable revision endpoint and display its verified original source."""
        for index, entry in enumerate(self.model.entries):
            if entry.row_id != row_id:
                continue
            target = self.proxy.mapFromSource(self.model.index(index, 0))
            if not target.isValid():
                self.search.clear()
                self.category.setCurrentText(entry.category)
                self.availability.setCurrentIndex(0)
                self._filter()
                target = self.proxy.mapFromSource(self.model.index(index, 0))
            if target.isValid():
                self.table.selectRow(target.row())
            return

    def preview_draft(self) -> None:
        """Show financial impact from an explicit, cancellable background recalculation."""
        if self.session is None or self.impact_worker is not None or not self.allow_leave():
            return
        worker = ReviewPreview(self.session, self.analysis_path, self)
        self.impact_worker = worker
        self._preview_generation = self._draft_generation
        self.open_button.setEnabled(False)
        self.editor.setEnabled(False)
        self.revisions.setEnabled(False)
        self.rebilling.setEnabled(False)
        self.impact.setText("Calculating the draft preview…")
        self.impact.show()
        worker.ready.connect(self._impact_ready)
        worker.failed.connect(self._impact_failed)
        worker.finished.connect(self._impact_finished)
        self._review_controls()
        worker.start()

    def _impact_failed(self, error: str) -> None:
        self.impact.setText("Preview could not be calculated. Accepted figures are unchanged.")
        self._message(error)

    def _impact_ready(self, snapshot: AnalysisSnapshot) -> None:
        if (
            self._shutdown
            or self._preview_generation != self._draft_generation
            or self.review is None
        ):
            return
        before, after = self.review.snapshot.summary, snapshot.summary
        self.impact.setText(
            f"Preview only · Charges £{before.charges:,.2f} → £{after.charges:,.2f} · Net credits £{before.net_credits:,.2f} → £{after.net_credits:,.2f}. Create a reviewed pack to publish these figures."
        )

    def _impact_finished(self) -> None:
        worker, self.impact_worker = self.impact_worker, None
        if worker:
            worker.deleteLater()
        self.open_button.setEnabled(True)
        self.editor.setEnabled(bool(self.editor.row_id))
        self.revisions.setEnabled(self.session is not None)
        self.rebilling.setEnabled(self.session is not None)
        self._review_controls()
        self._emit_idle()

    def regenerate_draft(self) -> None:
        """Ask the owning window to publish a new reviewed pack using its output preferences."""
        if self.session is not None and self.allow_leave():
            self.regenerate_requested.emit(self.session.payload())

    def has_workers(self) -> bool:
        """Keep owners alive until all native handles and worker signals are finished."""
        return (
            self.loader is not None
            or self.preview_worker is not None
            or self.opener is not None
            or self.impact_worker is not None
        )

    def shutdown(self) -> None:
        """Cancel loading/rendering and discard queued requests on application close."""
        self._shutdown = True
        self._pending_preview = None
        self._search_timer.stop()
        self._draft_timer.stop()
        self._focus_timer.stop()
        self.save_progress()
        for worker in (self.loader, self.preview_worker, self.opener, self.impact_worker):
            if worker is not None:
                worker.cancel_event.set()

    def _emit_idle(self) -> None:
        if not self.has_workers():
            self.idle.emit()
