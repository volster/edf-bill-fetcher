"""Record correction and explicit bill revision tools beside original evidence."""

from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from edf_bill_fetcher.application.evidence_review import text
from edf_bill_fetcher.application.record_review import ENTRY_TYPES, REVIEW_FIELDS
from edf_bill_fetcher.application.review_session import ReviewSession
from edf_bill_fetcher.processors.revisions import BILLS, COMPARE_FIELDS, NOTICE
from edf_bill_fetcher.ui.qt_widgets import button, label

STATUS_LABELS = (
    ("Needs review", "unreviewed"),
    ("Verified against source", "verified"),
    ("Needs information", "needs_information"),
)


class RecordEditor(QWidget):
    """Edit validated fields while the original PDF stays visible in the adjacent pane."""

    applied = Signal()
    failed = Signal(str)

    def __init__(self, parent=None):
        """Build accessible controls without changing accepted evidence."""
        super().__init__(parent)
        self.session: ReviewSession | None = None
        self.row_id = ""
        self.initial = {}
        layout = QVBoxLayout(self)
        layout.addWidget(
            label(
                "Edit observations here; source text, files and hashes stay original. Apply to the draft, then preview or create a reviewed pack.",
                "hint",
                wrap=True,
            )
        )
        area = QScrollArea()
        area.setWidgetResizable(True)
        content = QWidget()
        form = QFormLayout(content)
        form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapAllRows)
        self.fields = {}
        for field in REVIEW_FIELDS:
            control: QComboBox | QPlainTextEdit | QLineEdit
            if field in {"Entry Type", "Amount Kind", "Fuel", "Cancel/Rebill Admitted"}:
                control = QComboBox()
                options = (
                    ENTRY_TYPES
                    if field == "Entry Type"
                    else ("Balance", "Transaction", "Unknown")
                    if field == "Amount Kind"
                    else ("", "Electricity", "Gas", "Dual", "Unknown")
                    if field == "Fuel"
                    else ("No", "Yes")
                )
                control.addItems(options)
                control.setEditable(field == "Fuel")
            elif field == "Details":
                control = QPlainTextEdit()
                control.setFixedHeight(65)
            else:
                control = QLineEdit()
                if REVIEW_FIELDS[field] in {"money", "number"}:
                    control.setPlaceholderText(
                        "Blank = unknown; enter a number without £ or commas"
                    )
                elif REVIEW_FIELDS[field] == "date":
                    control.setPlaceholderText("DD/MM/YYYY or YYYY-MM-DD; blank = unknown")
            control.setAccessibleName("Correct " + field)
            control.setMinimumWidth(0)
            self.fields[field] = control
            form.addRow(field, control)
        self.include = QCheckBox("Include this record in the analysis and bundle")
        form.addRow(self.include)
        self.status = QComboBox()
        for title, value in STATUS_LABELS:
            self.status.addItem(title, value)
        form.addRow("Review state", self.status)
        self.reason = QPlainTextEdit()
        self.reason.setAccessibleName("Reason for review changes")
        self.reason.setPlaceholderText("What did you verify or change, and why?")
        self.reason.setFixedHeight(65)
        form.addRow("Reason", self.reason)
        area.setWidget(content)
        layout.addWidget(area, 1)
        actions = QHBoxLayout()
        self.apply_button = button("Apply to draft", self.apply)
        actions.addWidget(self.apply_button)
        actions.addWidget(button("Reset form", self.reset))
        layout.addLayout(actions)
        self.setEnabled(False)

    @staticmethod
    def _value(control):
        if isinstance(control, QPlainTextEdit):
            return control.toPlainText()
        if isinstance(control, QComboBox):
            return control.currentText()
        return control.text()

    def values(self) -> dict:
        """Return literal form values for shared correction validation."""
        return {name: self._value(control) for name, control in self.fields.items()}

    def set_record(self, session: ReviewSession | None, row_id: str = "") -> None:
        """Load the effective draft values while preserving the original source."""
        self.session, self.row_id = session, row_id
        self.setEnabled(bool(session and row_id))
        if session is None or not row_id:
            self.initial = {}
            return
        record = session.record(row_id)
        for name, control in self.fields.items():
            value = record.get(name)
            rendered = (
                ("Yes" if value else "No") if name == "Cancel/Rebill Admitted" else text(value)
            )
            if isinstance(control, QPlainTextEdit):
                control.setPlainText(rendered)
            elif isinstance(control, QComboBox):
                if control.findText(rendered) < 0:
                    control.addItem(rendered)
                control.setCurrentText(rendered)
            else:
                control.setText(rendered)
        decision = session.decision(row_id)
        annotation = session.statuses.get(row_id, {"status": "unreviewed", "reason": ""})
        self.include.setChecked(decision["include"])
        self.status.setCurrentIndex(self.status.findData(annotation["status"]))
        self.reason.setPlainText(decision["reason"] or annotation["reason"])
        self.initial = {
            "values": self.values(),
            "include": self.include.isChecked(),
            "status": self.status.currentData(),
            "reason": self.reason.toPlainText(),
        }

    def dirty(self) -> bool:
        """Detect unapplied changes before navigation or regeneration."""
        return bool(
            self.row_id
            and self.initial
            and self.initial
            != {
                "values": self.values(),
                "include": self.include.isChecked(),
                "status": self.status.currentData(),
                "reason": self.reason.toPlainText(),
            }
        )

    def apply(self) -> bool:
        """Validate and apply the edited fields as one undoable command."""
        if self.session is None or not self.row_id:
            return False
        if not self.dirty():
            return True
        try:
            values = self.values()
            changes = {
                name: value
                for name, value in values.items()
                if value != self.initial["values"][name]
            }
            self.session.update(
                self.row_id,
                changes,
                self.include.isChecked(),
                self.reason.toPlainText(),
                self.status.currentData(),
            )
        except (ValueError, TypeError) as error:
            self.failed.emit(str(error))
            return False
        self.set_record(self.session, self.row_id)
        self.applied.emit()
        return True

    def reset(self) -> None:
        """Discard unapplied form changes."""
        self.set_record(self.session, self.row_id)

    def allow_leave(self) -> bool:
        """Offer apply, discard or cancel before leaving an edited form."""
        if not self.dirty():
            return True
        choice = QMessageBox.question(
            self,
            "Unapplied record edits",
            "Apply your record edits to the draft before continuing?",
            QMessageBox.StandardButton.Save
            | QMessageBox.StandardButton.Discard
            | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Save,
        )
        if choice == QMessageBox.StandardButton.Save:
            return self.apply()
        if choice == QMessageBox.StandardButton.Discard:
            self.reset()
            return True
        return False


class RevisionPanel(QWidget):
    """Compare bills, inspect either source and record explicit relationship decisions."""

    changed = Signal()
    failed = Signal(str)
    show_record = Signal(str)

    def __init__(self, parent=None):
        """Build accessible controls without changing accepted evidence."""
        super().__init__(parent)
        self.session: ReviewSession | None = None
        self.links = []
        self._pair = (None, None)
        self._initial_decision = ("suggested", "")
        self._refreshing = False
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        area = QScrollArea()
        area.setWidgetResizable(True)
        content = QWidget()
        layout = QVBoxLayout(content)
        area.setWidget(content)
        outer.addWidget(area)
        layout.addWidget(label(NOTICE, "hint", wrap=True))
        self.summary = label("Open an analysis to compare bill revisions.", "hint", wrap=True)
        layout.addWidget(self.summary)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(("Earlier bill", "Later bill / credit", "Decision"))
        self.tree.setRootIsDecorated(False)
        self.tree.setMinimumHeight(100)
        self.tree.setMaximumHeight(180)
        self.tree.setAccessibleName("Bill revision relationships")
        layout.addWidget(self.tree, 1)
        selectors = QFormLayout()
        selectors.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapAllRows)
        self.older, self.newer = QComboBox(), QComboBox()
        for control in (self.older, self.newer):
            control.setEditable(True)
            control.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
            control.setMinimumWidth(0)
            control.setMinimumContentsLength(8)
            control.setSizeAdjustPolicy(
                QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon
            )
        self.older.setAccessibleName("Earlier bill")
        self.newer.setAccessibleName("Later bill or cancellation credit")
        selectors.addRow("Earlier bill", self.older)
        selectors.addRow("Later bill / credit", self.newer)
        layout.addLayout(selectors)
        self.compare = QTableWidget(0, 3)
        self.compare.setHorizontalHeaderLabels(("Field", "Earlier", "Later"))
        self.compare.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.compare.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.compare.setAccessibleName("Bill revision field comparison")
        self.compare.setMinimumHeight(100)
        self.compare.setMaximumHeight(230)
        layout.addWidget(self.compare, 1)
        sources = QVBoxLayout()
        sources.addWidget(button("Show earlier source", lambda: self._show(self.older)))
        sources.addWidget(button("Show later source", lambda: self._show(self.newer)))
        layout.addLayout(sources)
        self.decision = QComboBox()
        for title, value in (
            ("Unreviewed / clear decision", "suggested"),
            ("Confirm relationship", "confirmed"),
            ("Reject relationship", "rejected"),
        ):
            self.decision.addItem(title, value)
        layout.addWidget(self.decision)
        self.reason = QPlainTextEdit()
        self.reason.setFixedHeight(55)
        self.reason.setPlaceholderText("Explain the source evidence supporting your decision")
        self.reason.setAccessibleName("Reason for bill revision decision")
        layout.addWidget(self.reason)
        layout.addWidget(button("Save decision to draft", self.save_relationship))
        self.tree.itemSelectionChanged.connect(self._selected)
        self.older.currentIndexChanged.connect(self._compare)
        self.newer.currentIndexChanged.connect(self._compare)
        self.setEnabled(False)

    @staticmethod
    def _title(item: dict) -> str:
        row = item["record"]
        return " · ".join(
            text(row.get(k)) for k in ("Invoice #", "Date", "Account", "Fuel") if text(row.get(k))
        ) + (" · EXCLUDED" if not item["included"] else "")

    def refresh(self, session: ReviewSession | None) -> None:
        """Refresh suggestions and comparisons while preserving unapplied decisions."""
        pending = (
            (self.decision.currentData(), self.reason.toPlainText())
            if self.session is session and self.dirty()
            else None
        )
        self.session = session
        self.setEnabled(session is not None)
        if session is None:
            return
        self._refreshing = True
        current = self.older.currentData(), self.newer.currentData()
        self.tree.clear()
        for control in (self.older, self.newer):
            control.clear()
            control.addItem("Choose a record…", None)
        items = {item["id"]: item for item in session.items()}
        for row_id, item in items.items():
            kind = item["record"].get("Entry Type")
            if kind in BILLS:
                self.older.addItem(self._title(item), row_id)
            if kind in BILLS | {"Credit"}:
                self.newer.addItem(self._title(item), row_id)
        model = session.revision_model()
        self.links = model["relationships"]
        for link in self.links:
            tree_item = QTreeWidgetItem(
                (
                    self._title(items[link["older"]]),
                    self._title(items[link["newer"]]),
                    link["status"].title(),
                )
            )
            tree_item.setData(0, Qt.ItemDataRole.UserRole, (link["older"], link["newer"]))
            tree_item.setToolTip(0, link["basis"])
            self.tree.addTopLevelItem(tree_item)
        confirmed = sum(link["status"] == "confirmed" for link in self.links)
        self.summary.setText(
            f"{len(self.links)} relationships · {confirmed} confirmed. Use the selectors to add a link the suggestions missed."
            + (
                " Showing the first 500 suggestions; explicit decisions are retained."
                if model["truncated"]
                else ""
            )
        )
        for control, key in zip((self.older, self.newer), current, strict=True):
            control.setCurrentIndex(max(0, control.findData(key)))
        self._pair = (None, None)
        self._initial_decision = (self.decision.currentData(), self.reason.toPlainText())
        self._refreshing = False
        self._compare()
        if pending is not None:
            self.decision.setCurrentIndex(self.decision.findData(pending[0]))
            self.reason.setPlainText(pending[1])

    def _selected(self) -> None:
        selected = self.tree.selectedItems()
        if not selected or self._refreshing:
            return
        if not self.allow_leave():
            return
        older, newer = selected[0].data(0, Qt.ItemDataRole.UserRole)
        self._refreshing = True
        self.older.setCurrentIndex(self.older.findData(older))
        self.newer.setCurrentIndex(self.newer.findData(newer))
        self._refreshing = False
        self._compare()

    def _compare(self) -> None:
        if self._refreshing or self.session is None:
            return
        pair = self.older.currentData(), self.newer.currentData()
        if pair != self._pair and not self.allow_leave():
            self._refreshing = True
            for control, value in zip((self.older, self.newer), self._pair, strict=True):
                control.setCurrentIndex(max(0, control.findData(value)))
            self._refreshing = False
            return
        self._pair = pair
        older, newer = pair
        self.compare.setRowCount(0)
        if not older or not newer:
            self.decision.setCurrentIndex(0)
            self.reason.clear()
            self._initial_decision = ("suggested", "")
            return
        left, right = self.session.record(older), self.session.record(newer)
        self.compare.setRowCount(len(COMPARE_FIELDS))
        for index, field in enumerate(COMPARE_FIELDS):
            for column, value in enumerate((field, text(left.get(field)), text(right.get(field)))):
                cell = QTableWidgetItem(value)
                cell.setToolTip(value)
                if left.get(field) != right.get(field):
                    cell.setBackground(QColor("#EAF7F5"))
                self.compare.setItem(index, column, cell)
        decision = self.session.revisions.get((older, newer), {})
        self.decision.setCurrentIndex(self.decision.findData(decision.get("status", "suggested")))
        self.reason.setPlainText(decision.get("reason", ""))
        self._initial_decision = (self.decision.currentData(), self.reason.toPlainText())

    def _show(self, control: QComboBox) -> None:
        if (
            control.currentIndex() >= 0
            and control.currentText() == control.itemText(control.currentIndex())
            and (row_id := control.currentData())
        ):
            self.show_record.emit(row_id)

    def save_relationship(self) -> bool:
        """Validate and save an explicit relationship decision to the draft."""
        if self.session is None:
            return False
        if any(
            c.currentIndex() < 0 or c.currentText() != c.itemText(c.currentIndex())
            for c in (self.older, self.newer)
        ):
            self.failed.emit(
                "Select a matching record from each list; typed text alone does not select evidence."
            )
            return False
        older, newer = self._pair
        if not older or not newer:
            self.failed.emit("Choose an earlier bill and its proposed replacement or credit.")
            return False
        kind = (
            "cancellation_credit"
            if self.session.record(newer).get("Entry Type") == "Credit"
            else "replacement"
        )
        try:
            self.session.set_revision(
                older, newer, kind, self.decision.currentData(), self.reason.toPlainText()
            )
        except (ValueError, TypeError) as error:
            self.failed.emit(str(error))
            return False
        self._initial_decision = (self.decision.currentData(), self.reason.toPlainText())
        self.changed.emit()
        return True

    def dirty(self) -> bool:
        """Detect an unapplied relationship decision or reason."""
        return (self.decision.currentData(), self.reason.toPlainText()) != self._initial_decision

    def allow_leave(self) -> bool:
        """Preserve relationship edits when changing comparison or leaving the workspace."""
        if not self.dirty():
            return True
        choice = QMessageBox.question(
            self,
            "Unapplied revision decision",
            "Save the relationship decision to the draft before continuing?",
            QMessageBox.StandardButton.Save
            | QMessageBox.StandardButton.Discard
            | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Save,
        )
        if choice == QMessageBox.StandardButton.Save:
            return self.save_relationship()
        if choice == QMessageBox.StandardButton.Discard:
            self.decision.setCurrentIndex(self.decision.findData(self._initial_decision[0]))
            self.reason.setPlainText(self._initial_decision[1])
            return True
        return False
