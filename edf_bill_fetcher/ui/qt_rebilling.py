"""Full rebilling histories and editable dispute positions beside original PDF previews."""

from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QComboBox,
    QFormLayout,
    QHBoxLayout,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QScrollArea,
    QToolButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from edf_bill_fetcher.application.review_session import ReviewSession
from edf_bill_fetcher.models.output_tables import ledger_label
from edf_bill_fetcher.models.rebilling_tables import (
    comparison_table,
    document_label,
    history_table,
    position_state,
)
from edf_bill_fetcher.processors.rebilling_ledger import NOTICE, SCHEDULE_FIELDS, STATUSES
from edf_bill_fetcher.processors.revisions import BILLS
from edf_bill_fetcher.ui.qt_widgets import button, label


class RebillingPanel(QWidget):
    """Keep proposals separate from recorded account figures, with normal draft recovery."""

    changed = Signal()
    failed = Signal(str)
    show_record = Signal(str)

    def __init__(self, parent=None):
        """Build a scrollable history and human-review form for small desktop windows."""
        super().__init__(parent)
        self.session: ReviewSession | None = None
        self.row_id = ""
        self.model: dict = {}
        self._refreshing = False
        self._initial: tuple = ()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        area = QScrollArea()
        area.setWidgetResizable(True)
        content = QWidget()
        layout = QVBoxLayout(content)
        area.setWidget(content)
        outer.addWidget(area)
        layout.addWidget(
            label(
                "Trace bills and credits, then record a proposed challenge. Account balances stay separate.",
                "hint",
                wrap=True,
            )
        )
        self.help_button = QToolButton()
        self.help_button.setText("How to interpret these amounts")
        self.help_button.setCheckable(True)
        self.help_button.setArrowType(Qt.ArrowType.RightArrow)
        self.help_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.help_button.setAccessibleName("Explain rebilling amounts and dispute proposals")
        self.help_text = label(NOTICE, "hint", wrap=True)
        self.help_text.setVisible(False)
        self.help_button.toggled.connect(self._toggle_help)
        layout.addWidget(self.help_button)
        layout.addWidget(self.help_text)
        self.summary = label(
            "Open an analysis to inspect the rebilling history.", "hint", wrap=True
        )
        layout.addWidget(self.summary)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(("Document / date", "Type", "Dispute position"))
        self.tree.setMinimumHeight(140)
        self.tree.setMaximumHeight(230)
        self.tree.setAccessibleName("Complete rebilling document history")
        layout.addWidget(self.tree)
        self.details = QPlainTextEdit()
        self.details.setReadOnly(True)
        self.details.setMinimumHeight(120)
        self.details.setMaximumHeight(230)
        self.details.setAccessibleName("Rebilling components and comparisons")
        layout.addWidget(self.details)
        layout.addWidget(button("Show this original source", self._show))
        self.status = QComboBox()
        for value in STATUSES:
            self.status.addItem(value.replace("_", " ").title(), value)
        self.status.setAccessibleName("Dispute position status")
        self.challenged, self.reduction, self.first_demand = QLineEdit(), QLineEdit(), QLineEdit()
        for field, title in (
            (self.challenged, "Challenged amount"),
            (self.reduction, "Proposed reduction"),
            (self.first_demand, "Verified first demand date"),
        ):
            field.setAccessibleName(title)
        self.first_demand.setPlaceholderText("DD/MM/YYYY; blank if not verified")
        form = QFormLayout()
        form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapAllRows)
        for title, form_control in (
            ("Position", self.status),
            ("Challenged amount (£)", self.challenged),
            ("Proposed reduction (£)", self.reduction),
            ("First demand for the extra amount", self.first_demand),
        ):
            form.addRow(title, form_control)
        self.basis, self.request = QPlainTextEdit(), QPlainTextEdit()
        for control, title, placeholder in (
            (
                self.basis,
                "Dispute basis and scope",
                "Identify the charges challenged and source evidence. Explain how earlier demands and credits have been accounted for.",
            ),
            (
                self.request,
                "Evidence EDF must substantiate",
                "For example: earliest demand for each extra charge, meter readings, allocation between periods, cancellation and credit trail.",
            ),
        ):
            control.setFixedHeight(80)
            control.setAccessibleName(title)
            control.setPlaceholderText(placeholder)
            form.addRow(title, control)
        self.form_widget = QWidget()
        self.form_widget.setLayout(form)
        layout.addWidget(self.form_widget)
        self.use_adjustment = button("Use full positive bill adjustment", self._full_adjustment)
        layout.addWidget(self.use_adjustment)
        layout.addWidget(
            label(
                "Amounts are your proposed position. No percentage is conceded automatically. First demand is not inferred from a reading or repeated statement. Positions with overlapping periods must be reconciled before any total is claimed.",
                "hint",
                wrap=True,
            )
        )
        actions = QHBoxLayout()
        self.save_button = button("Save position to draft", self.save_position)
        self.clear_button = button("Clear position", self.clear_position)
        actions.addWidget(self.save_button)
        actions.addWidget(self.clear_button)
        layout.addLayout(actions)
        self.tree.itemSelectionChanged.connect(self._selected)
        self.setEnabled(False)

    def _toggle_help(self, expanded: bool) -> None:
        self.help_text.setVisible(expanded)
        self.help_button.setArrowType(
            Qt.ArrowType.DownArrow if expanded else Qt.ArrowType.RightArrow
        )

    def _values(self) -> tuple:
        return (
            self.status.currentData(),
            self.challenged.text(),
            self.reduction.text(),
            self.first_demand.text(),
            self.basis.toPlainText(),
            self.request.toPlainText(),
        )

    def dirty(self) -> bool:
        """Recognise unapplied edits independently of the saved draft."""
        return bool(self._initial and self._values() != self._initial)

    def refresh(self, session: ReviewSession | None) -> None:
        """Retain unapplied form text when other review tools update the shared draft."""
        pending = self._values() if self.session is session and self.dirty() else None
        old_id = self.row_id if self.session is session else ""
        self.session = session
        self.setEnabled(session is not None)
        self._refreshing = True
        self.tree.clear()
        self.model = session.ledger_model() if session else {}
        groups = {}
        selection = None
        docs = self.model.get("documents", [])
        for doc in docs:
            group_name = ledger_label(doc["record"])
            if group_name not in groups:
                group = QTreeWidgetItem((group_name, "", ""))
                self.tree.addTopLevelItem(group)
                group.setFlags(group.flags() & ~Qt.ItemFlag.ItemIsSelectable)
                group.setExpanded(True)
                groups[group_name] = group
            item = QTreeWidgetItem(
                (document_label(doc), str(doc["record"].get("Entry Type", "")), position_state(doc))
            )
            item.setData(0, Qt.ItemDataRole.UserRole, doc["row_id"])
            groups[group_name].addChild(item)
            if doc["row_id"] == old_id or selection is None:
                selection = item
        self.summary.setText(
            f"{len(docs)} documents · {sum(doc['position'] is not None for doc in docs)} saved positions · {sum(doc['needs_review'] for doc in docs)} need review. Proposals are not summed."
        )
        if selection is not None:
            self.tree.setCurrentItem(selection)
        self._refreshing = False
        self.row_id = selection.data(0, Qt.ItemDataRole.UserRole) if selection else ""
        self._load()
        if pending is not None:
            self._set_values(pending)

    def _set_values(self, values: tuple) -> None:
        self.status.setCurrentIndex(max(0, self.status.findData(values[0])))
        for control, value in zip(
            (self.challenged, self.reduction, self.first_demand), values[1:4], strict=True
        ):
            control.setText(value or "")
        self.basis.setPlainText(values[4] or "")
        self.request.setPlainText(values[5] or "")

    def _load(self) -> None:
        doc = next(
            (doc for doc in self.model.get("documents", []) if doc["row_id"] == self.row_id), None
        )
        if doc is None:
            self.details.clear()
            self._initial = ()
            return
        position = doc["position"] or {}
        self._set_values(
            tuple(
                position.get(field) or ("unreviewed" if field == "status" else "")
                for field in SCHEDULE_FIELDS
            )
        )
        self._initial = self._values()
        is_bill = doc["record"].get("Entry Type") in BILLS
        for control in (self.form_widget, self.save_button, self.clear_button, self.use_adjustment):
            control.setEnabled(is_bill)
        index = self.model["documents"].index(doc) + 1
        headers, values = history_table(self.model)[0], history_table(self.model)[index]
        detail = "\n\n".join(
            f"{header}\n{value}" for header, value in zip(headers, values, strict=True)
        )
        detail += "\n\nObserved components (not additional transactions):\n" + "\n".join(
            f"{part['kind'].title()}: £{part['amount'] or 'unknown'}; quantity {part['units'] if part['units'] is not None else 'unknown'}; rate {part['rate'] if part['rate'] is not None else 'unknown'}. {part['detail']}"
            for part in doc["components"]
        )
        comparisons = comparison_table(self.model)
        selected_comparisons = [
            row
            for link, row in zip(self.model["comparisons"], comparisons[1:], strict=True)
            if self.row_id in (link["older"], link["newer"])
        ]
        for row in selected_comparisons:
            detail += "\n\n" + "\n".join(row)
        if doc["needs_review"]:
            detail = (
                "NEEDS REVIEW: evidence or relationship decisions changed. Verify the position and save it again.\n\n"
                + detail
            )
        self.details.setPlainText(detail)

    def _selected(self) -> None:
        if self._refreshing:
            return
        item = self.tree.currentItem()
        key = item.data(0, Qt.ItemDataRole.UserRole) if item else ""
        if not key or key == self.row_id:
            return
        if not self.allow_leave():
            self.refresh(self.session)
            return
        self.row_id = key
        self._refreshing = True
        for n in range(self.tree.topLevelItemCount()):
            group = self.tree.topLevelItem(n)
            if group is None:
                continue
            for child in range(group.childCount()):
                candidate = group.child(child)
                if candidate is not None and candidate.data(0, Qt.ItemDataRole.UserRole) == key:
                    self.tree.setCurrentItem(candidate)
        self._refreshing = False
        self._load()

    def _show(self) -> None:
        if self.row_id:
            self.show_record.emit(self.row_id)

    def _full_adjustment(self) -> None:
        if self.session and self.row_id:
            amount = self.session.record(self.row_id).get("Period Charge (£)")
            if isinstance(amount, (int, float)) and amount > 0:
                self.challenged.setText(f"{amount:.2f}")
                self.reduction.setText(f"{amount:.2f}")
                self.status.setCurrentIndex(self.status.findData("challenged"))
            else:
                self.failed.emit(
                    "No positive net bill adjustment is available. Enter the specific replacement charges you challenge; a net credit does not cap them."
                )

    def save_position(self) -> bool:
        """Validate and save one human position through the shared review session."""
        if not self.session or not self.row_id:
            return False
        try:
            self.session.set_position(
                self.row_id, **dict(zip(SCHEDULE_FIELDS, self._values(), strict=True))
            )
        except (ValueError, TypeError) as error:
            self.failed.emit(str(error))
            return False
        self._initial = self._values()
        self.changed.emit()
        return True

    def clear_position(self) -> None:
        """Remove the proposal while retaining its accepted audit history and undo action."""
        if self.session and self.row_id:
            self.session.clear_position(self.row_id)
            self._initial = ()
            self.changed.emit()

    def allow_leave(self) -> bool:
        """Preserve uncommitted amounts and explanations before switching documents."""
        if not self.dirty():
            return True
        choice = QMessageBox.question(
            self,
            "Unapplied dispute position",
            "Save this position to the draft before continuing?",
            QMessageBox.StandardButton.Save
            | QMessageBox.StandardButton.Discard
            | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Save,
        )
        if choice == QMessageBox.StandardButton.Save:
            return self.save_position()
        if choice == QMessageBox.StandardButton.Discard:
            self._set_values(self._initial)
            return True
        return False
