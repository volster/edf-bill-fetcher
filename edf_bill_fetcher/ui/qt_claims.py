"""Human claim grouping, credit allocation and complaint wording in a separate workspace."""

from __future__ import annotations

from collections.abc import Callable

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPlainTextEdit,
    QScrollArea,
    QTableWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from edf_bill_fetcher.application.review_session import ReviewSession
from edf_bill_fetcher.helpers.formatting import fmt_money
from edf_bill_fetcher.processors.claim_reconciliation import has_active_claims
from edf_bill_fetcher.processors.preparation import number
from edf_bill_fetcher.processors.revisions import BILLS
from edf_bill_fetcher.ui.qt_widgets import WrappedCheckBox, button, label


class ClaimDialog(QDialog):
    """Edit one shared review session; publish proposals only through a new reviewed pack."""

    changed = Signal()
    show_record = Signal(str)

    def __init__(self, session: ReviewSession, parent: QWidget | None = None) -> None:
        """Keep a scrollable editor separate from the main evidence layout's size hints."""
        super().__init__(parent)
        self.session = session
        self.publish_requested = False
        self.group_id: str | None = None
        self._loading = False
        self._initial_group: tuple = ()
        self._initial_overlap: tuple[tuple[str, str] | None, str] = (None, "")
        self._initial_context: tuple = ()
        self.setWindowTitle("Claim reconciliation and complaint draft")
        self.setMinimumSize(640, 500)
        screen = self.screen()
        available = screen.availableGeometry() if screen is not None else None
        self.resize(
            max(640, min(980, available.width() - 48)) if available is not None else 980,
            max(500, min(760, available.height() - 80)) if available is not None else 760,
        )
        outer = QVBoxLayout(self)
        outer.addWidget(label("Reconcile the claim", "title"))
        outer.addWidget(
            label(
                "Group repeated demands for the same charges, then enter one consolidated proposal. The original balance remains unchanged.",
                "hint",
                wrap=True,
            )
        )
        self.summary = label("", "hint", wrap=True)
        outer.addWidget(self.summary)
        self.feedback = label("", "feedback", wrap=True)
        self.feedback.hide()
        outer.addWidget(self.feedback)
        self.groups = QTreeWidget()
        self.groups.setHeaderLabels(("Claim", "Title", "Review", "Further reduction"))
        self.groups.setRootIsDecorated(False)
        self.groups.header().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.groups.header().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.groups.setMaximumHeight(150)
        self.groups.setAccessibleName("Consolidated claim groups")
        outer.addWidget(self.groups)
        actions = QHBoxLayout()
        actions.addWidget(button("New group", self.new_group))
        actions.addWidget(button("Remove group", self.remove_group))
        actions.addStretch()
        outer.addLayout(actions)
        from PySide6.QtWidgets import QTabWidget

        self.tabs = QTabWidget()
        outer.addWidget(self.tabs, 1)
        self._build_group()
        self._build_overlap()
        self._build_context()
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Close
        )
        self.publish = buttons.addButton(
            "Create reviewed pack", QDialogButtonBox.ButtonRole.ActionRole
        )
        self.publish.clicked.connect(self.publish_pack)
        buttons.button(QDialogButtonBox.StandardButton.Save).clicked.connect(self.save_all)
        buttons.rejected.connect(self.reject)
        outer.addWidget(buttons)
        self.groups.itemSelectionChanged.connect(self._select_group)
        context = self.session.complaint_context
        self.complaint_summary.setPlainText(context["complaint_summary"])
        self.requested_outcome.setPlainText(context["requested_outcome"])
        self.include_pack.setChecked(context["include_pack"])
        self._initial_context = self._context_values()
        self.refresh()
        self._load_group(None)

    def _page(self, title: str) -> QVBoxLayout:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        body = QWidget()
        layout = QVBoxLayout(body)
        scroll.setWidget(body)
        self.tabs.addTab(scroll, title)
        return layout

    def _build_group(self) -> None:
        layout = self._page("Claims and credits")
        form = QFormLayout()
        form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapAllRows)
        self.claim_id, self.title = QLineEdit(), QLineEdit()
        self.claim_id.setAccessibleName("Claim group identifier")
        self.title.setAccessibleName("Claim group title")
        self.status = QComboBox()
        for text, key in (
            ("Draft", "draft"),
            ("Reviewed for totaling", "reviewed"),
            ("Withdrawn", "withdrawn"),
        ):
            self.status.addItem(text, key)
        self.challenged, self.reduction = QLineEdit(), QLineEdit()
        self.challenged.setAccessibleName("Consolidated amount challenged")
        self.reduction.setAccessibleName("Reduction sought before relief already received")
        for title, field in (
            ("Claim ID", self.claim_id),
            ("Title", self.title),
            ("Review state", self.status),
            ("Consolidated amount challenged (£, including applicable VAT)", self.challenged),
            ("Reduction sought before relief already received (£)", self.reduction),
        ):
            form.addRow(title, field)
        self.reason, self.account_link = QPlainTextEdit(), QPlainTextEdit()
        self.reason.setFixedHeight(85)
        self.reason.setAccessibleName("Claim scope and reconciliation reason")
        self.reason.setPlaceholderText(
            "Identify the charges challenged once, account for earlier demands, and explain the consolidated amount. Reviewed means you checked this allocation; it is not a legal finding."
        )
        self.account_link.setFixedHeight(60)
        self.account_link.setAccessibleName("Evidence linking different account references")
        self.account_link.setPlaceholderText(
            "Only for account migrations: identify the source linking the accounts and shared meter/supply."
        )
        form.addRow("Scope and source-backed reconciliation", self.reason)
        form.addRow("Account-link evidence, if applicable", self.account_link)
        layout.addLayout(form)
        layout.addWidget(label("Bills challenged by this group", "sectionTitle"))
        layout.addWidget(
            label(
                "Select the disputed documents representing the same economic charges. Earlier supporting bills remain in the history without being challenged separately.",
                "hint",
                wrap=True,
            )
        )
        self.members = QListWidget()
        self.members.setMinimumHeight(140)
        self.members.setMaximumHeight(220)
        self.members.setAccessibleName("Bills included in the consolidated challenge")
        layout.addWidget(self.members)
        layout.addWidget(button("Show selected bill in evidence workspace", self.reveal_bill))
        layout.addWidget(label("Credit allocations", "sectionTitle"))
        layout.addWidget(
            label(
                "Already applied relief reduces the further reduction requested. A cancellation offset explains the earlier bill/credit history and is not deducted again. Allocate only amounts supported by the original negative transaction.",
                "hint",
                wrap=True,
            )
        )
        self.credits = QTableWidget(0, 4)
        self.credits.setHorizontalHeaderLabels(
            (
                "Source credit / net credit bill",
                "Amount (£)",
                "Treatment",
                "Reason / allocation evidence",
            )
        )
        self.credits.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.credits.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.credits.horizontalHeader().setStretchLastSection(True)
        self.credits.setColumnWidth(0, 250)
        self.credits.setColumnWidth(1, 100)
        self.credits.setColumnWidth(2, 200)
        self.credits.setMinimumHeight(130)
        self.credits.setAccessibleName("Allocations of credits already recorded")
        layout.addWidget(self.credits)
        controls = QHBoxLayout()
        controls.addWidget(button("Add credit allocation", self.add_credit))
        controls.addWidget(button("Remove selected allocation", self.remove_credit))
        layout.addLayout(controls)
        self.group_notes = label("", "hint", wrap=True)
        layout.addWidget(self.group_notes)

    def _build_overlap(self) -> None:
        layout = self._page("Overlap review")
        layout.addWidget(
            label(
                "Potential overlaps withhold the total until resolved. Put repeated demands in one group. Confirm separate charges only when the source evidence supports it; a shared bill cannot be waived as independent.",
                "hint",
                wrap=True,
            )
        )
        self.overlaps = QTreeWidget()
        self.overlaps.setHeaderLabels(("Groups", "Reason for review", "Decision"))
        self.overlaps.setRootIsDecorated(False)
        self.overlaps.header().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.overlaps.header().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.overlaps.setMinimumHeight(140)
        self.overlaps.setAccessibleName("Potentially overlapping claims")
        layout.addWidget(self.overlaps)
        self.overlap_reason = QPlainTextEdit()
        self.overlap_reason.setFixedHeight(100)
        self.overlap_reason.setAccessibleName(
            "Evidence that overlapping claims represent separate charges"
        )
        layout.addWidget(self.overlap_reason)
        layout.addWidget(button("Confirm separate charges", self.save_overlap))
        layout.addWidget(button("Clear selected overlap decision", self.clear_overlap))
        self.overlaps.itemSelectionChanged.connect(self._select_overlap)

    def _build_context(self) -> None:
        layout = self._page("Complaint draft")
        layout.addWidget(
            label(
                "Describe the problem in your own words. The accepted schedule supplies the figures and sources; unresolved totals stay unquantified. The draft is not sent automatically.",
                "hint",
                wrap=True,
            )
        )
        self.complaint_summary, self.requested_outcome = QPlainTextEdit(), QPlainTextEdit()
        self.complaint_summary.setAccessibleName("Complaint summary")
        self.requested_outcome.setAccessibleName("Requested complaint outcome")
        self.complaint_summary.setMinimumHeight(120)
        self.requested_outcome.setMinimumHeight(100)
        layout.addWidget(label("Complaint summary", "fieldLabel"))
        layout.addWidget(self.complaint_summary)
        layout.addWidget(label("Outcome requested from EDF", "fieldLabel"))
        layout.addWidget(self.requested_outcome)
        self.include_pack = WrappedCheckBox(
            "Include a focused complaint pack when creating reviewed outputs"
        )
        layout.addWidget(self.include_pack)
        layout.addWidget(
            label(
                "The focused pack retains whole original PDFs, including any other pages in a shared document. Review and redact a separate copy before sharing if necessary.",
                "hint",
                wrap=True,
            )
        )

    @staticmethod
    def _record_label(item: dict) -> str:
        row = item["record"]
        return " · ".join(
            str(row.get(key) or "Unknown") for key in ("Invoice #", "Date", "Account")
        ) + (" · EXCLUDED" if not item["included"] else "")

    def refresh(self) -> None:
        """Refresh computed readiness while retaining unapplied group and complaint text."""
        self._loading = True
        model = self.session.claim_model()
        self.groups.clear()
        for group in model["groups"]:
            data = group["input"]
            state = (
                "Withdrawn"
                if data["status"] == "withdrawn"
                else "Needs review"
                if group["needs_review"]
                else "Reviewed"
                if group["ready"]
                else data["status"].title()
                + (" · unresolved" if data["status"] == "reviewed" else "")
            )
            item = QTreeWidgetItem(
                (data["id"], data["title"], state, fmt_money(group["further_reduction"]))
            )
            item.setData(0, Qt.ItemDataRole.UserRole, data["id"])
            item.setToolTip(2, "\n".join(group["issues"]))
            self.groups.addTopLevelItem(item)
            if data["id"] == self.group_id:
                self.groups.setCurrentItem(item)
        self.overlaps.clear()
        pair, pending = self._initial_overlap[0], self.overlap_reason.toPlainText()
        for overlap in model["overlaps"]:
            item = QTreeWidgetItem(
                (
                    " / ".join(overlap["groups"]),
                    overlap["kind"].replace("_", " "),
                    "Separate charges" if overlap["resolved"] else "Needs review",
                )
            )
            item.setData(0, Qt.ItemDataRole.UserRole, tuple(overlap["groups"]))
            self.overlaps.addTopLevelItem(item)
            if tuple(overlap["groups"]) == pair:
                self.overlaps.setCurrentItem(item)
        self.overlap_reason.setPlainText(pending)
        self.summary.setText(
            "No active claims. Withdrawn proposals remain in the review history."
            if not has_active_claims(model)
            else f"Further reduction proposed: {fmt_money(model['proposed_reduction'])}."
            if model["ready"]
            else f"Total withheld · reviewed subtotal {fmt_money(model['reviewed_subtotal'])} · {len(model['unresolved_positions'])} challenged documents still need reconciliation."
        )
        group = next(
            (group for group in model["groups"] if group["input"]["id"] == self.group_id), None
        )
        self.group_notes.setText(
            "\n".join(group["issues"])
            if group
            else "Save as a draft while gathering the evidence. Review it only after the consolidated figures and credit allocations are checked."
        )
        self._loading = False

    def _form_snapshot(self) -> tuple:
        members = tuple(
            sorted(
                self.members.item(n).data(Qt.ItemDataRole.UserRole)
                for n in range(self.members.count())
                if self.members.item(n).checkState() == Qt.CheckState.Checked
            )
        )
        credits = []
        for n in range(self.credits.rowCount()):
            source, amount, treatment, reason = (
                self.credits.cellWidget(n, column) for column in range(4)
            )
            assert isinstance(source, QComboBox)
            assert isinstance(amount, QLineEdit)
            assert isinstance(treatment, QComboBox)
            assert isinstance(reason, QLineEdit)
            credits.append(
                (
                    source.currentData(),
                    source.currentText(),
                    amount.text(),
                    treatment.currentData(),
                    reason.text(),
                )
            )
        return (
            self.claim_id.text(),
            self.title.text(),
            self.status.currentData(),
            self.challenged.text(),
            self.reduction.text(),
            self.reason.toPlainText(),
            self.account_link.toPlainText(),
            members,
            tuple(credits),
        )

    def _context_values(self) -> tuple:
        return (
            self.complaint_summary.toPlainText(),
            self.requested_outcome.toPlainText(),
            self.include_pack.isChecked(),
        )

    def _load_group(self, key: str | None) -> None:
        self.group_id = key
        data = self.session.claims.get(key, {}) if key else {}
        candidate = 1
        while f"C{candidate:03d}" in self.session.claims:
            candidate += 1
        self.claim_id.setText(data.get("id", f"C{candidate:03d}"))
        self.claim_id.setReadOnly(key is not None)
        self.title.setText(data.get("title", ""))
        self.status.setCurrentIndex(self.status.findData(data.get("status", "draft")))
        self.challenged.setText(data.get("challenged") or "")
        self.reduction.setText(data.get("reduction") or "")
        self.reason.setPlainText(data.get("reason", ""))
        self.account_link.setPlainText(data.get("account_link_reason", ""))
        self.members.clear()
        for item in self.session.items():
            if item["record"].get("Entry Type") not in BILLS:
                continue
            member = QListWidgetItem(self._record_label(item))
            member.setData(Qt.ItemDataRole.UserRole, item["id"])
            member.setFlags(member.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            member.setCheckState(
                Qt.CheckState.Checked
                if item["id"] in data.get("members", [])
                else Qt.CheckState.Unchecked
            )
            member.setToolTip(str(item["record"].get("Attachment Name") or "Source unavailable"))
            self.members.addItem(member)
        self.credits.setRowCount(0)
        for credit in data.get("credits", []):
            self.add_credit(credit)
        self._initial_group = self._form_snapshot()
        self.refresh()

    def add_credit(self, value: dict | None = None) -> None:
        """Add an editable allocation with an explicit source and treatment choice."""
        value = value or {}
        row = self.credits.rowCount()
        self.credits.insertRow(row)
        source, treatment = QComboBox(), QComboBox()
        source.addItem("Choose a source…", None)
        for item in self.session.items():
            available = number(item["record"].get("Transaction Amount (£)"))
            if item["record"].get("Entry Type") in BILLS | {"Credit"} and (
                (available is not None and available < 0) or item["id"] == value.get("row_id")
            ):
                source.addItem(
                    self._record_label(item)
                    + " · "
                    + (fmt_money(available) if available is not None else "amount unknown"),
                    item["id"],
                )
        source.setEditable(True)
        source.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        source.setMinimumContentsLength(8)
        source.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        source.setCurrentIndex(max(0, source.findData(value.get("row_id"))))
        source.setAccessibleName("Source for credit allocation")
        amount = QLineEdit(value.get("amount", ""))
        amount.setAccessibleName("Allocated credit amount")
        for title, key in (
            ("Choose treatment…", None),
            ("Cancellation offset", "cancellation_offset"),
            ("Already applied relief", "already_applied_relief"),
        ):
            treatment.addItem(title, key)
        treatment.setCurrentIndex(max(0, treatment.findData(value.get("treatment"))))
        treatment.setAccessibleName("Credit allocation treatment")
        reason = QLineEdit(value.get("reason", ""))
        reason.setMaxLength(4000)
        reason.setAccessibleName("Reason for this credit allocation")
        for column, control in enumerate((source, amount, treatment, reason)):
            self.credits.setCellWidget(row, column, control)
        self.credits.setRowHeight(row, 42)
        self.credits.setCurrentCell(row, 0)

    def remove_credit(self) -> None:
        """Remove the selected unapplied allocation; saving records the new group decision."""
        row = self.credits.currentRow()
        if row >= 0:
            self.credits.removeRow(row)

    def _error(self, error: object) -> None:
        self.feedback.setText(str(error))
        self.feedback.show()

    def save_group(self) -> bool:
        """Apply one complete group, preserving invalid form text for correction."""
        try:
            fields = self._form_snapshot()
            credits = []
            for row_number, (source, selected_text, amount, treatment, reason) in enumerate(
                fields[8]
            ):
                widget = self.credits.cellWidget(row_number, 0)
                if isinstance(widget, QComboBox) and (
                    widget.currentIndex() < 0
                    or selected_text != widget.itemText(widget.currentIndex())
                ):
                    raise ValueError(
                        "Choose a listed credit source; typed text alone does not select evidence"
                    )
                if source is None and not amount and treatment is None and not reason:
                    continue
                credits.append(
                    {"row_id": source, "amount": amount, "treatment": treatment, "reason": reason}
                )
            self.session.set_claim(
                fields[0],
                title=fields[1],
                status=fields[2],
                challenged=fields[3],
                reduction=fields[4],
                reason=fields[5],
                account_link_reason=fields[6],
                members=list(fields[7]),
                credits=credits,
            )
        except (ValueError, TypeError) as error:
            self._error(error)
            return False
        self._load_group(fields[0])
        self.feedback.hide()
        self.changed.emit()
        return True

    def save_context(self) -> bool:
        """Keep complaint wording and focused-pack inclusion in the shared audited draft."""
        try:
            self.session.set_complaint_context(*self._context_values())
        except (ValueError, TypeError) as error:
            self._error(error)
            return False
        self._initial_context = self._context_values()
        self.changed.emit()
        return True

    def new_group(self) -> None:
        """Start another group without discarding unapplied work."""
        if self._allow_group_leave():
            self._load_group(None)
            self.tabs.setCurrentIndex(0)

    def remove_group(self) -> None:
        """Remove a saved group and dependent overlap decisions with normal session undo."""
        if self.group_id and self._allow_group_leave():
            self.session.clear_claim(self.group_id)
            self._load_group(None)
            self.changed.emit()

    def _select_group(self) -> None:
        if self._loading:
            return
        item = self.groups.currentItem()
        key = item.data(0, Qt.ItemDataRole.UserRole) if item else None
        if key and key != self.group_id:
            if self._allow_group_leave():
                self._load_group(key)
            else:
                self.refresh()

    def _select_overlap(self) -> None:
        if self._loading:
            return
        item = self.overlaps.currentItem()
        pair = item.data(0, Qt.ItemDataRole.UserRole) if item else None
        if pair == self._initial_overlap[0]:
            return
        if not self._allow_overlap_leave():
            self.refresh()
            return
        decision = self.session.claim_separations.get(pair, {}) if pair else {}
        self.overlap_reason.setPlainText(decision.get("reason", ""))
        self._initial_overlap = (pair, self.overlap_reason.toPlainText())

    def save_overlap(self) -> bool:
        """Record the evidence that resolves a potential overlap as separate charges."""
        pair = self._initial_overlap[0]
        if not pair:
            self._error("Select a pair of potentially overlapping groups")
            return False
        try:
            self.session.separate_claims(*pair, self.overlap_reason.toPlainText())
        except (ValueError, TypeError) as error:
            self._error(error)
            return False
        self._initial_overlap = (pair, self.overlap_reason.toPlainText())
        self.refresh()
        self.changed.emit()
        return True

    def clear_overlap(self) -> None:
        """Remove an earlier separate-charges decision and withhold the total again."""
        pair = self._initial_overlap[0]
        if pair:
            self.session.clear_separation(*pair)
            self.overlap_reason.clear()
            self._initial_overlap = (pair, "")
            self.refresh()
            self.changed.emit()

    def _confirm_pending(
        self, title: str, save: Callable[[], bool], discard: Callable[[], None]
    ) -> bool:
        choice = QMessageBox.question(
            self,
            title,
            "Save these edits before continuing?",
            QMessageBox.StandardButton.Save
            | QMessageBox.StandardButton.Discard
            | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Save,
        )
        if choice == QMessageBox.StandardButton.Save:
            return bool(save())
        if choice == QMessageBox.StandardButton.Discard:
            discard()
            return True
        return False

    def _allow_group_leave(self) -> bool:
        if not self._initial_group or self._form_snapshot() == self._initial_group:
            return True
        return self._confirm_pending(
            "Unapplied claim group", self.save_group, lambda: self._load_group(self.group_id)
        )

    def _allow_overlap_leave(self) -> bool:
        if self.overlap_reason.toPlainText() == self._initial_overlap[1]:
            return True
        return self._confirm_pending(
            "Unapplied overlap review",
            self.save_overlap,
            lambda: self.overlap_reason.setPlainText(self._initial_overlap[1]),
        )

    def _allow_leave(self) -> bool:
        if not self._allow_group_leave() or not self._allow_overlap_leave():
            return False
        if self._context_values() != self._initial_context:
            return self._confirm_pending(
                "Unapplied complaint wording", self.save_context, self._restore_context
            )
        return True

    def _restore_context(self) -> None:
        summary, outcome, include = self._initial_context
        self.complaint_summary.setPlainText(summary)
        self.requested_outcome.setPlainText(outcome)
        self.include_pack.setChecked(include)

    def save_all(self) -> bool:
        """Apply only changed forms; an empty unused new-group form creates no claim."""
        if self._form_snapshot() != self._initial_group and not self.save_group():
            return False
        if (
            self.overlap_reason.toPlainText() != self._initial_overlap[1]
            and not self.save_overlap()
        ):
            return False
        if self._context_values() != self._initial_context and not self.save_context():
            return False
        return True

    def publish_pack(self) -> None:
        """Ask the owning evidence workspace to run the normal reviewed-pack job."""
        if self.save_all():
            self.publish_requested = True
            self.accept()

    def reveal_bill(self) -> None:
        """Return to the original evidence page after preserving the pending form."""
        item = self.members.currentItem()
        if item is not None and self._allow_leave():
            self.show_record.emit(item.data(Qt.ItemDataRole.UserRole))
            super().accept()

    def reject(self) -> None:
        """Protect pending text and allocations when closing or pressing Escape."""
        if self._allow_leave():
            super().reject()

    def closeEvent(self, event: QCloseEvent) -> None:
        """Keep the editor open when a pending decision cannot be saved or is cancelled."""
        event.accept() if self._allow_leave() else event.ignore()
