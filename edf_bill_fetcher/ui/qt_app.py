"""Modern desktop workspace for extraction, saved reports and preferences."""

from __future__ import annotations

import json
import math
import os
import sys
import tempfile
from datetime import date, datetime, timezone
from pathlib import Path
from typing import cast
from uuid import uuid4

from PySide6.QtCore import Qt, QTimer, QUrl, Slot
from PySide6.QtGui import QCloseEvent, QDesktopServices, QResizeEvent
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLineEdit,
    QMainWindow,
    QPlainTextEdit,
    QProgressBar,
    QRadioButton,
    QScrollArea,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from edf_bill_fetcher.application.run_service import RunOutcome
from edf_bill_fetcher.domain.json_values import freeze_json, object_value
from edf_bill_fetcher.domain.requests import RunRequest, SourceSelection
from edf_bill_fetcher.helpers.artifacts import write_json
from edf_bill_fetcher.models.report_sections import REPORT_SECTIONS
from edf_bill_fetcher.ui.qt_jobs import DesktopJob, RunWorker
from edf_bill_fetcher.ui.qt_theme import apply_theme
from edf_bill_fetcher.ui.qt_widgets import Card, PathField, button, label

SECTIONS = [
    ("cover", "Cover page"),
    ("toc", "Contents"),
    *[(section.key, section.title) for section in REPORT_SECTIONS],
]
FORMATS = {"both": ("pdf", "docx", "html"), "pdf": ("pdf",), "docx": ("docx",), "html": ("html",)}
DEFAULTS = {
    "use_anchors": True,
    "use_large": True,
    "use_reading_classification": True,
    "use_pdf_fields": True,
    "use_acc_filter": False,
    "filter_below": False,
    "save_filtered": True,
    "use_dedup": True,
    "save_dups": True,
    "amalgamate_duplicates": False,
    "use_domain_filter": True,
    "save_evidence_files": True,
    "scan_sap_dumps": True,
    "generate_reconciliation_sheet": True,
}


def scroll_page(title: str, description: str) -> tuple[QScrollArea, QVBoxLayout]:
    """Keep page content scrollable while the window actions remain visible."""
    content = QWidget()
    content.setObjectName("page")
    layout = QVBoxLayout(content)
    layout.setContentsMargins(28, 24, 28, 28)
    layout.setSpacing(18)
    heading = QVBoxLayout()
    heading.setSpacing(7)
    heading.addWidget(label("EVIDENCE WORKSPACE", "eyebrow"))
    heading.addWidget(label(title, "title", wrap=True))
    heading.addWidget(label(description, "subtitle", wrap=True))
    layout.addLayout(heading)
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setWidget(content)
    return scroll, layout


class ReportDialog(QDialog):
    """Remember report choices and keep saving/cancelling outside the scroll area."""

    def __init__(self, options: dict, parent: QWidget | None = None) -> None:
        """Show meaningful format descriptions and compact section selection."""
        super().__init__(parent)
        self.setWindowTitle("Report settings")
        self.resize(680, 680)
        self.setMinimumSize(560, 460)
        self.selection: dict | None = None
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        page, layout = scroll_page(
            "Report settings", "Choose the files and sections for your next report."
        )
        outer.addWidget(page, 1)
        formats = Card("File formats")
        self.formats = QButtonGroup(self)
        self.format_keys = list(FORMATS)
        grid = QGridLayout()
        for index, (key, title, hint) in enumerate(
            [
                ("both", "All formats", "PDF, Word and HTML copies"),
                ("pdf", "PDF", "Ready to share or print"),
                ("docx", "Word document", "Edit before sharing"),
                ("html", "HTML", "Open in your browser"),
            ]
        ):
            radio = QRadioButton(title)
            radio.setToolTip(hint)
            self.formats.addButton(radio, index)
            radio.setChecked(key == options.get("format", "both"))
            item = QVBoxLayout()
            item.addWidget(radio)
            item.addWidget(label(hint, "hint"))
            grid.addLayout(item, index // 2, index % 2)
        if self.formats.checkedId() < 0:
            self.formats.button(0).setChecked(True)
        formats.body.addLayout(grid)
        layout.addWidget(formats)
        sections = Card("Report content")
        tools = QHBoxLayout()
        self.section_count = label("", "hint")
        tools.addWidget(self.section_count, 1)
        tools.addWidget(button("Select all", lambda: self._select(True), "textButton"))
        tools.addSpacing(12)
        tools.addWidget(button("Clear", lambda: self._select(False), "textButton"))
        sections.body.addLayout(tools)
        selected = options.get("sections", [key for key, _ in SECTIONS])
        self.section_checks: dict[str, QCheckBox] = {}
        self.section_grid = QGridLayout()
        self.section_grid.setHorizontalSpacing(20)
        for index, (key, title) in enumerate(SECTIONS):
            check = QCheckBox(title.replace("&", "&&"))
            check.setChecked(key in selected)
            check.toggled.connect(self._update_count)
            self.section_checks[key] = check
            self.section_grid.addWidget(check, index // 2, index % 2)
        sections.body.addLayout(self.section_grid)
        layout.addWidget(sections)
        layout.addStretch()
        footer = QFrame()
        footer.setObjectName("footer")
        footer_layout = QVBoxLayout(footer)
        footer_layout.setContentsMargins(24, 12, 24, 16)
        self.validation = label("", "hint", wrap=True)
        footer_layout.addWidget(self.validation)
        actions = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        actions.button(QDialogButtonBox.StandardButton.Save).setText("Save report settings")
        actions.button(QDialogButtonBox.StandardButton.Save).setObjectName("primary")
        actions.accepted.connect(self.accept)
        actions.rejected.connect(self.reject)
        footer_layout.addWidget(actions)
        outer.addWidget(footer)
        self._update_count()

    def resizeEvent(self, event: QResizeEvent) -> None:
        """Keep long section names readable on small desktop screens."""
        super().resizeEvent(event)
        if hasattr(self, "section_grid"):
            columns = 2 if self.width() >= 650 else 1
            for index, check in enumerate(self.section_checks.values()):
                self.section_grid.removeWidget(check)
                self.section_grid.addWidget(check, index // columns, index % columns)

    def _select(self, checked: bool) -> None:
        for check in self.section_checks.values():
            check.setChecked(checked)

    def _update_count(self) -> None:
        self.section_count.setText(
            f"{sum(check.isChecked() for check in self.section_checks.values())} of {len(SECTIONS)} sections"
        )
        if hasattr(self, "validation"):
            self.validation.clear()

    def accept(self) -> None:
        """Save choices only after at least one content section is selected."""
        selected = [key for key, check in self.section_checks.items() if check.isChecked()]
        if not selected:
            self.validation.setText("Select at least one section to include in your report.")
            return
        self.selection = {
            "format": self.format_keys[self.formats.checkedId()],
            "sections": selected,
        }
        super().accept()


class EvidenceWindow(QMainWindow):
    """A responsive desktop interface backed by immutable, independently saved runs."""

    def __init__(self, settings_path: str | Path | None = None) -> None:
        """Build workflows and restore the existing user's preferences."""
        super().__init__()
        self.setWindowTitle("EDF · Evidence workspace")
        self.resize(1160, 720)
        self.setMinimumSize(820, 600)
        self.settings_path = Path(settings_path or Path.home() / ".edf_collector/config.json")
        self._loading = True
        self._busy = False
        self._closing = False
        self._wide: bool | None = None
        self.worker: RunWorker | None = None
        self.current_job: DesktopJob | None = None
        self.last_outcome: RunOutcome | None = None
        self._pending_outcome: RunOutcome | None = None
        self.report_options: dict = {"format": "both", "sections": [key for key, _ in SECTIONS]}
        self.checks: dict[str, QCheckBox] = {}
        self.fields: dict[str, QLineEdit | QDoubleSpinBox] = {}
        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(600)
        self._save_timer.timeout.connect(self.save_settings)
        self._build_ui()
        self.load_settings()
        self._loading = False
        self._refresh_state()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        shell = QHBoxLayout(central)
        shell.setContentsMargins(0, 0, 0, 0)
        shell.setSpacing(0)
        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(196)
        navigation = QVBoxLayout(sidebar)
        navigation.setContentsMargins(18, 30, 18, 22)
        navigation.setSpacing(8)
        navigation.addWidget(label("EDF", "eyebrow"))
        navigation.addWidget(label("Evidence\nworkspace", "brand"))
        navigation.addSpacing(32)
        self.nav_buttons = []
        for index, text in enumerate(["Build evidence pack", "Generate reports", "Preferences"]):
            action = button(text, lambda checked=False, index=index: self._navigate(index), "nav")
            action.setCheckable(True)
            self.nav_buttons.append(action)
            navigation.addWidget(action)
        navigation.addStretch()
        navigation.addWidget(label("LOCAL WORKSPACE", "eyebrow"))
        navigation.addWidget(
            label("Your source files and\nexports stay on your computer.", "hint", wrap=True)
        )
        shell.addWidget(sidebar)
        workspace = QVBoxLayout()
        workspace.setContentsMargins(0, 0, 0, 0)
        workspace.setSpacing(0)
        self.pages = QStackedWidget()
        self.pages.addWidget(self._build_sources())
        self.pages.addWidget(self._build_reports())
        self.pages.addWidget(self._build_preferences())
        workspace.addWidget(self.pages, 1)
        workspace.addWidget(self._build_footer())
        shell.addLayout(workspace, 1)
        for field in self.fields.values():
            if isinstance(field, QLineEdit):
                field.textChanged.connect(self._preferences_changed)
            else:
                field.valueChanged.connect(self._preferences_changed)
        for check in self.checks.values():
            check.toggled.connect(self._preferences_changed)
        for source_field in self.sources.values():
            source_field.entry.textChanged.connect(self._refresh_state)
        self.report_input.entry.textChanged.connect(self._refresh_state)
        self.include_reports.toggled.connect(self._preferences_changed)
        self.nav_buttons[0].setChecked(True)
        self._adapt_layout()

    def _build_sources(self) -> QScrollArea:
        page, layout = scroll_page(
            "Build your evidence pack",
            "Bring together bills, account history and email attachments.",
        )
        self.source_card = Card(
            "01  Choose your sources",
            "Add one or more sources. We’ll organise them into a single pack.",
        )
        self.sources: dict[str, PathField] = {}
        for key, title, hint in [
            ("pdf_dir", "PDF bills", "Folder · PDF"),
            ("htm_path", "Account history", "File · HTM / HTML"),
            ("pst_path", "Outlook archive", "File · PST / OST"),
            ("eml_msg_dir", "Email files", "Folder · EML / MSG"),
        ]:
            field = PathField(title, hint, lambda key=key: self._pick_source(key))
            self.sources[key] = field
            self.source_card.body.addWidget(field)
        self.source_card.body.addStretch()
        self.output_card = Card(
            "02  Save & export", "An Excel workbook is included in every evidence pack."
        )
        destination = PathField("Save to", "Output folder", self._pick_output)
        destination.entry.setPlaceholderText("Beside the first source")
        self.fields["output_folder"] = destination.entry
        self.output_card.body.addWidget(destination)
        self.output_card.body.addWidget(label("Pack name", "fieldLabel"))
        name = QLineEdit("EDF_Dispute_Evidence")
        name.setAccessibleName("Evidence pack name")
        self.fields["output_name"] = name
        self.output_card.body.addWidget(name)
        self.output_card.body.addWidget(
            self._check("save_evidence_files", "Include original source files")
        )
        self.include_reports = QCheckBox("Include PDF, Word or HTML reports")
        self.output_card.body.addWidget(self.include_reports)
        self.report_summary = label("", "hint", wrap=True)
        self.output_card.body.addWidget(self.report_summary)
        self.report_settings_button = button("Report settings…", self._open_report_settings)
        self.output_card.body.addWidget(
            self.report_settings_button, alignment=Qt.AlignmentFlag.AlignLeft
        )
        self.output_card.body.addStretch()
        self.output_card.body.addWidget(
            label("Each run gets a new folder. Earlier exports are kept.", "hint", wrap=True)
        )
        self.source_grid = QGridLayout()
        self.source_grid.setSpacing(18)
        layout.addLayout(self.source_grid)
        layout.addStretch()
        return page

    def _build_reports(self) -> QScrollArea:
        page, layout = scroll_page(
            "Generate reports", "Turn saved analysis into reports without repeating extraction."
        )
        card = Card(
            "Choose a saved analysis",
            "Use the saved analysis from an earlier run to keep its original calculations and evidence.",
        )
        self.report_input = PathField(
            "Analysis or workbook", "File · JSON / Excel", self._pick_report_input
        )
        self.report_input.entry.setPlaceholderText(
            "Choose result.json from an earlier evidence pack"
        )
        card.body.addWidget(self.report_input)
        card.body.addWidget(
            label(
                "You can also import an Excel workbook. Workbook imports have limited provenance because original source files and decisions are unavailable.",
                "hint",
                wrap=True,
            )
        )
        layout.addWidget(card)
        options = Card("Save your reports")
        destination = PathField("Save to", "Output folder", self._pick_output)
        destination.entry.setPlaceholderText("Same location as the selected analysis")
        main_destination = cast(QLineEdit, self.fields["output_folder"])
        destination.entry.textChanged.connect(main_destination.setText)
        main_destination.textChanged.connect(destination.entry.setText)
        options.body.addWidget(destination)
        options.body.addWidget(
            button("Choose formats & sections…", self._open_report_settings),
            alignment=Qt.AlignmentFlag.AlignLeft,
        )
        self.saved_report_summary = label("", "hint", wrap=True)
        options.body.addWidget(self.saved_report_summary)
        self.report_excel = QCheckBox("Also create an Excel workbook")
        self.report_excel.toggled.connect(self._refresh_state)
        options.body.addWidget(self.report_excel)
        options.body.addWidget(
            label(
                "Reports are saved in a new folder alongside a preserved analysis.",
                "hint",
                wrap=True,
            )
        )
        layout.addWidget(options)
        layout.addStretch()
        return page

    def _check(self, key: str, text: str) -> QCheckBox:
        check = QCheckBox(text)
        check.setChecked(DEFAULTS[key])
        self.checks[key] = check
        return check

    def _text_field(self, parent: Card, key: str, title: str, default: str = "") -> QLineEdit:
        parent.body.addWidget(label(title, "fieldLabel"))
        field = QLineEdit(default)
        field.setAccessibleName(title)
        self.fields[key] = field
        parent.body.addWidget(field)
        return field

    def _amount_field(self, parent: Card, key: str, title: str, value: float) -> None:
        parent.body.addWidget(label(title, "fieldLabel"))
        field = QDoubleSpinBox()
        field.setRange(0, 1_000_000_000)
        field.setDecimals(2)
        field.setPrefix("£ ")
        field.setValue(value)
        field.setAccessibleName(title)
        self.fields[key] = field
        parent.body.addWidget(field)

    def _build_preferences(self) -> QScrollArea:
        page, layout = scroll_page(
            "Preferences", "Fine-tune processing. These choices apply to your next extraction."
        )
        extraction = Card("Bill extraction")
        for key, title in [
            ("use_anchors", "Recognise common bill patterns"),
            ("use_large", "Use amount fallback when bill fields are missing"),
            ("use_reading_classification", "Identify actual and estimated readings"),
            ("use_pdf_fields", "Extract usage, rates and invoice numbers"),
        ]:
            extraction.body.addWidget(self._check(key, title))
        layout.addWidget(extraction)
        filters = Card(
            "Accounts & filters",
            "Filtering can exclude evidence. Kept side sheets make those decisions reviewable.",
        )
        filters.body.addWidget(self._check("use_acc_filter", "Only include a specific account"))
        self._text_field(filters, "acc_num", "Account number")
        filters.body.addWidget(self._check("use_domain_filter", "Filter emails by sender"))
        self._text_field(
            filters, "domain_filter", "Sender domains or addresses", "edfenergy.com"
        ).setPlaceholderText("Separate multiple values with commas")
        filters.body.addWidget(
            self._check("filter_below", "Exclude results below a minimum amount")
        )
        self._amount_field(filters, "min_amount", "Minimum amount", 50)
        filters.body.addWidget(
            self._check("save_filtered", "Keep excluded records on a side sheet")
        )
        self._amount_field(filters, "analysis_min", "Minimum charge for bill analysis", 0)
        self._text_field(filters, "report_account_ref", "Account reference shown on reports")
        layout.addWidget(filters)
        duplicates = Card(
            "Duplicate evidence",
            "Supported identities establish duplicates. Matching amounts alone do not.",
        )
        for key, title in [
            ("use_dedup", "Remove supported duplicates across sources"),
            ("save_dups", "Keep removed duplicates on a side sheet"),
            ("amalgamate_duplicates", "Merge missing fields from duplicate records"),
        ]:
            duplicates.body.addWidget(self._check(key, title))
        layout.addWidget(duplicates)
        sap = Card(
            "SAP records & reconciliation",
            "Include contract, meter and financial history supplied by EDF.",
        )
        sap.body.addWidget(self._check("scan_sap_dumps", "Read SAP records in PDF files"))
        sap.body.addWidget(
            self._check("generate_reconciliation_sheet", "Compare SAP records with bill evidence")
        )
        layout.addWidget(sap)
        layout.addStretch()
        return page

    def _build_footer(self) -> QFrame:
        footer = QFrame()
        footer.setObjectName("footer")
        layout = QVBoxLayout(footer)
        layout.setContentsMargins(28, 12, 28, 16)
        layout.setSpacing(10)
        self.feedback = label("", "feedback", wrap=True)
        self.feedback.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.feedback.hide()
        layout.addWidget(self.feedback)
        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setRange(0, 100)
        self.progress.hide()
        layout.addWidget(self.progress)
        actions = QHBoxLayout()
        status = QVBoxLayout()
        status.setSpacing(4)
        self.status = label("Ready when you are", "status", wrap=True)
        self.status.setAccessibleName("Run status")
        self.readiness = label("Choose at least one source to get started.", "hint", wrap=True)
        status.addWidget(self.status)
        status.addWidget(self.readiness)
        actions.addLayout(status, 1)
        self.details_button = button("Details…", self._show_details)
        self.details_button.hide()
        actions.addWidget(self.details_button)
        self.open_folder_button = button("Open folder", self._open_folder)
        self.open_folder_button.hide()
        actions.addWidget(self.open_folder_button)
        self.run_button = button("Create evidence pack", self._start_or_cancel, "primary")
        self.run_button.setMinimumWidth(185)
        actions.addWidget(self.run_button)
        layout.addLayout(actions)
        return footer

    def _navigate(self, index: int) -> None:
        self.pages.setCurrentIndex(index)
        for selected, action in enumerate(self.nav_buttons):
            action.setChecked(selected == index)
        self._refresh_state()

    def _adapt_layout(self) -> None:
        wide = self.width() >= 1080
        if self._wide == wide:
            return
        self._wide = wide
        self.source_grid.removeWidget(self.source_card)
        self.source_grid.removeWidget(self.output_card)
        self.source_grid.addWidget(self.source_card, 0, 0)
        self.source_grid.addWidget(self.output_card, 0 if wide else 1, 1 if wide else 0)
        self.source_grid.setColumnStretch(0, 3 if wide else 1)
        self.source_grid.setColumnStretch(1, 2 if wide else 0)

    def resizeEvent(self, event: QResizeEvent) -> None:
        """Stack the two source steps when a narrow window needs more reading space."""
        super().resizeEvent(event)
        if hasattr(self, "source_grid"):
            self._adapt_layout()

    def _pick_source(self, key: str) -> None:
        if key in {"pdf_dir", "eml_msg_dir"}:
            path = QFileDialog.getExistingDirectory(
                self, "Choose " + ("PDF bills" if key == "pdf_dir" else "email files")
            )
        else:
            filters = (
                "Outlook archives (*.pst *.ost)"
                if key == "pst_path"
                else "Account history (*.htm *.html)"
            )
            path, _ = QFileDialog.getOpenFileName(self, "Choose source file", "", filters)
        if path:
            self.sources[key].entry.setText(path)

    def _pick_output(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self, "Choose output folder", str(self._field_value("output_folder"))
        )
        if path:
            cast(QLineEdit, self.fields["output_folder"]).setText(path)

    def _pick_report_input(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open saved analysis or workbook",
            "",
            "Saved analysis (*.json);;Excel workbook (*.xlsx)",
        )
        if path:
            self.report_input.entry.setText(path)

    def _open_report_settings(self) -> None:
        dialog = ReportDialog(self.report_options, self)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.selection is not None:
            self.report_options = dialog.selection
            self._preferences_changed()

    def _preferences_changed(self) -> None:
        if not self._loading:
            self._refresh_state()
            self._save_timer.start()

    def _field_value(self, key: str) -> str | float:
        field = self.fields[key]
        if isinstance(field, QDoubleSpinBox):
            field.interpretText()
            return field.value()
        return field.text().strip()

    def _refresh_state(self) -> None:
        if self._loading:
            return
        self.fields["acc_num"].setEnabled(self.checks["use_acc_filter"].isChecked())
        self.fields["domain_filter"].setEnabled(self.checks["use_domain_filter"].isChecked())
        self.fields["min_amount"].setEnabled(self.checks["filter_below"].isChecked())
        self.checks["save_filtered"].setEnabled(self.checks["filter_below"].isChecked())
        self.checks["save_dups"].setEnabled(self.checks["use_dedup"].isChecked())
        self.checks["amalgamate_duplicates"].setEnabled(
            self.checks["use_dedup"].isChecked() and self.checks["save_dups"].isChecked()
        )
        formats = FORMATS[self.report_options["format"]]
        section_count = len(self.report_options["sections"])
        self.report_summary.setText(
            " · ".join(fmt.upper() if fmt != "docx" else "Word" for fmt in formats)
            + f" · {section_count} section{'s' if section_count != 1 else ''}"
        )
        self.saved_report_summary.setText(self.report_summary.text())
        if self._busy:
            return
        page = self.pages.currentIndex()
        if page == 2:
            self.run_button.setText("Back to workspace")
            self.run_button.setEnabled(True)
            self.readiness.setText("Preferences are saved automatically.")
        elif page == 1:
            self.run_button.setText("Generate reports")
            self.run_button.setEnabled(bool(self.report_input.value()))
            self.readiness.setText(
                "Saved analysis keeps its original calculations."
                if self.report_input.value()
                else "Choose a saved analysis or workbook to continue."
            )
        else:
            count = sum(bool(field.value()) for field in self.sources.values())
            self.run_button.setText("Create evidence pack")
            self.run_button.setEnabled(count > 0)
            self.readiness.setText(
                f"{count} source{'s' if count != 1 else ''} selected · Excel"
                + (" + reports" if self.include_reports.isChecked() else " workbook")
                if count
                else "Choose at least one source to get started."
            )

    def load_settings(self) -> None:
        """Reuse existing settings; incomplete or malformed files retain safe defaults."""
        try:
            payload = json.loads(self.settings_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        if not isinstance(payload, dict) or not isinstance(payload.get("gui_state", {}), dict):
            return
        values = dict(payload.get("gui_state", {}))
        for alias, key in {
            "use_reading_class": "use_reading_classification",
            "use_acc_filt": "use_acc_filter",
        }.items():
            if key not in values and alias in values:
                values[key] = values[alias]
        for key, check in self.checks.items():
            if type(values.get(key)) is bool:
                check.setChecked(values[key])
        for key, field in self.fields.items():
            value = values.get(key, payload.get(key))
            if (
                isinstance(field, QDoubleSpinBox)
                and isinstance(value, (int, float))
                and not isinstance(value, bool)
            ):
                try:
                    if math.isfinite(value):
                        field.setValue(value)
                except OverflowError:
                    pass
            elif isinstance(field, QLineEdit) and isinstance(value, str):
                field.setText(value.removesuffix(".xlsx") if key == "output_name" else value)
        self.include_reports.setChecked(values.get("auto_generate_report") is True)
        options = payload.get("report_options")
        if (
            isinstance(options, dict)
            and isinstance(options.get("format"), str)
            and options["format"] in FORMATS
        ):
            selected = options.get("sections", [key for key, _ in SECTIONS])
            if (
                isinstance(selected, list)
                and selected
                and all(isinstance(key, str) and key in dict(SECTIONS) for key in selected)
            ):
                self.report_options = {
                    "format": options["format"],
                    "sections": list(dict.fromkeys(selected)),
                }

    def _configuration(self) -> dict:
        config: dict = {key: check.isChecked() for key, check in self.checks.items()}
        config.update(
            {
                key: self._field_value(key)
                for key in self.fields
                if key not in {"output_name", "output_folder"}
            }
        )
        config["as_of"] = date.today().isoformat()
        config["report_sections"] = list(self.report_options["sections"])
        return config

    def save_settings(self) -> None:
        """Publish preferences atomically without retaining selected source paths."""
        values = self._configuration()
        values.pop("as_of")
        values.pop("report_sections")
        values.update(
            output_folder=self._field_value("output_folder"),
            output_name=str(self._field_value("output_name")) + ".xlsx",
            auto_generate_report=self.include_reports.isChecked(),
        )
        try:
            self.settings_path.parent.mkdir(parents=True, exist_ok=True)
            write_json(
                {
                    "gui_state": values,
                    "output_folder": values["output_folder"],
                    "report_options": self.report_options,
                },
                str(self.settings_path),
            )
            os.chmod(self.settings_path, 0o600)
        except OSError as error:
            self._feedback("Preferences could not be saved: " + str(error), "warning")

    def capture_job(self) -> DesktopJob:
        """Validate local inputs and reserve an isolated directory before starting work."""
        config = self._configuration()
        report_mode = self.pages.currentIndex() == 1
        inputs = []
        if report_mode:
            path = Path(self.report_input.value())
            if not path.is_file() or path.suffix.lower() not in {".json", ".xlsx"}:
                raise ValueError(
                    "Choose an existing saved analysis (.json) or Excel workbook (.xlsx)."
                )
            inputs.append(SourceSelection("records_json", str(path)))
            mode = "workbook" if path.suffix.lower() == ".xlsx" else "saved_result"
            formats = FORMATS[self.report_options["format"]]
            if self.report_excel.isChecked():
                formats = ("xlsx", *formats)
            stem = path.stem
        else:
            mode = "extract"
            for key, field in self.sources.items():
                value = field.value()
                if not value:
                    continue
                path = Path(value)
                folder = key in {"pdf_dir", "eml_msg_dir"}
                if not (path.is_dir() if folder else path.is_file()):
                    raise ValueError(
                        "Source not found. Choose an existing "
                        + ("folder: " if folder else "file: ")
                        + value
                    )
                if key == "eml_msg_dir":
                    inputs.extend(SourceSelection(kind, value) for kind in ("eml_dir", "msg_dir"))
                else:
                    kind = {
                        "pdf_dir": "pdf_dir",
                        "htm_path": "htm",
                        "pst_path": "ost" if path.suffix.lower() == ".ost" else "pst",
                    }[key]
                    inputs.append(SourceSelection(kind, value))
            if not inputs:
                raise ValueError("Add at least one source before creating your evidence pack.")
            if config["use_acc_filter"] and not config["acc_num"]:
                raise ValueError(
                    "Enter an account number in Preferences, or switch off the account filter."
                )
            if (
                config["use_domain_filter"]
                and any(source.kind in {"pst", "ost", "eml_dir", "msg_dir"} for source in inputs)
                and not config["domain_filter"]
            ):
                raise ValueError(
                    "Enter a sender domain in Preferences, or switch off the sender filter."
                )
            stem = str(self._field_value("output_name"))
            formats = (
                ("xlsx", *FORMATS[self.report_options["format"]])
                if self.include_reports.isChecked()
                else ("xlsx",)
            )
        if (
            not stem
            or stem in {".", ".."}
            or any(char in stem for char in '/\\<>:"|?*')
            or any(ord(char) < 32 for char in stem)
            or len(stem) > 180
            or stem.endswith((".", " "))
        ):
            raise ValueError("Use a pack name without folder paths or special filename characters.")
        reserved = {
            "CON",
            "PRN",
            "AUX",
            "NUL",
            *[f"{prefix}{number}" for prefix in ("COM", "LPT") for number in range(1, 10)],
        }
        if stem.split(".")[0].upper() in reserved:
            raise ValueError("Choose a different pack name; this name is reserved by Windows.")
        destination = str(self._field_value("output_folder"))
        first = Path(inputs[0].path)
        base = Path(destination) if destination else first.parent
        base.mkdir(parents=True, exist_ok=True)
        run_directory = tempfile.mkdtemp(
            prefix=stem + "_" + date.today().strftime("%Y%m%d") + "_", dir=base
        )
        request = RunRequest(
            uuid4().hex,
            datetime.now(timezone.utc),
            date.fromisoformat(config["as_of"]),
            tuple(inputs),
            freeze_json(config),
        )
        return DesktopJob(
            request,
            run_directory,
            stem,
            tuple(formats),
            tuple(self.report_options["sections"]),
            mode,
            inputs[0].path if report_mode else "",
        )

    def _start_or_cancel(self) -> None:
        if self._busy:
            if self.worker is not None:
                self.worker.cancel_event.set()
            self.run_button.setText("Cancelling…")
            self.run_button.setEnabled(False)
            self.status.setText("Finishing the current step…")
            return
        if self.pages.currentIndex() == 2:
            self._navigate(0)
            return
        try:
            job = self.capture_job()
        except (ValueError, OSError) as error:
            self._feedback(str(error), "error")
            return
        self.save_settings()
        self.current_job = job
        self.last_outcome = None
        self._pending_outcome = None
        self._busy = True
        self.feedback.hide()
        self.details_button.hide()
        self.open_folder_button.hide()
        self.pages.setEnabled(False)
        for action in self.nav_buttons:
            action.setEnabled(False)
        self.progress.setRange(0, 0)
        self.progress.show()
        self.status.setText("Starting your run…")
        self.readiness.setText("You can cancel after the current step completes.")
        self.run_button.setText("Cancel run")
        self.run_button.setEnabled(True)
        worker = RunWorker(job, self)
        self.worker = worker
        worker.progress.connect(self._on_progress)
        worker.percentage.connect(self._on_percentage)
        worker.outcome_ready.connect(self._receive_outcome)
        worker.finished.connect(self._worker_finished)
        worker.start()

    @Slot(str)
    def _on_progress(self, text: str) -> None:
        self.status.setText(text)
        self.progress.setRange(0, 0)

    @Slot(int)
    def _on_percentage(self, value: int) -> None:
        self.progress.setRange(0, 100)
        self.progress.setValue(value)

    @Slot(object)
    def _receive_outcome(self, outcome: RunOutcome) -> None:
        self._pending_outcome = outcome

    @Slot()
    def _worker_finished(self) -> None:
        worker = self.worker
        self.worker = None
        self._busy = False
        self.pages.setEnabled(True)
        for action in self.nav_buttons:
            action.setEnabled(True)
        self.progress.setRange(0, 100)
        self.progress.hide()
        self._refresh_state()
        if self._pending_outcome is not None:
            self.last_outcome = self._pending_outcome
            self._present_outcome(self.last_outcome)
        if worker is not None:
            worker.deleteLater()
        if self._closing:
            self.close()

    def _feedback(self, text: str, severity: str) -> None:
        self.feedback.setText(text)
        self.feedback.setProperty("severity", severity)
        self.feedback.style().unpolish(self.feedback)
        self.feedback.style().polish(self.feedback)
        self.feedback.show()

    def _present_outcome(self, outcome: RunOutcome) -> None:
        manifest = object_value(outcome.manifest)
        files = len(manifest.get("artifacts", {}))
        records = len(outcome.snapshot.prepared.extraction.observations) if outcome.snapshot else 0
        titles = {
            "complete": "Your files are ready",
            "review_required": "Created with review notes",
            "failed": "Run needs attention",
            "empty": "No billing records found",
            "cancelled": "Run cancelled",
        }
        self.status.setText(titles.get(outcome.status, "Run finished"))
        self.readiness.setText(
            f"{records} original records · {files} exported file{'s' if files != 1 else ''}"
        )
        if outcome.status == "complete":
            self._feedback(
                "Your files and saved analysis are ready in the output folder.", "success"
            )
        elif outcome.status == "review_required":
            self._feedback(
                "Files were saved, but some evidence needs review. Read the details before sharing reports.",
                "warning",
            )
        elif outcome.status == "cancelled":
            self._feedback(
                f"{files} exported files were saved before cancellation."
                if files
                else "The run was cancelled. No reports were exported.",
                "warning",
            )
        elif outcome.status == "empty":
            self._feedback("Check your selected sources and filters, then try again.", "warning")
        else:
            retained = (
                " Your saved analysis is available in the output folder."
                if manifest.get("result")
                else ""
            )
            self._feedback((outcome.error or "The run could not finish.")[:350] + retained, "error")
        self.details_button.setVisible(outcome.status != "complete")
        self.open_folder_button.setVisible(
            self.current_job is not None and Path(self.current_job.directory).is_dir()
        )

    def _open_folder(self) -> None:
        if self.current_job is not None and not QDesktopServices.openUrl(
            QUrl.fromLocalFile(self.current_job.directory)
        ):
            self._feedback(
                "The folder could not be opened. Your files are saved at "
                + self.current_job.directory,
                "warning",
            )

    def _show_details(self) -> None:
        if self.last_outcome is None:
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Run details")
        dialog.resize(640, 420)
        layout = QVBoxLayout(dialog)
        layout.addWidget(label(self.status.text(), "sectionTitle"))
        details = QPlainTextEdit()
        details.setReadOnly(True)
        lines = [self.last_outcome.error or ""]
        lines.extend(object_value(self.last_outcome.manifest).get("warnings", []))
        if self.last_outcome.snapshot is not None:
            state = object_value(self.last_outcome.snapshot.prepared.extraction.legacy_state)
            lines.extend(str(note) for note in state.get("error_log", []))
        if self.current_job is not None:
            lines.append("Output folder: " + self.current_job.directory)
        details.setPlainText("\n\n".join(line for line in lines if line))
        layout.addWidget(details)
        actions = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        actions.rejected.connect(dialog.reject)
        layout.addWidget(actions)
        dialog.exec()

    def closeEvent(self, event: QCloseEvent) -> None:
        """Keep the worker alive until cancellation and cleanup finish."""
        if self.worker is not None and self.worker.isRunning():
            self._closing = True
            self._start_or_cancel()
            event.ignore()
            return
        self._save_timer.stop()
        self.save_settings()
        event.accept()


def launch(argv: list[str] | None = None) -> int:
    """Launch the Qt desktop; CLI paths never import the GUI toolkit."""
    existing = QApplication.instance()
    application = existing if isinstance(existing, QApplication) else QApplication(argv or sys.argv)
    apply_theme(application)
    application.setApplicationName("EDF Evidence Collector")
    window = EvidenceWindow()
    window.show()
    return application.exec()
