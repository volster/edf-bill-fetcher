"""Complete typed configuration controls without an optional web dependency."""

from __future__ import annotations

from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QLabel,
    QLineEdit,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from edf_bill_fetcher.application.configuration import validate_settings
from edf_bill_fetcher.models.config import ConfigDict


class ConfigurationDialog(QDialog):
    """Expose every non-section analytical setting; report sections keep their existing dialog."""

    def __init__(self, values: dict, parent: QWidget | None = None) -> None:
        """Present canonical keys with explicit defaults and inline validation feedback."""
        super().__init__(parent)
        self.setWindowTitle("All analysis settings")
        self.resize(680, 720)
        self.result_config: dict = {}
        self.controls: dict[str, QLineEdit | QComboBox] = {}
        outer = QVBoxLayout(self)
        introduction = QLabel(
            "Blank fields use application defaults. These settings affect new analysis; accepted results keep their saved calculations."
        )
        introduction.setWordWrap(True)
        outer.addWidget(introduction)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        body = QWidget()
        form = QFormLayout(body)
        for key, annotation in ConfigDict.__annotations__.items():
            if key == "report_sections":
                continue
            control: QLineEdit | QComboBox
            if "bool" in str(annotation):
                control = QComboBox()
                control.addItems(("Application default", "Yes", "No"))
                control.setCurrentIndex(
                    1 if values.get(key) is True else 2 if values.get(key) is False else 0
                )
            else:
                control = QLineEdit(str(values.get(key, "")))
                control.setPlaceholderText(
                    "YYYY-MM-DD (blank = today)" if key == "as_of" else "Application default"
                )
            title = key.replace("_", " ").title()
            control.setAccessibleName(title)
            self.controls[key] = control
            form.addRow(title, control)
        scroll.setWidget(body)
        outer.addWidget(scroll)
        self.error = QLabel()
        self.error.setWordWrap(True)
        outer.addWidget(self.error)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        outer.addWidget(buttons)

    def accept(self) -> None:
        """Validate complete settings before committing any changes to desktop preferences."""
        values: dict = {}
        try:
            for key, control in self.controls.items():
                if isinstance(control, QComboBox):
                    if control.currentIndex():
                        values[key] = control.currentIndex() == 1
                elif text := control.text().strip():
                    annotation = str(ConfigDict.__annotations__[key])
                    values[key] = (
                        float(text)
                        if "float" in annotation
                        else int(text)
                        if "int" in annotation
                        else text
                    )
            validate_settings(values)
        except ValueError as error:
            self.error.setText(str(error))
            return
        self.result_config = values
        super().accept()
