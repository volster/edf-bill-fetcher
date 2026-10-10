"""Small reusable widgets for the evidence workspace."""

from collections.abc import Callable

from PySide6.QtCore import QPoint, QSize, Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSizePolicy,
    QStyle,
    QStyleOptionButton,
    QVBoxLayout,
    QWidget,
)


def label(text: str, role: str = "", *, wrap: bool = False) -> QLabel:
    """Create text with a semantic visual role and optional wrapping."""
    widget = QLabel(text)
    widget.setObjectName(role)
    widget.setWordWrap(wrap)
    widget.setTextFormat(Qt.TextFormat.PlainText)
    return widget


def button(text: str, action: Callable, role: str = "") -> QPushButton:
    """Create an action with a predictable keyboard focus target."""
    widget = QPushButton(text.replace("&", "&&"))
    widget.setObjectName(role)
    widget.setCursor(Qt.CursorShape.PointingHandCursor)
    widget.clicked.connect(lambda checked=False: action())
    return widget


class WrappedCheckBox(QCheckBox):
    """Keep long option labels readable without forcing a horizontal scrollbar."""

    def __init__(self, text: str) -> None:
        """Use a native indicator and a wrapping caption in one clickable control."""
        super().__init__()
        self.caption = label(text, wrap=True)
        self.caption.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.setAccessibleName(text)
        self.setToolTip(text)
        option = QStyleOptionButton()
        self.initStyleOption(option)
        indicator = self.style().pixelMetric(QStyle.PixelMetric.PM_IndicatorWidth, option, self)
        spacing = self.style().pixelMetric(QStyle.PixelMetric.PM_CheckBoxLabelSpacing, option, self)
        self.body = QHBoxLayout(self)
        self.body.setContentsMargins(indicator + spacing, 5, 0, 5)
        self.body.addWidget(self.caption)
        policy = QSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Preferred)
        policy.setHeightForWidth(True)
        self.setSizePolicy(policy)

    def text(self) -> str:
        """Return the full caption for callers inspecting this native checkbox."""
        return self.caption.text()

    def sizeHint(self) -> QSize:
        """Include the caption and native indicator in the preferred size."""
        return self.body.sizeHint()

    def minimumSizeHint(self) -> QSize:
        """Allow the caption to wrap instead of requiring its full line width."""
        return self.body.minimumSize()

    def heightForWidth(self, width: int) -> int:
        """Reserve enough height for every line at the actual available width."""
        return self.body.heightForWidth(width)

    def hitButton(self, position: QPoint) -> bool:
        """Clicking the caption toggles the same control as the indicator."""
        return self.rect().contains(position)


class Card(QFrame):
    """A titled surface that groups one step or closely related preferences."""

    def __init__(self, title: str, description: str = "") -> None:
        """Build a card with consistent padding and heading spacing."""
        super().__init__()
        self.setObjectName("card")
        self.body = QVBoxLayout(self)
        self.body.setContentsMargins(22, 20, 22, 20)
        self.body.setSpacing(12)
        self.body.addWidget(label(title, "sectionTitle"))
        if description:
            self.body.addWidget(label(description, "hint", wrap=True))


class PathField(QWidget):
    """An editable local path and clearly named browse action."""

    def __init__(self, title: str, hint: str, browse: Callable) -> None:
        """Connect the entire browse button to its source-specific picker."""
        super().__init__()
        self.setObjectName("row")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        heading = QHBoxLayout()
        heading.addWidget(label(title, "fieldLabel"))
        heading.addStretch()
        heading.addWidget(label(hint, "hint"))
        layout.addLayout(heading)
        row = QHBoxLayout()
        row.setSpacing(8)
        self.entry = QLineEdit()
        self.entry.setPlaceholderText("Choose a file or folder…")
        self.entry.setAccessibleName(title)
        self.entry.setToolTip(hint)
        self.browse_button = button("Browse…", browse)
        self.browse_button.setAccessibleName("Browse " + title)
        row.addWidget(self.entry, 1)
        row.addWidget(self.browse_button)
        layout.addLayout(row)

    def value(self) -> str:
        """Return the captured path without surrounding whitespace."""
        return self.entry.text().strip()
