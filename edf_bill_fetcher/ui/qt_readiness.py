"""Show extraction prerequisites without blocking the desktop during native probes."""

from __future__ import annotations

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Signal, Slot
from shiboken6 import isValid

from edf_bill_fetcher.helpers.capabilities import Capability, check_capabilities
from edf_bill_fetcher.ui.qt_widgets import Card, WrappedCheckBox, button, label


class _ProbeSignals(QObject):
    finished = Signal(object)


class _Probe(QRunnable):
    def __init__(self) -> None:
        super().__init__()
        self.signals = _ProbeSignals()
        self.check = check_capabilities

    def run(self) -> None:
        try:
            checks = self.check()
        except Exception as error:
            checks = (Capability("runtime", "Installation checks", False, str(error)),)
        try:
            self.signals.finished.emit(checks)
        except RuntimeError:
            # QApplication cleanup can delete the signal while native checks
            # finish. A closed desktop has no recipient for these results.
            if isValid(self.signals):
                raise


class ReadinessCard(Card):
    """Require informed consent to limited extraction; absent archive readers are fatal."""

    changed = Signal()

    def __init__(self) -> None:
        """Start bounded checks in the thread pool, never on the widget thread."""
        super().__init__(
            "03  Extraction readiness", "Check support before processing your evidence."
        )
        self.checks: tuple[Capability, ...] | None = None
        self.checking = False
        self.sources: set[str] = set()
        self.summary = label("Checking extraction support…", "hint", wrap=True)
        self.body.addWidget(self.summary)
        self.limitations = label("", "feedback", wrap=True)
        self.body.addWidget(self.limitations)
        self.allow_limited = WrappedCheckBox("Continue with the limitations shown above")
        self.allow_limited.hide()
        self.allow_limited.toggled.connect(self.changed)
        self.body.addWidget(self.allow_limited)
        self.refresh_button = button("Refresh checks", self.refresh)
        self.body.addWidget(self.refresh_button)
        self.refresh()

    def refresh(self) -> None:
        """Clear previous consent when the available capabilities change."""
        if self.checking:
            return
        self.checking = True
        self.allow_limited.setChecked(False)
        self.refresh_button.setEnabled(False)
        self.summary.setText("Checking extraction support…")
        probe = _Probe()
        probe.signals.finished.connect(self._checked)
        QThreadPool.globalInstance().start(probe)
        self.changed.emit()

    @Slot(object)
    def _checked(self, checks: tuple[Capability, ...]) -> None:
        self.checks = checks
        self.checking = False
        self.refresh_button.setEnabled(True)
        self.update_sources(self.sources)
        self.changed.emit()

    def update_sources(self, sources: set[str]) -> None:
        """Explain limitations relevant to the selected sources without scanning their files."""
        if sources != self.sources:
            self.allow_limited.setChecked(False)
        self.sources = sources
        if self.checks is None or self.checking:
            return
        self.summary.setText(
            "\n".join(
                f"{'Ready' if check.available else 'Unavailable'}: {check.title}"
                for check in self.checks
            )
        )
        issues = self.issues()
        self.limitations.setText("\n".join(issues))
        self.limitations.setVisible(bool(issues))
        self.allow_limited.setVisible(bool(issues) and not self.blocking_issue())
        self.summary.setToolTip(
            "\n".join(f"{check.title}: {check.detail}" for check in self.checks)
        )

    def blocking_issue(self) -> str:
        """Never silently skip a selected PST archive or a required output dependency."""
        if self.checking or self.checks is None:
            return "Checking extraction support…"
        available = {check.key: check.available for check in self.checks}
        if not available.get("excel", False):
            return "Excel support is unavailable. Repair the installation before creating a pack."
        if not available.get("pdf", False) and self.sources & {
            "pdf_dir",
            "pst_path",
            "eml_msg_dir",
        }:
            return "PDF support is unavailable. Repair the installation before processing these sources."
        if "pst_path" in self.sources and not available.get("pst", False):
            return "Outlook archive support is unavailable. Install libpff-python or remove that source."
        return ""

    def issues(self) -> list[str]:
        """Keep source-specific consequences distinct from technical diagnostic details."""
        blocking = self.blocking_issue()
        if blocking:
            return [blocking]
        available = {check.key: check.available for check in self.checks or ()}
        issues = []
        if self.sources and not available.get("ocr", False):
            issues.append(
                "OCR is unavailable. Scanned pages and image-based account history cannot be read. Install Tesseract with English language data, or explicitly continue with this limitation."
            )
        if "eml_msg_dir" in self.sources and not available.get("msg", False):
            issues.append(
                "MSG support is unavailable. Outlook .msg files will not be read; .eml files remain supported. Install the [msg] extra, or explicitly continue with this limitation."
            )
        return issues

    def reason_not_ready(self) -> str:
        """Return the actionable explanation used by both the button and start guard."""
        blocking = self.blocking_issue()
        if blocking:
            return blocking
        if self.issues() and not self.allow_limited.isChecked():
            return "Review extraction limitations above before continuing."
        return ""
