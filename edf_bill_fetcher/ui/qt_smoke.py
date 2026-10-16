"""Exercise the delivered desktop window without touching user preferences."""

from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from edf_bill_fetcher.ui.qt_app import EvidenceWindow
from edf_bill_fetcher.ui.qt_theme import apply_theme


def check_gui(report_path: Path, *, platform: str | None = None) -> int:
    """Show the real workspace, record its state after startup, and exit."""
    arguments = ["EDF GUI check"]
    if platform:
        arguments.extend(["-platform", platform])
    existing = QApplication.instance()
    application = existing if isinstance(existing, QApplication) else QApplication(arguments)
    apply_theme(application)
    application.setApplicationName("EDF Evidence Collector")
    with tempfile.TemporaryDirectory(prefix="edf-gui-check-") as directory:
        window = EvidenceWindow(Path(directory) / "preferences.json")

        def finish() -> None:
            try:
                report = {
                    "schema": 1,
                    "visible": window.isVisible(),
                    "pages": window.pages.count(),
                    "platform": application.platformName(),
                    "frozen": bool(getattr(sys, "frozen", False)),
                }
                report_path.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
                application.exit(0 if report["visible"] and report["pages"] == 4 else 1)
            except Exception:
                # Exceptions escaping a Qt callback do not reliably fail the process.
                application.exit(1)

        window.show()
        QTimer.singleShot(500, finish)
        try:
            return application.exec()
        finally:
            window.close()
            window.deleteLater()
