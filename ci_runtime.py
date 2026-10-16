"""Record the selected CI runtime and exercise its real Tcl/Tk and Qt libraries."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import sys
import tempfile
import traceback
from pathlib import Path


def runtime_info() -> dict:
    """Collect an allowlist of runtime details without dumping environment secrets."""
    return {
        "executable": sys.executable,
        "version": sys.version,
        "prefix": sys.prefix,
        "platform": platform.platform(),
        "environment": {
            key: os.environ[key]
            for key in (
                "pythonLocation",
                "TCL_LIBRARY",
                "TK_LIBRARY",
                "QT_QPA_PLATFORM",
                "RUNNER_OS",
                "RUNNER_ARCH",
                "ImageOS",
                "ImageVersion",
            )
            if key in os.environ
        },
    }


def check_tk() -> dict:
    """Load Tcl/Tk repeatedly, including themed widgets, using the selected Python."""
    import tkinter as tk
    from tkinter import ttk

    libraries: dict[str, str] = {}
    for _ in range(20):
        root = tk.Tk()
        try:
            root.withdraw()
            libraries = {
                "tcl": str(root.tk.eval("info library")),
                "tk": str(root.tk.eval("set tk_library")),
                "patchlevel": str(root.tk.eval("info patchlevel")),
            }
            style = ttk.Style(root)
            for theme in style.theme_names():
                style.theme_use(theme)
                widget = ttk.Treeview(root)
                widget.insert("", "end", text="Runtime check")
                widget.pack()
                root.update_idletasks()
                widget.destroy()
        finally:
            root.destroy()
    # A manifest distinguishes an absent/unreadable file from a guessed scanner
    # race. Read actual bytes; enumerating FileInfo objects does not warm files.
    scripts = {}
    for name in ("tcl", "tk"):
        base = Path(libraries[name])
        scripts[name] = {
            str(path.relative_to(base)): hashlib.sha256(path.read_bytes()).hexdigest()
            for path in sorted(base.rglob("*.tcl"))
        }
    return {"libraries": libraries, "scripts_sha256": scripts, "startup_cycles": 20}


def check_qt() -> dict:
    """Construct and display the application's real desktop with isolated settings."""
    from PySide6.QtCore import QCoreApplication, QEvent, qVersion
    from PySide6.QtWidgets import QApplication

    from edf_bill_fetcher.ui.qt_app import EvidenceWindow
    from edf_bill_fetcher.ui.qt_theme import apply_theme

    application = QApplication.instance() or QApplication([])
    if not isinstance(application, QApplication):
        raise RuntimeError("A QWidget application could not be created")
    apply_theme(application)
    with tempfile.TemporaryDirectory(prefix="edf-ci-gui-") as directory:
        window = EvidenceWindow(Path(directory) / "settings.json")
        try:
            window.show()
            application.processEvents()
            if not window.isVisible() or window.pages.count() != 4:
                raise RuntimeError("The evidence workspace did not open completely")
        finally:
            window.close()
            window.deleteLater()
            QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
            application.processEvents()
    return {"version": qVersion(), "platform": application.platformName(), "workspace": "opened"}


def main(argv: list[str] | None = None) -> int:
    """Persist diagnostics even when a GUI runtime raises during startup."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--inventory", action="store_true", help="Record Python without loading GUIs"
    )
    parser.add_argument("--report", type=Path, default=Path(".ci-diagnostics/runtime.json"))
    args = parser.parse_args(argv)
    report = runtime_info()
    args.report.parent.mkdir(parents=True, exist_ok=True)

    def save() -> None:
        args.report.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")

    save()
    try:
        expected = os.environ.get("pythonLocation")
        if expected and Path(sys.executable).resolve().parent != Path(expected).resolve():
            raise RuntimeError("Selected Python does not match setup-python's pythonLocation")
        if not args.inventory:
            report["tk"] = check_tk()
            save()
            report["qt"] = check_qt()
        report["status"] = "success"
        return 0
    except Exception:
        report["status"] = "failure"
        report["error"] = traceback.format_exc()
        traceback.print_exc()
        return 1
    finally:
        save()


if __name__ == "__main__":
    raise SystemExit(main())
