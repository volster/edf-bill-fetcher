"""Run CI tests, retrying only transient Windows Tcl/Tk fixture startup errors."""

from __future__ import annotations

import logging
import os
import subprocess
import sys
from pathlib import Path
from xml.etree import ElementTree


def _only_tk_setup_errors(path: Path) -> bool:
    """Recognize the runner's missing Tk library without accepting test failures."""
    try:
        root = ElementTree.parse(path).getroot()
    except (OSError, ElementTree.ParseError):
        return False
    errors = list(root.iter("error"))
    if not errors or root.find(".//failure") is not None:
        return False
    for error in errors:
        message = error.get("message", "")
        detail = message + " " + " ".join(error.itertext())
        if not message.startswith("failed on setup") or "_tkinter.TclError" not in detail:
            return False
        if not any(
            marker in detail
            for marker in ("Can't find a usable tk.tcl", "Can't find a usable init.tcl")
        ):
            return False
    return True


def main() -> int:
    """Run every test and replace reports with one complete retry when eligible."""
    junit = Path("test-results.xml")
    junit.unlink(missing_ok=True)
    command = [
        sys.executable,
        "-m",
        "pytest",
        "--cov=.",
        "--cov-report=xml",
        "--cov-report=term-missing",
        "--junitxml=" + str(junit),
    ]
    code = subprocess.call(command)
    if (
        code != 1
        or sys.platform != "win32"
        or os.environ.get("GITHUB_ACTIONS") != "true"
        or not _only_tk_setup_errors(junit)
    ):
        return code
    logging.warning(
        "Windows runner could not load Tcl/Tk during fixture setup. "
        "Retrying the entire suite once; ordinary failures and repeated startup errors remain fatal."
    )
    # Do not append coverage or select only failed tests: the published XML,
    # coverage gate and final exit code must describe a complete suite.
    junit.unlink(missing_ok=True)
    return subprocess.call(command)


if __name__ == "__main__":
    raise SystemExit(main())
