"""Run CI tests, retrying only transient Windows Tcl/Tk fixture startup errors."""

from __future__ import annotations

import json
import logging
import os
import shutil
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
    """Keep each attempt's evidence and publish one complete suite as the final gate."""
    junit = Path("test-results.xml")
    coverage = Path("coverage.xml")
    diagnostics = Path(".ci-diagnostics/tests")
    if diagnostics.exists():
        shutil.rmtree(diagnostics)
    diagnostics.mkdir(parents=True, exist_ok=True)
    attempts: list[dict] = []
    junit.unlink(missing_ok=True)
    coverage.unlink(missing_ok=True)
    command = [
        sys.executable,
        "-m",
        "pytest",
        "--cov=.",
        "--cov-report=xml",
        "--cov-report=term-missing",
        "--junitxml=" + str(junit),
    ]

    def run() -> int:
        code = subprocess.call(command)
        attempt = len(attempts) + 1
        destination = diagnostics / f"attempt-{attempt}"
        destination.mkdir(exist_ok=True)
        for report in (junit, coverage):
            if report.is_file():
                shutil.copyfile(report, destination / report.name)
        attempts.append({"attempt": attempt, "exit_code": code})
        (diagnostics / "attempts.json").write_text(
            json.dumps(attempts, indent=2) + "\n", encoding="utf-8"
        )
        return code

    code = run()
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
    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a", encoding="utf-8") as handle:
            handle.write(
                "### Windows Tcl/Tk infrastructure retry\n\n"
                "The first full suite failed solely during native library setup. "
                "Both attempts' JUnit and coverage reports are retained in the diagnostics artifact. "
                "The second full suite determines this job's result.\n\n"
            )
    # Do not append coverage or select only failed tests: the published XML,
    # coverage gate and final exit code must describe a complete suite.
    junit.unlink(missing_ok=True)
    coverage.unlink(missing_ok=True)
    return run()


if __name__ == "__main__":
    raise SystemExit(main())
