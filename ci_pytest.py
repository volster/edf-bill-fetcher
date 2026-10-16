"""Run the complete CI suite once and retain its reports and exit status."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from pathlib import Path


def _publish_tk_recoveries() -> None:
    """Publish outside pytest capture so recovered starts remain visible on GitHub."""
    report = Path(".ci-diagnostics/tk-startup.jsonl")
    if not report.is_file():
        return
    try:
        events = [json.loads(line) for line in report.read_text(encoding="utf-8").splitlines()]
        recovered = sum(event.get("recovered") is True for event in events)
    except (OSError, ValueError, AttributeError):
        sys.stderr.write("Could not summarize native Tk diagnostics; see the retained artifact.\n")
        return
    if recovered:
        sys.stdout.write(
            f"::warning title=Windows Tcl/Tk startup recovered::{recovered} native startup(s) "
            "recovered after verified script reads; no tests were rerun. "
            "See tk-startup.jsonl in the diagnostics artifact.\n"
        )
    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a", encoding="utf-8") as handle:
            handle.write(
                f"### Native Tcl/Tk startup\n\n{len(events)} startup error(s); "
                f"{recovered} recovered after one native-call retry. "
                "Script readability and hashes are retained in `tk-startup.jsonl`. "
                "The full suite ran once.\n\n"
            )


def main() -> int:
    """Preserve one full suite’s evidence and its original exit code."""
    junit = Path("test-results.xml")
    coverage = Path("coverage.xml")
    diagnostics = Path(".ci-diagnostics/tests")
    if diagnostics.exists():
        shutil.rmtree(diagnostics)
    diagnostics.mkdir(parents=True, exist_ok=True)
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

    code = subprocess.call(command)
    destination = diagnostics / "attempt-1"
    destination.mkdir()
    for report in (junit, coverage):
        if report.is_file():
            shutil.copyfile(report, destination / report.name)
    (diagnostics / "attempts.json").write_text(
        json.dumps([{"attempt": 1, "exit_code": code}], indent=2) + "\n", encoding="utf-8"
    )
    _publish_tk_recoveries()
    return code


if __name__ == "__main__":
    raise SystemExit(main())
