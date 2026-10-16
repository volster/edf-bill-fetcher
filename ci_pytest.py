"""Run the complete CI suite once and retain its reports and exit status."""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path


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
    return code


if __name__ == "__main__":
    raise SystemExit(main())
