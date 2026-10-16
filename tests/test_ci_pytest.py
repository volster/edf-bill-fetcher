"""The CI runner must preserve failures and their evidence without suite retries."""

import json
from pathlib import Path

import pytest

import ci_pytest


@pytest.mark.parametrize("code", [0, 1, 2, 3, 4, 5])
def test_exit_codes_and_failure_reports_are_retained_without_rerunning(tmp_path, monkeypatch, code):
    monkeypatch.chdir(tmp_path)
    calls = []
    stale = Path(".ci-diagnostics/tests/attempt-2")
    stale.mkdir(parents=True)
    (stale / "test-results.xml").write_text("old results")
    Path("test-results.xml").write_text("old results")
    Path("coverage.xml").write_text("old coverage")

    def execute(command):
        assert not Path("test-results.xml").exists()
        assert not Path("coverage.xml").exists()
        calls.append(command)
        Path("test-results.xml").write_text("current failure or success")
        Path("coverage.xml").write_text("current coverage")
        return code

    monkeypatch.setattr(ci_pytest.subprocess, "call", execute)
    assert ci_pytest.main() == code
    assert len(calls) == 1
    assert calls[0][:3] == [ci_pytest.sys.executable, "-m", "pytest"]
    diagnostics = Path(".ci-diagnostics/tests")
    assert not stale.exists()
    assert (diagnostics / "attempt-1/test-results.xml").read_text() == "current failure or success"
    assert (diagnostics / "attempt-1/coverage.xml").read_text() == "current coverage"
    assert json.loads((diagnostics / "attempts.json").read_text()) == [
        {"attempt": 1, "exit_code": code}
    ]


def test_aborted_suite_cannot_publish_stale_results(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    Path("test-results.xml").write_text("old success")
    Path("coverage.xml").write_text("old coverage")
    monkeypatch.setattr(ci_pytest.subprocess, "call", lambda command: 2)
    assert ci_pytest.main() == 2
    assert not Path("test-results.xml").exists()
    assert not Path("coverage.xml").exists()
    assert not Path(".ci-diagnostics/tests/attempt-1/test-results.xml").exists()


def test_native_recovery_is_published_after_the_suite_without_changing_its_failure(
    tmp_path, monkeypatch, capsys
):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("GITHUB_STEP_SUMMARY", str(tmp_path / "summary.md"))

    def execute(command):
        Path(".ci-diagnostics/tk-startup.jsonl").write_text('{"recovered": true}\n')
        return 1

    monkeypatch.setattr(ci_pytest.subprocess, "call", execute)
    assert ci_pytest.main() == 1
    assert (
        "::warning title=Windows Tcl/Tk startup recovered::1 native startup"
        in capsys.readouterr().out
    )
    assert "The full suite ran once" in Path("summary.md").read_text()


def test_malformed_native_diagnostics_do_not_replace_the_test_exit_code(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    def execute(command):
        Path(".ci-diagnostics/tk-startup.jsonl").write_text("partial JSON\n")
        return 2

    monkeypatch.setattr(ci_pytest.subprocess, "call", execute)
    assert ci_pytest.main() == 2
