"""Keep CI infrastructure retries bounded and independent of test correctness."""

import json
from pathlib import Path
from xml.etree import ElementTree

import pytest

import ci_pytest


def _junit(path, errors=(), failures=()):
    root = ElementTree.Element("testsuites")
    suite = ElementTree.SubElement(root, "testsuite")
    case = ElementTree.SubElement(suite, "testcase")
    for tag, messages in (("error", errors), ("failure", failures)):
        for message in messages:
            ElementTree.SubElement(case, tag, message=message).text = message
    ElementTree.ElementTree(root).write(path)


TK_ERROR = 'failed on setup with "_tkinter.TclError: Can\'t find a usable tk.tcl"'


@pytest.mark.parametrize(
    ("errors", "failures", "expected"),
    [
        ([TK_ERROR], [], True),
        ([TK_ERROR.replace("tk.tcl", "init.tcl")], [], True),
        ([TK_ERROR], ["AssertionError: wrong balance"], False),
        ([TK_ERROR, "failed on setup with OSError: disk full"], [], False),
        ([TK_ERROR.replace("setup", "teardown")], [], False),
        (["collection failure: " + TK_ERROR], [], False),
        (["failed on setup with _tkinter.TclError: unknown widget"], [], False),
        ([], [TK_ERROR], False),
        ([], [], False),
    ],
)
def test_only_native_library_fixture_errors_are_retryable(tmp_path, errors, failures, expected):
    report = tmp_path / "junit.xml"
    _junit(report, errors, failures)
    assert ci_pytest._only_tk_setup_errors(report) is expected


def test_missing_and_malformed_reports_never_qualify(tmp_path):
    report = tmp_path / "junit.xml"
    assert not ci_pytest._only_tk_setup_errors(report)
    report.write_text("not XML")
    assert not ci_pytest._only_tk_setup_errors(report)


@pytest.mark.parametrize("second_code", [0, 1])
def test_windows_ci_retries_entire_suite_once_and_keeps_final_result(
    tmp_path, monkeypatch, second_code
):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(ci_pytest.sys, "platform", "win32")
    monkeypatch.setenv("GITHUB_ACTIONS", "true")
    commands = []
    monkeypatch.setenv("GITHUB_STEP_SUMMARY", str(tmp_path / "summary.md"))

    def execute(command):
        assert not Path("test-results.xml").exists()
        commands.append(command.copy())
        Path("coverage.xml").write_text(f"coverage attempt {len(commands)}")
        _junit("test-results.xml", errors=[TK_ERROR] if len(commands) == 1 else [])
        return 1 if len(commands) == 1 else second_code

    monkeypatch.setattr(ci_pytest.subprocess, "call", execute)
    assert ci_pytest.main() == second_code
    assert len(commands) == 2
    assert commands[0] == commands[1]
    assert commands[0][1:3] == ["-m", "pytest"]
    assert "--cov-append" not in commands[0]
    assert "--lf" not in commands[0]
    diagnostics = tmp_path / ".ci-diagnostics/tests"
    assert ci_pytest._only_tk_setup_errors(diagnostics / "attempt-1/test-results.xml")
    assert not ci_pytest._only_tk_setup_errors(diagnostics / "attempt-2/test-results.xml")
    assert (diagnostics / "attempt-1/coverage.xml").read_text() == "coverage attempt 1"
    assert (diagnostics / "attempt-2/coverage.xml").read_text() == "coverage attempt 2"
    assert Path("coverage.xml").read_text() == "coverage attempt 2"
    assert json.loads((diagnostics / "attempts.json").read_text()) == [
        {"attempt": 1, "exit_code": 1},
        {"attempt": 2, "exit_code": second_code},
    ]
    assert "Both attempts" in (tmp_path / "summary.md").read_text()


def test_failed_retry_without_reports_does_not_publish_stale_first_attempt(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(ci_pytest.sys, "platform", "win32")
    monkeypatch.setenv("GITHUB_ACTIONS", "true")
    calls = []
    stale = Path(".ci-diagnostics/tests/attempt-2")
    stale.mkdir(parents=True)
    (stale / "test-results.xml").write_text("stale previous invocation")

    def execute(command):
        calls.append(command)
        if len(calls) == 1:
            _junit("test-results.xml", errors=[TK_ERROR])
            Path("coverage.xml").write_text("first attempt only")
            return 1
        assert not Path("coverage.xml").exists()
        return 2

    monkeypatch.setattr(ci_pytest.subprocess, "call", execute)
    assert ci_pytest.main() == 2
    assert not Path("test-results.xml").exists()
    assert not Path("coverage.xml").exists()
    assert Path(".ci-diagnostics/tests/attempt-1/test-results.xml").is_file()
    assert not Path(".ci-diagnostics/tests/attempt-2/test-results.xml").exists()


@pytest.mark.parametrize(
    ("platform", "actions", "code", "error"),
    [
        ("linux", "true", 1, TK_ERROR),
        ("win32", "false", 1, TK_ERROR),
        ("win32", "true", 2, TK_ERROR),
        ("win32", "true", 0, TK_ERROR),
        ("win32", "true", 1, "failed on setup with ValueError: invalid data"),
    ],
)
def test_other_failures_and_local_runs_are_never_retried(
    tmp_path, monkeypatch, platform, actions, code, error
):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(ci_pytest.sys, "platform", platform)
    monkeypatch.setenv("GITHUB_ACTIONS", actions)
    calls = []
    _junit("test-results.xml", errors=[TK_ERROR])

    def execute(command):
        assert not Path("test-results.xml").exists(), "A previous run's report must be removed"
        calls.append(command)
        _junit("test-results.xml", errors=[error])
        return code

    monkeypatch.setattr(ci_pytest.subprocess, "call", execute)
    assert ci_pytest.main() == code
    assert len(calls) == 1
