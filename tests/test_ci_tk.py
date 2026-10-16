"""Recover native startup only; genuine missing files and application errors stay fatal."""

import _tkinter
import hashlib
import json
import tkinter as tk

import pytest

import ci_tk


@pytest.fixture(autouse=True)
def isolated_actions(monkeypatch):
    monkeypatch.delenv("GITHUB_ACTIONS", raising=False)
    monkeypatch.setattr(ci_tk.time, "sleep", lambda delay: None)


def _error(path, reason="no such file or directory"):
    return _tkinter.TclError(
        "Can't find a usable tk.tcl: couldn't read file \"" + str(path) + '": ' + reason
    )


@pytest.mark.parametrize("reason", ["no such file or directory", "permission denied", "No error"])
def test_verified_readable_script_retries_native_call_and_retains_hash(tmp_path, reason):
    script = tmp_path / "Tk libraries/button.tcl"
    script.parent.mkdir()
    script.write_bytes(b"# a synthetic Tcl script")
    report = tmp_path / "startup.jsonl"
    calls = []
    sentinel = object()

    def create(*args, **kwargs):
        calls.append((args, kwargs))
        if len(calls) == 1:
            raise _error(script, reason)
        return sentinel

    assert ci_tk._create_with_retry(create, ("screen",), {"wantTk": True}, report) is sentinel
    assert calls == [(("screen",), {"wantTk": True})] * 2
    evidence = json.loads(report.read_text())
    assert evidence["recovered"] is True
    assert evidence["scripts"][0]["sha256"] == hashlib.sha256(script.read_bytes()).hexdigest()
    assert evidence["scripts"][0]["size"] == len(script.read_bytes())


@pytest.mark.parametrize("kind", ["missing", "empty", "directory"])
def test_genuinely_unreadable_script_is_fatal_without_retry(tmp_path, kind):
    script = tmp_path / "button.tcl"
    if kind == "empty":
        script.touch()
    elif kind == "directory":
        script.mkdir()
    calls = []
    error = _error(script)

    def create():
        calls.append(True)
        raise error

    report = tmp_path / "startup.jsonl"
    with pytest.raises(_tkinter.TclError) as caught:
        ci_tk._create_with_retry(create, (), {}, report)
    assert caught.value is error
    assert len(calls) == 1
    evidence = json.loads(report.read_text())
    assert not evidence["recovered"]
    assert not evidence["retryable"]


@pytest.mark.parametrize(
    "error",
    [
        AssertionError("Wrong balance"),
        ValueError("Invalid report"),
        _tkinter.TclError("unknown widget"),
        _tkinter.TclError("Can't find a usable tk.tcl: version conflict"),
        _tkinter.TclError("Can't find a usable tk.tcl: no such file or directory"),
    ],
)
def test_other_errors_are_not_retried(tmp_path, error):
    calls = []

    def create():
        calls.append(True)
        raise error

    with pytest.raises(type(error)) as caught:
        ci_tk._create_with_retry(create, (), {}, tmp_path / "startup.jsonl")
    assert caught.value is error
    assert len(calls) == 1


def test_repeated_library_failure_stays_fatal_after_one_retry(tmp_path):
    script = tmp_path / "button.tcl"
    script.write_text("# synthetic")
    calls = []

    def create():
        calls.append(True)
        raise _error(script)

    report = tmp_path / "startup.jsonl"
    with pytest.raises(_tkinter.TclError):
        ci_tk._create_with_retry(create, (), {}, report)
    assert len(calls) == 2
    evidence = json.loads(report.read_text())
    assert not evidence["recovered"]
    assert "retry_error" in evidence


def test_files_are_verified_after_the_bounded_settling_delay(tmp_path, monkeypatch):
    script = tmp_path / "button.tcl"
    calls = []
    delays = []

    def settle(delay):
        delays.append(delay)
        script.write_text("# readable after a temporary lock settles")

    def create():
        calls.append(True)
        if len(calls) == 1:
            raise _error(script)
        return "native interpreter"

    monkeypatch.setattr(ci_tk.time, "sleep", settle)
    assert (
        ci_tk._create_with_retry(create, (), {}, tmp_path / "startup.jsonl") == "native interpreter"
    )
    assert delays == [0.25]
    assert len(calls) == 2


@pytest.mark.parametrize(("platform", "actions"), [("linux", "true"), ("win32", "false")])
def test_guard_leaves_local_runs_and_other_platforms_untouched(monkeypatch, platform, actions):
    original = _tkinter.create
    monkeypatch.setattr(ci_tk.sys, "platform", platform)
    monkeypatch.setenv("GITHUB_ACTIONS", actions)
    with ci_tk.windows_tk_startup_retry():
        assert _tkinter.create is original
    assert _tkinter.create is original


def test_native_recovery_inside_test_body_still_constructs_real_widgets(
    tmp_path, monkeypatch, capsys
):
    original = _tkinter.create
    script = tmp_path / "button.tcl"
    script.write_text("# synthetic read-race evidence")
    calls = []

    def create(*args, **kwargs):
        calls.append(True)
        if len(calls) == 1:
            raise _error(script)
        return original(*args, **kwargs)

    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(_tkinter, "create", create)
    monkeypatch.setattr(ci_tk.sys, "platform", "win32")
    monkeypatch.setenv("GITHUB_ACTIONS", "true")
    with ci_tk.windows_tk_startup_retry():
        root = tk.Tk()
        try:
            root.withdraw()
            widget = tk.Button(root, text="Actual Tk button")
            assert widget.cget("text") == "Actual Tk button"
        finally:
            root.destroy()
    assert _tkinter.create is create
    assert len(calls) == 2
    assert "::warning title=Windows Tcl/Tk startup recovered::" in capsys.readouterr().out


def test_native_guard_restores_constructor_after_failure(tmp_path, monkeypatch):
    def create(*args, **kwargs):
        raise AssertionError("Real application failure")

    monkeypatch.setattr(_tkinter, "create", create)
    monkeypatch.setattr(ci_tk.sys, "platform", "win32")
    monkeypatch.setenv("GITHUB_ACTIONS", "true")
    with pytest.raises(AssertionError, match="Real application failure"):
        with ci_tk.windows_tk_startup_retry():
            _tkinter.create()
    assert _tkinter.create is create
