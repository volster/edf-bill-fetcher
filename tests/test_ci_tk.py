"""Recover native startup only; genuine missing files and application errors stay fatal."""

import _tkinter
import hashlib
import json
import os
import sys
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


@pytest.mark.parametrize("failure", ["script_read", "autoload_bootstrap", "direct_library_read"])
@pytest.mark.parametrize("configured_libraries", [False, True])
def test_native_recovery_inside_test_body_still_constructs_real_widgets(
    tmp_path, monkeypatch, capsys, failure, configured_libraries
):
    if configured_libraries:
        probe = tk.Tk()
        try:
            monkeypatch.setenv("TCL_LIBRARY", str(probe.tk.eval("info library")))
            monkeypatch.setenv("TK_LIBRARY", str(probe.tk.eval("set tk_library")))
        finally:
            probe.destroy()
    native_prefix = sys.base_prefix
    native_libraries = {name: os.environ.get(name) for name in ("TCL_LIBRARY", "TK_LIBRARY")}
    original = _tkinter.create
    script = tmp_path / "button.tcl"
    script.write_text("# synthetic read-race evidence")
    if failure in {"autoload_bootstrap", "direct_library_read"}:
        files = _bootstrap_files(tmp_path, monkeypatch)
        if failure == "direct_library_read":
            script = files["auto.tcl"]
    calls = []

    def create(*args, **kwargs):
        calls.append(True)
        if len(calls) == 1:
            if failure == "autoload_bootstrap":
                raise _tkinter.TclError('invalid command name "tcl_findLibrary"')
            if failure == "direct_library_read":
                raise _tkinter.TclError(
                    f'couldn\'t read file "{script}": no such file or directory'
                )
            raise _error(script)
        # Synthetic scripts prove retry eligibility; they cannot initialize a
        # real interpreter. In particular, Conda may require its library vars.
        with monkeypatch.context() as native:
            native.setattr(sys, "base_prefix", native_prefix)
            for name, value in native_libraries.items():
                if value is None:
                    native.delenv(name, raising=False)
                else:
                    native.setenv(name, value)
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


def _bootstrap_files(tmp_path, monkeypatch):
    monkeypatch.setattr(ci_tk.sys, "base_prefix", str(tmp_path))
    monkeypatch.delenv("TCL_LIBRARY", raising=False)
    monkeypatch.delenv("TK_LIBRARY", raising=False)
    directory = tmp_path / "tcl"
    tcl = directory / f"tcl{_tkinter.TCL_VERSION}"
    tk_library = directory / f"tk{_tkinter.TK_VERSION}"
    tcl.mkdir(parents=True)
    tk_library.mkdir()
    files = {
        "init.tcl": tcl / "init.tcl",
        "auto.tcl": tcl / "auto.tcl",
        "tclIndex": tcl / "tclIndex",
        "tk.tcl": tk_library / "tk.tcl",
    }
    files["init.tcl"].write_text("# synthetic Tcl initialization")
    files["auto.tcl"].write_text("proc tcl_findLibrary {args} {return}")
    files["tclIndex"].write_text("set auto_index(tcl_findLibrary) {source auto.tcl}")
    files["tk.tcl"].write_text("# synthetic Tk initialization")
    return files


def test_missing_bootstrap_command_requires_verified_definition_and_autoload_index(
    tmp_path, monkeypatch
):
    files = _bootstrap_files(tmp_path, monkeypatch)
    calls = []
    sentinel = object()

    def create(*args, **kwargs):
        calls.append((args, kwargs))
        if len(calls) == 1:
            raise _tkinter.TclError('invalid command name "tcl_findLibrary"')
        return sentinel

    report = tmp_path / "startup.jsonl"
    assert ci_tk._create_with_retry(create, ("screen",), {"wantTk": True}, report) is sentinel
    assert calls == [(("screen",), {"wantTk": True})] * 2
    evidence = json.loads(report.read_text())
    assert evidence["recovered"] is True
    assert len(evidence["scripts"]) == 4
    for script in evidence["scripts"]:
        assert (
            script["sha256"]
            == hashlib.sha256(files[ci_tk.Path(script["path"]).name].read_bytes()).hexdigest()
        )
    assert any(script.get("defines_bootstrap_command") for script in evidence["scripts"])
    assert any(script.get("indexes_bootstrap_command") for script in evidence["scripts"])


@pytest.mark.parametrize("problem", ["missing", "empty", "definition", "index", "override"])
def test_missing_or_inconsistent_bootstrap_scripts_are_fatal(tmp_path, monkeypatch, problem):
    files = _bootstrap_files(tmp_path, monkeypatch)
    if problem == "missing":
        files["auto.tcl"].unlink()
    elif problem == "empty":
        files["init.tcl"].write_text("")
    elif problem == "definition":
        files["auto.tcl"].write_text("# No bootstrap definition")
    elif problem == "index":
        files["tclIndex"].write_text("# No bootstrap index")
    else:
        monkeypatch.setenv("TCL_LIBRARY", str(tmp_path / "custom"))
    calls = []

    def create():
        calls.append(True)
        raise _tkinter.TclError('invalid command name "tcl_findLibrary"')

    report = tmp_path / "startup.jsonl"
    with pytest.raises(_tkinter.TclError, match="tcl_findLibrary"):
        ci_tk._create_with_retry(create, (), {}, report)
    assert len(calls) == 1
    assert json.loads(report.read_text())["retryable"] is False


def test_repeated_bootstrap_error_is_still_fatal_after_one_retry(tmp_path, monkeypatch):
    _bootstrap_files(tmp_path, monkeypatch)
    calls = []

    def create():
        calls.append(True)
        raise _tkinter.TclError('invalid command name "tcl_findLibrary"')

    report = tmp_path / "startup.jsonl"
    with pytest.raises(_tkinter.TclError, match="tcl_findLibrary"):
        ci_tk._create_with_retry(create, (), {}, report)
    assert len(calls) == 2
    evidence = json.loads(report.read_text())
    assert evidence["recovered"] is False
    assert "retry_error" in evidence


@pytest.mark.parametrize("reason", ["no such file or directory", "permission denied", "No error"])
def test_direct_installed_library_read_failure_verifies_file_before_retry(
    tmp_path, monkeypatch, reason
):
    files = _bootstrap_files(tmp_path, monkeypatch)
    script = files["auto.tcl"]
    calls = []

    def create():
        calls.append(True)
        if len(calls) == 1:
            raise _tkinter.TclError(f'couldn\'t read file "{script}": {reason}')
        return "native interpreter"

    report = tmp_path / "startup.jsonl"
    assert ci_tk._create_with_retry(create, (), {}, report) == "native interpreter"
    assert len(calls) == 2
    evidence = json.loads(report.read_text())
    assert evidence["recovered"] is True
    assert evidence["scripts"][0]["sha256"] == hashlib.sha256(script.read_bytes()).hexdigest()


@pytest.mark.parametrize("problem", ["missing", "empty", "outside_installation", "repeated"])
def test_direct_library_read_recovery_stays_scoped_and_bounded(tmp_path, monkeypatch, problem):
    files = _bootstrap_files(tmp_path, monkeypatch)
    script = files["auto.tcl"]
    if problem == "missing":
        script.unlink()
    elif problem == "empty":
        script.write_text("")
    elif problem == "outside_installation":
        script = tmp_path / "application.tcl"
        script.write_text("# outside the selected runtime")
    calls = []

    def create():
        calls.append(True)
        raise _tkinter.TclError(f'couldn\'t read file "{script}": no such file or directory')

    report = tmp_path / "startup.jsonl"
    with pytest.raises(_tkinter.TclError, match="couldn't read file"):
        ci_tk._create_with_retry(create, (), {}, report)
    assert len(calls) == (2 if problem == "repeated" else 1)
    if report.exists():
        assert json.loads(report.read_text())["recovered"] is False
