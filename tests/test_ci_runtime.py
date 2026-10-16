"""Runtime failures must leave useful diagnostics without leaking environment values."""

import json
from pathlib import Path

import pytest

import ci_runtime


@pytest.fixture
def isolated_runtime(monkeypatch):
    monkeypatch.delenv("pythonLocation", raising=False)
    monkeypatch.delenv("GITHUB_ACTIONS", raising=False)
    monkeypatch.setattr(ci_runtime, "check_tk", lambda: {"tcl": "checked"})
    monkeypatch.setattr(ci_runtime, "check_qt", lambda: {"workspace": "opened"})


def test_inventory_contains_selected_python_and_only_allowlisted_environment(
    tmp_path, monkeypatch, isolated_runtime
):
    monkeypatch.setenv("GH_TOKEN", "must-never-be-recorded")
    monkeypatch.setenv("TK_LIBRARY", "/runtime/tk")
    monkeypatch.setattr(ci_runtime, "check_tk", lambda: pytest.fail("Inventory loaded Tk"))
    path = tmp_path / "diagnostics/runtime.json"
    assert ci_runtime.main(["--inventory", "--report", str(path)]) == 0
    text = path.read_text()
    report = json.loads(text)
    assert report["executable"] == ci_runtime.sys.executable
    assert report["environment"]["TK_LIBRARY"] == "/runtime/tk"
    assert "must-never-be-recorded" not in text
    assert "GH_TOKEN" not in text
    assert "tk" not in report


def test_complete_preflight_records_both_gui_checks(tmp_path, isolated_runtime):
    path = tmp_path / "runtime.json"
    assert ci_runtime.main(["--report", str(path)]) == 0
    report = json.loads(path.read_text())
    assert report["status"] == "success"
    assert report["tk"] == {"tcl": "checked"}
    assert report["qt"] == {"workspace": "opened"}


@pytest.mark.parametrize("check", ["check_tk", "check_qt"])
def test_failed_native_library_probe_retains_runtime_and_traceback(
    tmp_path, monkeypatch, isolated_runtime, check
):
    def unavailable():
        raise OSError("Native library cannot be read")

    monkeypatch.setattr(ci_runtime, check, unavailable)
    path = tmp_path / "runtime.json"
    assert ci_runtime.main(["--report", str(path)]) == 1
    report = json.loads(path.read_text())
    assert report["status"] == "failure"
    assert report["executable"] == ci_runtime.sys.executable
    assert "Native library cannot be read" in report["error"]
    if check == "check_qt":
        assert report["tk"] == {"tcl": "checked"}


def test_path_interpreter_mismatch_fails_before_gui_loading(
    tmp_path, monkeypatch, isolated_runtime
):
    monkeypatch.setenv("pythonLocation", str(tmp_path / "another-python"))
    monkeypatch.setattr(
        ci_runtime, "check_tk", lambda: pytest.fail("Loaded wrong interpreter's Tk")
    )
    path = tmp_path / "runtime.json"
    assert ci_runtime.main(["--report", str(path)]) == 1
    report = json.loads(path.read_text())
    assert "pythonLocation" in report["error"]


def test_matching_interpreter_location_is_accepted(tmp_path, monkeypatch, isolated_runtime):
    monkeypatch.setenv("pythonLocation", str(Path(ci_runtime.sys.prefix).resolve()))
    assert ci_runtime.main(["--inventory", "--report", str(tmp_path / "runtime.json")]) == 0


@pytest.mark.parametrize("binary", ["bin/python", "python.exe"])
def test_setup_python_prefix_supports_posix_and_windows_binary_layouts(
    tmp_path, monkeypatch, isolated_runtime, binary
):
    prefix = tmp_path / "python"
    monkeypatch.setenv("pythonLocation", str(prefix))
    monkeypatch.setattr(ci_runtime.sys, "prefix", str(prefix))
    monkeypatch.setattr(ci_runtime.sys, "executable", str(prefix / binary))
    assert ci_runtime.main(["--inventory", "--report", str(tmp_path / "runtime.json")]) == 0


def test_actions_failure_annotation_escapes_traceback_lines(
    tmp_path, monkeypatch, isolated_runtime, capsys
):
    monkeypatch.setenv("GITHUB_ACTIONS", "true")
    monkeypatch.setenv("pythonLocation", str(tmp_path / "wrong-prefix"))
    assert ci_runtime.main(["--inventory", "--report", str(tmp_path / "runtime.json")]) == 1
    output = capsys.readouterr().out
    assert output.startswith("::error title=CI runtime preflight::")
    assert "%0A" in output
    assert len(output.splitlines()) == 1
