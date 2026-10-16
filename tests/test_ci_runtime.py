"""Runtime failures must leave useful diagnostics without leaking environment values."""

import json
from pathlib import Path

import pytest

import ci_runtime


@pytest.fixture
def isolated_runtime(monkeypatch):
    monkeypatch.delenv("pythonLocation", raising=False)
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
    monkeypatch.setenv("pythonLocation", str(Path(ci_runtime.sys.executable).resolve().parent))
    assert ci_runtime.main(["--inventory", "--report", str(tmp_path / "runtime.json")]) == 0
