"""Runtime failures must leave useful diagnostics without leaking environment values."""

import json

import pytest

import ci_runtime


@pytest.fixture
def isolated_runtime(monkeypatch):
    monkeypatch.delenv("pythonLocation", raising=False)
    monkeypatch.delenv("GITHUB_ACTIONS", raising=False)
    monkeypatch.setattr(ci_runtime, "check_qt", lambda: {"workspace": "opened"})


def test_inventory_contains_selected_python_and_only_allowlisted_environment(
    tmp_path, monkeypatch, isolated_runtime
):
    monkeypatch.setenv("GH_TOKEN", "must-never-be-recorded")
    monkeypatch.setenv("QT_QPA_PLATFORM", "offscreen")
    monkeypatch.setattr(ci_runtime, "check_qt", lambda: pytest.fail("Inventory loaded Qt"))
    path = tmp_path / "diagnostics/runtime.json"
    assert ci_runtime.main(["--inventory", "--report", str(path)]) == 0
    text = path.read_text()
    report = json.loads(text)
    assert report["executable"] == ci_runtime.sys.executable
    assert report["environment"]["QT_QPA_PLATFORM"] == "offscreen"
    assert "must-never-be-recorded" not in text
    assert "GH_TOKEN" not in text
    assert "tk" not in report


def test_complete_preflight_records_the_qt_workspace(tmp_path, isolated_runtime):
    path = tmp_path / "runtime.json"
    assert ci_runtime.main(["--report", str(path)]) == 0
    report = json.loads(path.read_text())
    assert report["status"] == "success"
    assert report["qt"] == {"workspace": "opened"}


def test_failed_native_library_probe_retains_runtime_and_traceback(
    tmp_path, monkeypatch, isolated_runtime
):
    def unavailable():
        raise OSError("Native library cannot be read")

    monkeypatch.setattr(ci_runtime, "check_qt", unavailable)
    path = tmp_path / "runtime.json"
    assert ci_runtime.main(["--report", str(path)]) == 1
    report = json.loads(path.read_text())
    assert report["status"] == "failure"
    assert report["executable"] == ci_runtime.sys.executable
    assert "Native library cannot be read" in report["error"]


def test_inventory_accepts_macos_framework_and_toolcache_prefix_differences(
    tmp_path, monkeypatch, isolated_runtime
):
    monkeypatch.setattr(ci_runtime.sys, "prefix", str(tmp_path / "framework-prefix"))
    monkeypatch.setenv("pythonLocation", str(tmp_path / "toolcache-prefix"))
    assert ci_runtime.main(["--inventory", "--report", str(tmp_path / "runtime.json")]) == 0


def test_actions_failure_annotation_escapes_traceback_lines(
    tmp_path, monkeypatch, isolated_runtime, capsys
):
    def unavailable():
        raise OSError("Native library cannot be read")

    monkeypatch.setattr(ci_runtime, "check_qt", unavailable)
    monkeypatch.setenv("GITHUB_ACTIONS", "true")
    assert ci_runtime.main(["--report", str(tmp_path / "runtime.json")]) == 1
    output = capsys.readouterr().out
    assert output.startswith("::error title=CI runtime preflight::")
    assert "%0A" in output
    assert len(output.splitlines()) == 1
