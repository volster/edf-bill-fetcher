"""Detect unusable native prerequisites, not merely their installed metadata."""

import json
import subprocess
from types import SimpleNamespace

import pytest

from edf_bill_fetcher.helpers import build_info, capabilities
from edf_bill_fetcher.io.adapters import ocr


def test_native_import_failure_is_visible_instead_of_claiming_support(monkeypatch):
    def load(name):
        raise OSError("Native DLL could not load")

    monkeypatch.setattr(capabilities.importlib, "import_module", load)
    result = capabilities._module_check("pst", "Archives", (("pypff", "file"),))
    assert result.available is False
    assert "DLL" in result.detail


@pytest.mark.parametrize(
    "failure", ["missing-data", "missing-english", "broken-executable", "timeout"]
)
def test_ocr_probe_detects_each_unusable_engine(tmp_path, monkeypatch, failure):
    monkeypatch.setattr(
        ocr,
        "find_tesseract",
        lambda: ocr.TesseractRuntime(
            tmp_path / "tesseract", None if failure == "missing-data" else tmp_path
        ),
    )

    def run(command, **kwargs):
        if failure == "broken-executable":
            raise OSError("Missing native library")
        if failure == "timeout":
            raise subprocess.TimeoutExpired(command, 5)
        return SimpleNamespace(
            stdout="tesseract 5\n" if "--version" in command else "Languages:\nfra\n"
        )

    monkeypatch.setattr(capabilities.subprocess, "run", run)
    result = capabilities._ocr_check()
    assert result.available is False
    assert result.detail


def test_doctor_writes_diagnostics_and_returns_failure_for_partial_install(tmp_path, monkeypatch):
    monkeypatch.setattr(
        capabilities,
        "check_capabilities",
        lambda: (capabilities.Capability("ocr", "Scans", False, "No engine"),),
    )
    monkeypatch.setattr(build_info, "get_build_info", lambda: {"commit": "abc", "mode": "source"})
    target = tmp_path / "diagnostics.json"
    assert capabilities.main(["--output", str(target)]) == 1
    report = json.loads(target.read_text())
    assert report["ready"] is False
    assert report["capabilities"][0]["detail"] == "No engine"
    assert report["build"]["commit"] == "abc"


def test_frozen_provenance_never_uses_an_unrelated_git_checkout(tmp_path, monkeypatch):
    package = tmp_path / "data"
    package.mkdir()
    (package / "build-info.json").write_text(
        json.dumps({"version": "1.2.3", "commit": "a" * 40, "dirty": False}), encoding="utf-8"
    )
    monkeypatch.setattr(build_info.sys, "frozen", True, raising=False)
    monkeypatch.setattr(build_info, "files", lambda name: package)
    monkeypatch.setattr(
        build_info, "source_build_info", lambda root: pytest.fail("Frozen apps must not query Git")
    )
    assert build_info.get_build_info() == {
        "version": "1.2.3",
        "commit": "a" * 40,
        "dirty": False,
        "mode": "packaged",
    }


def test_source_provenance_captures_revision_and_local_changes(tmp_path, monkeypatch):
    def run(command, **kwargs):
        return SimpleNamespace(stdout="b" * 40 if "rev-parse" in command else " M app.py\n")

    monkeypatch.setattr(build_info.subprocess, "run", run)
    info = build_info.source_build_info(tmp_path)
    assert info["commit"] == "b" * 40
    assert info["dirty"] is True


def test_frozen_missing_metadata_is_unknown_rather_than_guessed(tmp_path, monkeypatch):
    monkeypatch.setattr(build_info.sys, "frozen", True, raising=False)
    monkeypatch.setattr(build_info, "files", lambda name: tmp_path)
    assert build_info.get_build_info()["commit"] == "unknown"


def test_ocr_success_checks_executable_and_english_language(tmp_path, monkeypatch):
    monkeypatch.setattr(
        ocr, "find_tesseract", lambda: ocr.TesseractRuntime(tmp_path / "tesseract", tmp_path)
    )
    commands = []

    def run(command, **kwargs):
        commands.append(command)
        return SimpleNamespace(
            stdout="tesseract 5\n" if "--version" in command else "Languages:\neng\n"
        )

    monkeypatch.setattr(capabilities.subprocess, "run", run)
    assert capabilities._ocr_check().available is True
    assert "--version" in commands[0]
    assert "--tessdata-dir" in commands[1]
