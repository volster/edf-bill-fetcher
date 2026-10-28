"""Source setup must fail honestly and select the requested development toolchain."""

import subprocess
from pathlib import Path

import pytest

import bootstrap
import build_all


@pytest.mark.parametrize(
    ("mode", "extras", "last"),
    [
        ("run", ".[msg]", "main.py"),
        ("develop", ".[dev,build,msg]", None),
        ("build", ".[build,msg]", "ci_build.py"),
    ],
)
def test_setup_installs_requested_tools_and_checks_before_launch(
    tmp_path, monkeypatch, mode, extras, last
):
    python = bootstrap.environment_python(tmp_path)
    python.parent.mkdir(parents=True, exist_ok=True)
    python.touch()
    commands: list[list[str]] = []
    monkeypatch.setattr(bootstrap, "run", commands.append)
    monkeypatch.setattr(
        bootstrap.subprocess, "run", lambda *args, **kwargs: subprocess.CompletedProcess(args, 0)
    )
    assert bootstrap.main([mode, "--venv", str(tmp_path)]) == 0
    installation = next(command for command in commands if "-e" in command)
    assert installation[-1] == extras
    assert str(bootstrap.ROOT / ".github/ci-constraints.txt") in installation
    checks = next(
        index
        for index, command in enumerate(commands)
        if "edf_bill_fetcher.helpers.capabilities" in command
    )
    if last:
        assert commands[-1] == [str(python), str(bootstrap.ROOT / last)]
        assert checks < len(commands) - 1
    else:
        assert checks == len(commands) - 1


def test_missing_ocr_fails_setup_without_launching_or_building(tmp_path, monkeypatch):
    python = bootstrap.environment_python(tmp_path)
    python.parent.mkdir(parents=True, exist_ok=True)
    python.touch()
    commands: list[list[str]] = []

    def run(command):
        commands.append(command)
        if "edf_bill_fetcher.helpers.capabilities" in command:
            raise subprocess.CalledProcessError(1, command)

    monkeypatch.setattr(bootstrap, "run", run)
    monkeypatch.setattr(
        bootstrap.subprocess, "run", lambda *args, **kwargs: subprocess.CompletedProcess(args, 0)
    )
    assert bootstrap.main(["build", "--venv", str(tmp_path)]) == 1
    assert not any(str(bootstrap.ROOT / "ci_build.py") in command for command in commands)


def test_check_only_does_not_install_or_modify_environment(tmp_path, monkeypatch):
    python = bootstrap.environment_python(tmp_path)
    python.parent.mkdir(parents=True, exist_ok=True)
    python.touch()
    commands: list[list[str]] = []
    monkeypatch.setattr(bootstrap, "run", commands.append)
    assert bootstrap.main(["--check", "--venv", str(tmp_path)]) == 0
    assert len(commands) == 2
    assert commands[0][:2] == [str(python), "-c"]
    assert commands[-1] == [str(python), "-m", "edf_bill_fetcher.helpers.capabilities"]
    assert bootstrap.main(["--check", "--venv", str(tmp_path / "missing")]) == 1
    assert not (tmp_path / "missing").exists()


def test_setup_does_not_overwrite_unrelated_directory(tmp_path):
    existing = tmp_path / "notes.txt"
    existing.write_text("keep me", encoding="utf-8")
    assert bootstrap.main(["develop", "--venv", str(tmp_path)]) == 1
    assert existing.read_text() == "keep me"


def test_build_download_correlates_request_and_requires_every_platform(tmp_path, monkeypatch):
    calls = []

    def gh(arguments, *, capture=False):
        calls.append(arguments)
        if arguments[:2] == ["run", "list"]:
            import json

            return json.dumps(
                [
                    {
                        "databaseId": 999,
                        "displayTitle": "EDF build someone-else",
                        "headSha": "wrong",
                        "url": "unrelated",
                    },
                    {
                        "databaseId": 123,
                        "displayTitle": "EDF build requested",
                        "headSha": "wanted",
                        "url": "correct",
                    },
                ]
            )
        if arguments[:2] == ["run", "download"]:
            target = Path(arguments[arguments.index("--dir") + 1])
            for platform, filename in {
                "windows-2025": "EDF_Evidence_Collector.exe",
                "macos-15": "EDF_Evidence_Collector-macos-arm64-ad-hoc.zip",
                "ubuntu-24.04": "EDF_Evidence_Collector",
            }.items():
                folder = target / f"EDF_Evidence_Collector-{platform}-manual-unvalidated"
                folder.mkdir()
                (folder / filename).write_bytes(b"downloaded")
        return ""

    monkeypatch.setattr(build_all, "gh", gh)
    monkeypatch.setattr(build_all, "uuid4", lambda: type("Request", (), {"hex": "requested"})())
    result = build_all.build_all("owner/repo", "branch", tmp_path)
    assert result == tmp_path / "123"
    assert calls[0][-2:] == ["-f", "request_id=requested"]
    assert ["run", "watch", "123", "--repo", "owner/repo", "--exit-status"] in calls
    assert not any("999" in command for command in calls)


def test_failed_build_never_downloads_older_artifacts(tmp_path, monkeypatch):
    calls = []

    def gh(arguments, *, capture=False):
        calls.append(arguments)
        if arguments[:2] == ["run", "list"]:
            return '[{"databaseId":123,"displayTitle":"EDF build requested","headSha":"wanted","url":"correct"}]'
        if arguments[:2] == ["run", "watch"]:
            raise subprocess.CalledProcessError(1, arguments)
        return ""

    monkeypatch.setattr(build_all, "gh", gh)
    monkeypatch.setattr(build_all, "uuid4", lambda: type("Request", (), {"hex": "requested"})())
    with pytest.raises(subprocess.CalledProcessError):
        build_all.build_all("owner/repo", "branch", tmp_path)
    assert not any(command[:2] == ["run", "download"] for command in calls)
    assert list(tmp_path.iterdir()) == []


@pytest.mark.parametrize("platform", ["win32", "darwin", "linux"])
def test_native_setup_uses_platform_package_manager_without_disabling_verification(
    monkeypatch, platform, tmp_path
):
    monkeypatch.setenv("ProgramFiles", str(tmp_path))
    monkeypatch.setattr(
        bootstrap.shutil,
        "which",
        lambda name: name if name in {"winget", "brew", "apt-get"} else None,
    )
    commands = bootstrap.system_commands(platform)
    assert any("tesseract" in argument.lower() for command in commands for argument in command)
    assert not any(
        "--ignore-security-hash" in command or "--allow-unauthenticated" in command
        for command in commands
    )
    if platform == "linux":
        assert commands[0][:2] == ["sudo", "apt-get"]
        assert bootstrap.system_commands(platform, administrator=True)[0] == ["apt-get", "update"]


def test_build_main_reports_missing_cli_and_dispatch_errors(tmp_path, monkeypatch):
    monkeypatch.setattr(build_all.shutil, "which", lambda name: None)
    assert build_all.main([]) == 1
    monkeypatch.setattr(build_all.shutil, "which", lambda name: name)

    def fail(*args):
        raise subprocess.CalledProcessError(1, "gh")

    monkeypatch.setattr(build_all, "build_all", fail)
    assert build_all.main(["--output", str(tmp_path)]) == 1


def test_unsupported_python_fails_before_installation(monkeypatch, tmp_path):
    monkeypatch.setattr(bootstrap.sys, "version_info", (3, 13))
    monkeypatch.setattr(
        bootstrap, "run", lambda command: pytest.fail("Must reject Python before running setup")
    )
    assert bootstrap.main(["develop", "--venv", str(tmp_path / "new")]) == 1


def test_build_correlation_timeout_never_falls_back_to_another_run(tmp_path, monkeypatch):
    calls = []
    clock = iter([0, 0, 121])
    monkeypatch.setattr(build_all.time, "monotonic", lambda: next(clock))
    monkeypatch.setattr(build_all.time, "sleep", lambda duration: None)

    def gh(arguments, *, capture=False):
        calls.append(arguments)
        return '[{"databaseId":999,"displayTitle":"unrelated"}]'

    monkeypatch.setattr(build_all, "gh", gh)
    with pytest.raises(RuntimeError, match="no older artifacts"):
        build_all.build_all("owner/repo", "branch", tmp_path)
    assert not any(command[:2] in [["run", "download"], ["run", "watch"]] for command in calls)


def test_installed_windows_ocr_is_not_reinstalled_on_repeated_setup(tmp_path, monkeypatch):
    engine = tmp_path / "Tesseract-OCR/tesseract.exe"
    engine.parent.mkdir()
    engine.touch()
    (engine.parent / "tessdata").mkdir()
    (engine.parent / "tessdata/eng.traineddata").write_bytes(b"installed")
    monkeypatch.setenv("ProgramFiles", str(tmp_path))
    monkeypatch.setattr(bootstrap.shutil, "which", lambda name: None)
    assert bootstrap.system_commands("win32") == []
