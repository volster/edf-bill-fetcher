"""Choose the project environment without installing into base or hiding setup failures."""

import json
import ntpath
import os
import subprocess
from pathlib import Path
from types import SimpleNamespace

import pytest

import bootstrap


@pytest.fixture
def conda_setup(tmp_path, monkeypatch):
    base = tmp_path / "Conda installation"
    project = base / "envs" / "edf-bill-fetcher"
    unrelated = base / "envs" / "another-project"
    state: dict = {"envs": [str(base), str(unrelated)], "root_prefix": str(base)}
    commands: list[list[str]] = []
    monkeypatch.setattr(bootstrap, "ROOT", tmp_path)
    monkeypatch.setattr(bootstrap, "find_conda", lambda: "conda")

    def execute(command, **kwargs):
        commands.append(command)
        if command[1:] == ["info", "--envs", "--json"]:
            return subprocess.CompletedProcess(command, 0, json.dumps({"envs": state["envs"]}))
        if command[1:] == ["info", "--json"]:
            return subprocess.CompletedProcess(command, 0, json.dumps(state))
        if command[1:2] == ["create"]:
            state["envs"].append(str(project))
            python = project / ("python.exe" if bootstrap.os.name == "nt" else "bin/python")
            python.parent.mkdir(parents=True, exist_ok=True)
            python.touch()
        return subprocess.CompletedProcess(command, 0)

    monkeypatch.setattr(bootstrap.subprocess, "run", execute)
    return SimpleNamespace(base=base, project=project, state=state, commands=commands)


@pytest.mark.parametrize("mode", ["run", "develop", "build"])
def test_new_conda_project_is_created_once_and_all_work_uses_its_runtime(conda_setup, mode):
    setup = conda_setup
    assert bootstrap.main([mode]) == 0
    assert [
        "conda",
        "create",
        "--yes",
        "--name",
        "edf-bill-fetcher",
        "python=3.11",
        "pip",
    ] in setup.commands
    prefix = ["conda", "run", "--no-capture-output", "--prefix", str(setup.project), "python"]
    work = [command for command in setup.commands if command[1:2] not in (["info"], ["create"])]
    assert work
    assert all(command[: len(prefix)] == prefix for command in work)
    assert not (bootstrap.ROOT / ".venv").exists()
    install = next(command for command in work if "-e" in command)
    assert (
        install[-1]
        == {"run": ".[msg]", "develop": ".[dev,build,msg]", "build": ".[build,msg]"}[mode]
    )
    setup.commands.clear()
    assert bootstrap.main([mode]) == 0
    assert not any("create" in command for command in setup.commands)


def test_existing_project_reused_even_when_launcher_python_is_newer(conda_setup, monkeypatch):
    assert bootstrap.main(["develop"]) == 0
    conda_setup.commands.clear()
    monkeypatch.setattr(bootstrap.sys, "version_info", (3, 13))
    assert bootstrap.main(["--check", "--install-system"]) == 0
    assert not any("install" in command or "create" in command for command in conda_setup.commands)
    assert any(
        "edf_bill_fetcher.helpers.capabilities" in command for command in conda_setup.commands
    )


def test_check_missing_conda_environment_is_read_only(conda_setup):
    assert bootstrap.main(["--check", "--install-system"]) == 1
    assert conda_setup.commands == [["conda", "info", "--json"]]
    assert not conda_setup.project.exists()


def test_custom_project_name_reuses_registered_environment(conda_setup):
    other = conda_setup.base / "envs" / "edf-review"
    python = other / ("python.exe" if bootstrap.os.name == "nt" else "bin/python")
    python.parent.mkdir(parents=True)
    python.touch()
    conda_setup.state["envs"].append(str(other))
    assert bootstrap.main(["--check", "--conda-env", "edf-review"]) == 0
    assert any(str(other) in command for command in conda_setup.commands)
    assert not conda_setup.project.exists()


@pytest.mark.parametrize("name", ["", "base", "root", "../other", "bad name", "-bad", "bad&name"])
def test_unsafe_or_base_environment_names_are_rejected_without_commands(conda_setup, name):
    assert bootstrap.main(["develop", "--conda-env=" + name]) == 1
    assert conda_setup.commands == []


def test_base_with_project_basename_is_not_reused(conda_setup):
    conda_setup.state["root_prefix"] = str(conda_setup.project)
    conda_setup.state["envs"] = [str(conda_setup.project)]
    assert bootstrap.main(["--check"]) == 1
    assert conda_setup.commands == [["conda", "info", "--json"]]


def test_duplicate_names_are_rejected_instead_of_modifying_arbitrary_environment(conda_setup):
    conda_setup.state["envs"] += [
        str(conda_setup.project),
        str(conda_setup.base / "other" / "edf-bill-fetcher"),
    ]
    assert bootstrap.main(["develop"]) == 1
    assert conda_setup.commands == [["conda", "info", "--json"]]


def test_environment_without_python_is_not_overwritten(conda_setup):
    conda_setup.state["envs"].append(str(conda_setup.project))
    conda_setup.project.mkdir(parents=True)
    sentinel = conda_setup.project / "notes.txt"
    sentinel.write_text("keep", encoding="utf-8")
    assert bootstrap.main(["develop"]) == 1
    assert sentinel.read_text(encoding="utf-8") == "keep"
    assert not any("create" in command for command in conda_setup.commands)


@pytest.mark.parametrize(
    "output", ["not json", "[]", "{}", '{"root_prefix":"/base","envs":[null]}']
)
def test_invalid_conda_inventory_does_not_trigger_creation_or_venv(
    conda_setup, monkeypatch, output
):
    monkeypatch.setattr(
        bootstrap.subprocess, "run", lambda *a, **k: subprocess.CompletedProcess(a, 0, output)
    )
    assert bootstrap.main(["develop"]) == 1
    assert not (bootstrap.ROOT / ".venv").exists()


@pytest.mark.parametrize("failed_step", ["info", "create", "install", "capabilities", "version"])
def test_conda_failure_stops_setup_without_switching_environment(
    conda_setup, monkeypatch, failed_step
):
    execute = bootstrap.subprocess.run

    def fail(command, **kwargs):
        if (
            (failed_step in {"info", "create"} and command[1] == failed_step)
            or (failed_step == "install" and "install" in command)
            or (
                failed_step == "capabilities" and "edf_bill_fetcher.helpers.capabilities" in command
            )
            or (failed_step == "version" and "-c" in command)
        ):
            raise subprocess.CalledProcessError(1, command, stderr="setup failure")
        return execute(command, **kwargs)

    monkeypatch.setattr(bootstrap.subprocess, "run", fail)
    assert bootstrap.main(["run"]) == 1
    assert not (bootstrap.ROOT / ".venv").exists()
    assert not any(str(bootstrap.ROOT / "main.py") in command for command in conda_setup.commands)


def test_windows_conda_uses_root_python_and_activated_run(conda_setup, monkeypatch):
    monkeypatch.setattr(
        bootstrap, "os", SimpleNamespace(name="nt", path=ntpath, environ=os.environ)
    )
    assert bootstrap.main(["develop"]) == 0
    assert (conda_setup.project / "python.exe").is_file()
    assert not (conda_setup.project / "Scripts/python.exe").exists()


def test_no_conda_creates_then_reuses_default_venv(tmp_path, monkeypatch):
    monkeypatch.setattr(bootstrap, "ROOT", tmp_path)
    monkeypatch.setattr(bootstrap, "find_conda", lambda: None)
    commands: list[list[str]] = []
    created: list[Path] = []

    def create(self, path):
        created.append(path)
        python = bootstrap.environment_python(path)
        python.parent.mkdir(parents=True)
        python.touch()

    monkeypatch.setattr(bootstrap.venv.EnvBuilder, "create", create)
    monkeypatch.setattr(bootstrap, "run", commands.append)
    monkeypatch.setattr(
        bootstrap.subprocess, "run", lambda *a, **k: subprocess.CompletedProcess(a, 0)
    )
    assert bootstrap.main(["develop"]) == 0
    assert bootstrap.main(["develop"]) == 0
    assert created == [tmp_path / ".venv"]
    assert all(
        command[0] == str(bootstrap.environment_python(tmp_path / ".venv")) for command in commands
    )


def test_explicit_venv_bypasses_conda_detection(tmp_path, monkeypatch):
    python = bootstrap.environment_python(tmp_path)
    python.parent.mkdir(parents=True, exist_ok=True)
    python.touch()
    monkeypatch.setattr(
        bootstrap, "find_conda", lambda: pytest.fail("Explicit venv must bypass Conda")
    )
    monkeypatch.setattr(bootstrap, "run", lambda command: None)
    assert bootstrap.main(["--check", "--venv", str(tmp_path)]) == 0


def test_explicit_conda_does_not_fall_back_when_unavailable(tmp_path, monkeypatch):
    monkeypatch.setattr(bootstrap, "ROOT", tmp_path)
    monkeypatch.setattr(bootstrap, "find_conda", lambda: None)
    assert bootstrap.main(["develop", "--conda-env", "edf-review"]) == 1
    assert not (tmp_path / ".venv").exists()


@pytest.mark.parametrize("source", ["CONDA_EXE", "PATH", "home"])
def test_conda_discovery_without_requiring_shell_activation(tmp_path, monkeypatch, source):
    executable = tmp_path / "miniconda3" / ("Scripts/conda.exe" if os.name == "nt" else "bin/conda")
    executable.parent.mkdir(parents=True)
    executable.touch()
    monkeypatch.delenv("CONDA_EXE", raising=False)
    monkeypatch.setattr(
        bootstrap.shutil, "which", lambda name: str(executable) if source == "PATH" else None
    )
    monkeypatch.setattr(Path, "home", lambda: tmp_path)
    monkeypatch.setattr(bootstrap.sys, "prefix", str(tmp_path / "absent"))
    monkeypatch.setattr(bootstrap.sys, "base_prefix", str(tmp_path / "absent"))
    if source == "CONDA_EXE":
        monkeypatch.setenv("CONDA_EXE", str(executable))
    assert bootstrap.find_conda() == str(executable.resolve())
