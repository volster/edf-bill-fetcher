"""Prepare a complete source installation, then run, develop or build the app."""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import venv
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def system_commands(platform: str, *, administrator: bool = False) -> list[list[str]]:
    """Use native package managers; never download and execute an installer script."""
    if platform == "win32":
        engine = shutil.which("tesseract")
        candidates = [Path(engine)] if engine else []
        candidates.append(
            Path(os.environ.get("ProgramFiles", r"C:\Program Files"))
            / "Tesseract-OCR/tesseract.exe"
        )
        if any(
            path.is_file() and (path.parent / "tessdata/eng.traineddata").is_file()
            for path in candidates
        ):
            return []
        if shutil.which("winget"):
            return [
                [
                    "winget",
                    "install",
                    "--exact",
                    "--id",
                    "UB-Mannheim.TesseractOCR",
                    "--accept-source-agreements",
                    "--accept-package-agreements",
                ]
            ]
        if shutil.which("choco"):
            return [["choco", "install", "tesseract", "--yes", "--no-progress"]]
        raise RuntimeError(
            "Install Tesseract OCR with English data in C:\\Program Files\\Tesseract-OCR, then rerun."
        )
    if platform == "darwin":
        if not shutil.which("brew"):
            raise RuntimeError(
                "Install Homebrew, or install Tesseract with English data and the native build tools yourself."
            )
        return [["brew", "install", "tesseract", "autoconf", "automake", "libtool", "pkg-config"]]
    if platform == "linux" and shutil.which("apt-get"):
        prefix = [] if administrator else ["sudo"]
        return [
            [*prefix, "apt-get", "update"],
            [
                *prefix,
                "apt-get",
                "install",
                "-y",
                "tesseract-ocr",
                "tesseract-ocr-eng",
                "python3-dev",
                "python3-venv",
                "build-essential",
                "autoconf",
                "automake",
                "libtool",
                "pkg-config",
                "libxcb-cursor0",
                "libxkbcommon-x11-0",
                "libegl1",
            ],
        ]
    raise RuntimeError(
        "Install Tesseract with English language data and your distribution's Python/Qt native dependencies, then rerun without --install-system."
    )


def run(command: list[str]) -> None:
    """Keep setup commands visible and preserve every subprocess failure."""
    sys.stdout.write("Running: " + subprocess.list2cmdline(command) + "\n")
    sys.stdout.flush()
    subprocess.run(command, cwd=ROOT, check=True)


def environment_python(directory: Path) -> Path:
    """Use the environment directly, avoiding shell activation and execution policies."""
    return directory / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def find_conda() -> str | None:
    """Find Conda from its shell configuration, PATH or usual installation roots."""
    candidates = [os.environ.get("CONDA_EXE"), shutil.which("conda")]
    roots = [Path(sys.prefix), Path(sys.base_prefix)]
    for parent in (Path.home(), Path.home() / "opt", Path("/opt")):
        roots.extend(parent / name for name in ("miniconda3", "anaconda3", "miniforge3"))
    for variable in ("LOCALAPPDATA", "ProgramData"):
        if os.environ.get(variable):
            roots.extend(
                Path(os.environ[variable]) / name
                for name in ("miniconda3", "anaconda3", "miniforge3")
            )
    for root in roots:
        candidates.append(str(root / ("Scripts/conda.exe" if os.name == "nt" else "bin/conda")))
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            path = Path(candidate)
            # Prefer the executable over condabin's Windows batch wrapper.
            executable = path.parent.parent / "Scripts/conda.exe"
            if path.suffix.lower() == ".bat" and executable.is_file():
                path = executable
            return str(path.resolve())
    return None


def conda_environment(conda: str, name: str) -> Path | None:
    """Resolve a named project environment without ever selecting Conda's base."""
    # --envs alone omits root_prefix in newer Conda versions. Full JSON contains both.
    response = subprocess.run(
        [conda, "info", "--json"],
        cwd=ROOT,
        capture_output=True,
        encoding="utf-8",
        check=True,
        timeout=30,
    )
    try:
        info = json.loads(response.stdout)
        if (
            not isinstance(info, dict)
            or not isinstance(info.get("root_prefix"), str)
            or not info["root_prefix"]
            or not Path(info["root_prefix"]).is_absolute()
            or not isinstance(info.get("envs"), list)
            or not all(isinstance(path, str) and Path(path).is_absolute() for path in info["envs"])
        ):
            raise ValueError("missing or invalid environment paths")
        base = Path(info["root_prefix"]).resolve()
        matches = {
            Path(path).resolve()
            for path in info["envs"]
            if os.path.normcase(Path(path).name) == os.path.normcase(name)
            and Path(path).resolve() != base
        }
    except (ValueError, TypeError) as error:
        raise RuntimeError(f"Cannot read Conda's environment list: {error}") from error
    if len(matches) > 1:
        raise RuntimeError(
            f"Several Conda environments are named {name!r}. "
            "Use --conda-env with a unique project environment name."
        )
    return next(iter(matches), None)


def conda_command(conda: str, name: str, *, check: bool) -> list[str]:
    """Reuse or create the project environment and run with Conda activation semantics."""
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]*", name) or name.lower() in {"base", "root"}:
        raise RuntimeError(
            "Choose a project --conda-env name using letters, digits, _, - or .; not base/root."
        )
    directory = conda_environment(conda, name)
    if directory is None:
        if check:
            raise RuntimeError(
                f"Conda environment {name!r} does not exist. Run setup without --check."
            )
        run([conda, "create", "--yes", "--name", name, "python=3.11", "pip"])
        directory = conda_environment(conda, name)
        if directory is None:
            raise RuntimeError("Conda did not register the new environment; setup cannot continue.")
    python = directory / ("python.exe" if os.name == "nt" else "bin/python")
    if not python.is_file():
        raise RuntimeError(
            f"Conda environment {directory} contains no Python. Repair it or choose another --conda-env name."
        )
    sys.stdout.write(f"Using Conda environment: {directory}\n")
    return [conda, "run", "--no-capture-output", "--prefix", str(directory), "python"]


def venv_command(directory: Path, *, check: bool) -> list[str]:
    """Reuse a project venv or create one without overwriting unrelated files."""
    directory = directory.expanduser().resolve()
    python = environment_python(directory)
    if not python.is_file():
        if check:
            raise RuntimeError("The virtual environment does not exist. Run setup without --check.")
        if sys.version_info[:2] not in {(3, 10), (3, 11), (3, 12)}:
            raise RuntimeError(
                "Creating a venv requires Python 3.10, 3.11 or 3.12; 3.11 is recommended."
            )
        if directory.exists() and (not directory.is_dir() or any(directory.iterdir())):
            raise RuntimeError(
                "The environment directory is not empty and contains no Python. Choose another --venv directory."
            )
        venv.EnvBuilder(with_pip=True).create(directory)
    sys.stdout.write(f"Using virtual environment: {directory}\n")
    return [str(python)]


def main(argv: list[str] | None = None) -> int:
    """Prefer a project Conda environment, falling back to venv when Conda is absent."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mode", choices=("run", "develop", "build"), nargs="?", default="run")
    parser.add_argument(
        "--install-system",
        action="store_true",
        help="Install native prerequisites using winget/Chocolatey, Homebrew or apt (may request administrator access)",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Only check the existing environment; do not install or launch anything",
    )
    environments = parser.add_mutually_exclusive_group()
    environments.add_argument(
        "--venv",
        type=Path,
        help="Explicitly use this venv directory instead of Conda (default fallback: .venv)",
    )
    environments.add_argument(
        "--conda-env",
        help="Use this Conda environment name (default: edf-bill-fetcher); requires Conda",
    )
    options = parser.parse_args(argv)
    try:
        conda = None if options.venv is not None else find_conda()
        if options.conda_env is not None and conda is None:
            raise RuntimeError("Conda was not found. Open a Conda prompt or use --venv .venv.")
        python = (
            conda_command(
                conda,
                options.conda_env if options.conda_env is not None else "edf-bill-fetcher",
                check=options.check,
            )
            if conda
            else None
        )
        # System packages can supply venv/ensurepip support before venv creation.
        if options.install_system and not options.check:
            for command in system_commands(
                sys.platform, administrator=hasattr(os, "geteuid") and os.geteuid() == 0
            ):
                run(command)
        if python is None:
            python = venv_command(options.venv or ROOT / ".venv", check=options.check)
        run(
            [
                *python,
                "-c",
                "import sys, struct; sys.exit("
                "'Use a Python 3.10-3.12 environment' if not (3, 10) <= sys.version_info[:2] <= (3, 12) "
                "else 'Use a 64-bit Python environment' if struct.calcsize('P') != 8 else 0)",
            ]
        )
        if not options.check:
            pip = subprocess.run([*python, "-m", "pip", "--version"], cwd=ROOT, capture_output=True)
            if pip.returncode:
                run([*python, "-m", "ensurepip", "--upgrade"])
            run([*python, "-m", "pip", "install", "--upgrade", "pip"])
            extras = (
                "dev,build,msg"
                if options.mode == "develop"
                else "build,msg"
                if options.mode == "build"
                else "msg"
            )
            run(
                [
                    *python,
                    "-m",
                    "pip",
                    "install",
                    "-c",
                    str(ROOT / ".github/ci-constraints.txt"),
                    "-e",
                    f".[{extras}]",
                ]
            )
        run([*python, "-m", "edf_bill_fetcher.helpers.capabilities"])
        if not options.check and options.mode in {"run", "build"}:
            run([*python, str(ROOT / ("main.py" if options.mode == "run" else "ci_build.py"))])
        elif not options.check:
            sys.stdout.write(
                "Development environment ready. Run tests with: "
                + subprocess.list2cmdline([*python, "-m", "pytest"])
                + "\n"
            )
        return 0
    except (OSError, RuntimeError, subprocess.SubprocessError) as error:
        sys.stderr.write(f"Setup did not complete: {error}\n")
        if isinstance(error, subprocess.CalledProcessError) and error.stderr:
            sys.stderr.write(str(error.stderr) + "\n")
        sys.stderr.write(
            "See docs/BUILDING.md for environment and native dependency setup. "
            "Conda failures are not silently retried in a different environment.\n"
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
