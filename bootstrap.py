"""Prepare a complete source installation, then run, develop or build the app."""

from __future__ import annotations

import argparse
import os
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
                "python3-tk",
                "tk-dev",
                "build-essential",
                "autoconf",
                "automake",
                "libtool",
                "pkg-config",
                "libxcb-cursor0",
                "libxkbcommon-x11-0",
                "libegl1",
                "xvfb",
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


def main(argv: list[str] | None = None) -> int:
    """Provision a virtual environment and refuse to claim readiness with missing OCR."""
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
    parser.add_argument(
        "--venv", type=Path, default=ROOT / ".venv", help="Virtual environment directory"
    )
    options = parser.parse_args(argv)
    try:
        if sys.version_info[:2] not in {(3, 10), (3, 11), (3, 12)}:
            raise RuntimeError(
                "Use Python 3.10, 3.11 or 3.12; Python 3.11 is recommended and matches the executable builds."
            )
        python = environment_python(options.venv.resolve())
        if options.check:
            if not python.is_file():
                raise RuntimeError(
                    "The virtual environment does not exist. Run bootstrap.py without --check first."
                )
        else:
            if options.install_system:
                for command in system_commands(
                    sys.platform, administrator=hasattr(os, "geteuid") and os.geteuid() == 0
                ):
                    run(command)
            if not python.is_file():
                if options.venv.exists() and any(options.venv.iterdir()):
                    raise RuntimeError(
                        "The environment directory is not empty and contains no Python. Choose another --venv directory."
                    )
                venv.EnvBuilder(with_pip=True).create(options.venv)
            run(
                [
                    str(python),
                    "-c",
                    "import sys; assert (3, 10) <= sys.version_info[:2] <= (3, 12), 'Use a Python 3.10-3.12 environment'",
                ]
            )
            pip = subprocess.run(
                [str(python), "-m", "pip", "--version"], cwd=ROOT, capture_output=True
            )
            if pip.returncode:
                run([str(python), "-m", "ensurepip", "--upgrade"])
            run([str(python), "-m", "pip", "install", "--upgrade", "pip"])
            extras = (
                "dev,build,msg"
                if options.mode == "develop"
                else "build,msg"
                if options.mode == "build"
                else "msg"
            )
            run(
                [
                    str(python),
                    "-m",
                    "pip",
                    "install",
                    "-c",
                    str(ROOT / ".github/ci-constraints.txt"),
                    "-e",
                    f".[{extras}]",
                ]
            )
        run([str(python), "-m", "edf_bill_fetcher.helpers.capabilities"])
        if not options.check and options.mode in {"run", "build"}:
            run([str(python), str(ROOT / ("main.py" if options.mode == "run" else "ci_build.py"))])
        elif not options.check:
            sys.stdout.write(f"Development environment ready. Run tests with: {python} -m pytest\n")
        return 0
    except (OSError, RuntimeError, subprocess.SubprocessError) as error:
        sys.stderr.write(f"Setup did not complete: {error}\n")
        sys.stderr.write(
            "For missing native dependencies, rerun with --install-system or follow docs/BUILDING.md.\n"
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
