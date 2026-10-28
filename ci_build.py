"""Build the same frozen application in CI and tagged releases."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from edf_bill_fetcher.helpers.build_info import build_metadata

HIDDEN_IMPORTS = (
    "PySide6.QtCore",
    "PySide6.QtGui",
    "PySide6.QtWidgets",
    "pypdfium2",
    "reportlab",
    "reportlab.platypus",
    "reportlab.lib",
    "reportlab.lib.colors",
    "reportlab.lib.styles",
    "reportlab.lib.units",
    "reportlab.lib.pagesizes",
    "reportlab.lib.enums",
    "reportlab.pdfbase",
    "reportlab.pdfbase.pdfmetrics",
    "reportlab.pdfbase.ttfonts",
    "scipy",
    "scipy.stats",
    "statsmodels",
    "statsmodels.tsa.holtwinters",
    "pypff",
)


def main() -> None:
    """Use a standard signed .app on macOS and one-file binaries elsewhere."""
    macos = sys.platform == "darwin"
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--onedir" if macos else "--onefile",
        "--windowed",
        "--name=EDF_Evidence_Collector",
        "--additional-hooks-dir=.github/pyinstaller_hooks",
        *[f"--hidden-import={module}" for module in HIDDEN_IMPORTS],
        "--exclude-module=tkinter",
        "--exclude-module=_tkinter",
        "--collect-all=extract_msg",
        "--collect-data=edf_bill_fetcher",
        "--copy-metadata=edf-bill-fetcher",
    ]
    environment = os.environ.copy()
    if macos:
        command.append("--osx-bundle-identifier=io.github.volster.edf-bill-fetcher")
        # PyInstaller signs collected Mach-O code and then the enclosing bundle.
        # Its default is to warn on a failed outer signature; distribution must fail.
        environment["PYINSTALLER_STRICT_BUNDLE_CODESIGN_ERROR"] = "1"
        environment["PYINSTALLER_VERIFY_BUNDLE_SIGNATURE"] = "1"
    with tempfile.TemporaryDirectory(prefix="edf-build-") as directory:
        metadata = Path(directory) / "build-info.json"
        metadata.write_text(
            json.dumps(build_metadata(Path(__file__).resolve().parent)), encoding="utf-8"
        )
        subprocess.run(
            [*command, f"--add-data={metadata}{os.pathsep}edf_bill_fetcher/data", "main.py"],
            env=environment,
            check=True,
        )


if __name__ == "__main__":
    main()
