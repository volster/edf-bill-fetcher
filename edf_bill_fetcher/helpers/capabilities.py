"""Probe extraction prerequisites before starting a potentially long evidence run."""

from __future__ import annotations

import argparse
import importlib
import json
import subprocess
import sys
from dataclasses import asdict, dataclass
from pathlib import Path


@dataclass(frozen=True)
class Capability:
    """One independently checked capability and an actionable failure explanation."""

    key: str
    title: str
    available: bool
    detail: str


def _module_check(key: str, title: str, modules: tuple[tuple[str, str], ...]) -> Capability:
    try:
        for name, attribute in modules:
            if not callable(getattr(importlib.import_module(name), attribute, None)):
                raise RuntimeError(f"{name}.{attribute} is unavailable")
    except Exception as error:
        return Capability(key, title, False, f"Dependency could not load: {error}")
    return Capability(key, title, True, "Available")


def _ocr_check() -> Capability:
    try:
        from edf_bill_fetcher.io.adapters.ocr import find_tesseract

        runtime = find_tesseract()
        if runtime.tessdata is None:
            raise RuntimeError("English OCR language data (eng.traineddata) is missing")
        version = subprocess.run(
            [str(runtime.executable), "--version"],
            capture_output=True,
            text=True,
            check=True,
            timeout=5,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        ).stdout.splitlines()
        languages = subprocess.run(
            [str(runtime.executable), "--list-langs", "--tessdata-dir", str(runtime.tessdata)],
            capture_output=True,
            text=True,
            check=True,
            timeout=5,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        ).stdout.splitlines()
        if "eng" not in {line.strip() for line in languages}:
            raise RuntimeError("The OCR engine cannot find English language data")
        return Capability(
            "ocr", "Scanned PDFs and image pages", True, version[0] if version else "Available"
        )
    except Exception as error:
        return Capability("ocr", "Scanned PDFs and image pages", False, str(error))


def check_capabilities() -> tuple[Capability, ...]:
    """Load native libraries as well as modules; installed metadata alone is insufficient."""
    return (
        _module_check(
            "pdf",
            "Text PDFs and page previews",
            (("pdfplumber", "open"), ("pypdfium2", "PdfDocument")),
        ),
        _module_check("excel", "Excel workbooks", (("openpyxl", "Workbook"),)),
        _module_check("qt", "Desktop interface", (("PySide6.QtWidgets", "QApplication"),)),
        _module_check("pst", "Outlook PST / OST archives", (("pypff", "file"),)),
        _module_check("msg", "Outlook MSG email files", (("extract_msg", "Message"),)),
        _ocr_check(),
    )


def diagnostic_report() -> dict:
    """Return portable diagnostics without selected files, account details or secrets."""
    from edf_bill_fetcher.helpers.build_info import get_build_info

    checks = check_capabilities()
    return {
        "build": get_build_info(),
        "python": sys.version.split()[0],
        "ready": all(check.available for check in checks),
        "capabilities": [asdict(check) for check in checks],
    }


def main(argv: list[str] | None = None) -> int:
    """Check a source installation or save diagnostics from a windowed executable."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, help="Write JSON diagnostics to this file")
    options = parser.parse_args(argv)
    report = diagnostic_report()
    if options.output:
        options.output.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    elif sys.stdout is not None:
        for check in report["capabilities"]:
            sys.stdout.write(
                f"{'OK' if check['available'] else 'MISSING'}: {check['title']} — {check['detail']}\n"
            )
    return 0 if report["ready"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
