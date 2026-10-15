"""Exercise bundled OCR through the delivered application's real extraction command."""

import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

candidates = [
    Path("dist/EDF_Evidence_Collector.app/Contents/MacOS/EDF_Evidence_Collector"),
    Path("dist/EDF_Evidence_Collector.exe" if os.name == "nt" else "dist/EDF_Evidence_Collector"),
]
executable = next(path.resolve() for path in candidates if path.is_file())
with tempfile.TemporaryDirectory(prefix="edf-frozen-ocr-") as directory:
    root = Path(directory)
    inputs = root / "input"
    inputs.mkdir()
    shutil.copyfile("tests/fixtures/scanned_bill.pdf", inputs / "scanned.pdf")
    records_path = root / "records.json"
    # Hiding PATH prevents accidentally exercising a machine-installed engine.
    environment = os.environ.copy()
    environment["PATH"] = os.environ.get("SystemRoot", "/usr") + (
        "/System32" if os.name == "nt" else "/bin"
    )
    subprocess.run(
        [
            str(executable),
            "--extract",
            "--pdf-dir",
            str(inputs),
            "--output",
            str(root / "report.xlsx"),
            "--records-json",
            str(records_path),
            "--no-filter-below",
            "--strict",
        ],
        env=environment,
        check=True,
        timeout=180,
    )
    result = json.loads(records_path.read_text(encoding="utf-8"))
    rows = result["records"]
    assert len(rows) == 1, rows
    row = rows[0]
    assert row["Invoice #"] == "KI-12345", row
    assert row["Balance (£)"] == 250, row
    assert row["Period Charge (£)"] == 125, row
    assert row["Text Extraction"] == "OCR", row
    assert row["OCR Pages"] == [1], row
    assert row["OCR Review Required"] is True, row
    assert result["status"] == "review_required", result["status"]
