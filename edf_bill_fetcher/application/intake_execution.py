"""Publish a reusable mixed-input inventory before any billing analysis runs."""

from __future__ import annotations

import argparse
from collections.abc import Callable
from pathlib import Path
from uuid import uuid4

from openpyxl import Workbook

from edf_bill_fetcher.helpers.artifacts import write_json
from edf_bill_fetcher.helpers.excel_utils import literal_cell
from edf_bill_fetcher.io.adapters.intake import IntakeLimits, discover


def inspect_sources(
    paths: list[str],
    directory: str,
    *,
    cancelled: Callable[[], bool] = lambda: False,
    progress: Callable[[str], None] = lambda message: None,
) -> list[dict]:
    """Retain expanded originals and an inventory reusable by every interface."""
    target = Path(directory)
    target.mkdir(parents=True, exist_ok=True)
    rows: list[dict] = []
    expanded = 0
    for index, path in enumerate(paths):
        if len(rows) >= 100_000:
            raise ValueError("Mixed intake exceeds the aggregate file-count limit")
        for item in discover(
            Path(path),
            target / "intake_files" / str(index),
            cancelled=cancelled,
            limits=IntakeLimits(
                files=100_000 - len(rows), expanded_bytes=max(0, 64 * 1024**3 - expanded)
            ),
        ):
            rows.append({**item, "id": uuid4().hex})
            if item.get("container"):
                expanded += item.get("size", 0)
            if len(rows) <= 20 or len(rows) % 100 == 0:
                progress(f"Inspected {len(rows):,} files: " + item["name"] + ": " + item["status"])
    if cancelled():
        from edf_bill_fetcher.application.run_service import RunCancelled

        raise RunCancelled("Intake cancelled")
    write_json(rows, str(target / "intake.json"))
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Intake inventory"
    columns = ("name", "kind", "status", "detail", "size", "origin")
    for column, key in enumerate(columns, 1):
        literal_cell(sheet, 1, column, key.replace("_", " ").title())
        sheet.column_dimensions[sheet.cell(1, column).column_letter].width = (
            25 if column < 4 else 55
        )
    for row_number, row in enumerate(rows, 2):
        for column, key in enumerate(columns, 1):
            literal_cell(sheet, row_number, column, row.get(key))
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    workbook.save(target / "Intake.xlsx")
    workbook.close()
    return rows


def main(argv: list[str]) -> None:
    """Inspect mixed inputs without starting a web server or analysing bills."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("paths", nargs="+")
    parser.add_argument("--output-dir", required=True)
    options = parser.parse_args(argv)
    directory = Path(options.output_dir)
    if directory.exists() and any(directory.iterdir()):
        raise ValueError("Inspect into a fresh empty folder")
    inspect_sources(options.paths, str(directory))
