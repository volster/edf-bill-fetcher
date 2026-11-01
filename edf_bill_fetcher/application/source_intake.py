"""Route a mixed folder/archive through the existing collectors, preserving provenance."""

from __future__ import annotations

import tempfile
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from edf_bill_fetcher.domain.json_values import freeze_json
from edf_bill_fetcher.domain.requests import RunRequest, SourceSelection
from edf_bill_fetcher.io.adapters.intake import IntakeLimits, discover


def collect_mixed(
    engine: Any, path: str, *, pff_module: Any = None, plan: list[dict] | None = None
) -> None:
    """Collect recognized originals; saved results/workbooks require an explicit import choice."""
    from edf_bill_fetcher.application.run_service import collect_engine

    if engine._attachment_dir is None:
        engine._attachment_dir = tempfile.TemporaryDirectory(prefix="edf-intake-")
    staging = Path(engine._attachment_dir.name) / "mixed"
    remaining = 100_000 - len(engine.intake_inventory)
    if remaining <= 0:
        raise ValueError("Mixed intake exceeds the aggregate file-count limit")
    expanded = sum(item.get("size", 0) for item in engine.intake_inventory if item.get("container"))
    items = (
        plan
        if plan is not None
        else discover(
            Path(path),
            staging,
            cancelled=engine.is_cancelled,
            limits=IntakeLimits(files=remaining, expanded_bytes=max(0, 64 * 1024**3 - expanded)),
        )
    )
    for item in items:
        if engine.is_cancelled():
            return
        engine.intake_inventory.append({key: value for key, value in item.items() if key != "path"})
        if item["status"] != "ready":
            if item["status"] not in {"container", "skipped"}:
                engine.log_error(
                    "Intake " + item["name"],
                    item["detail"] or "Choose an explicit saved-result/workbook import workflow",
                )
            continue
        current = Path(item["path"])
        if not engine._can_read_source(str(current)):
            engine.log_error(
                "Intake " + item["name"], "Source is outside the configured evidence folders"
            )
            continue
        if item.get("size") is not None and (
            not current.is_file()
            or current.stat().st_size != item["size"]
            or current.stat().st_mtime_ns != item["mtime_ns"]
        ):
            engine.log_error(
                "Intake " + item["name"],
                "Source changed after inspection; inspect it again before extracting",
            )
            continue
        context = {
            "Intake Origin": str(item.get("origin", Path(path).resolve())) + "!/" + item["name"],
            "Intake Member": item["name"],
        }
        if item["container"]:
            context.update(item["container"])
        engine._intake_origins[str(current.resolve())] = context
        engine._intake_current = context
        engine._source_context = {}
        kind = item["kind"]
        try:
            if kind == "pdf":
                engine.pdf_count += 1
                engine.process_pdf_file(
                    str(current),
                    "Mixed PDF evidence",
                    current.name,
                    "N/A",
                    attachment_name=current.name,
                )
            elif kind == "htm":
                engine.process_htm_file(str(current))
            elif kind in {"eml", "msg"}:
                from edf_bill_fetcher.io.adapters.eml import parse_eml_message
                from edf_bill_fetcher.io.adapters.msg import parse_msg_message

                # A private one-message folder lets the shared message pipeline
                # preserve its sender gates, attachment handling and diagnostics.
                folder = staging / ("message-" + str(len(engine.intake_inventory)))
                folder.mkdir(parents=True, exist_ok=True)
                target = folder / (
                    current.name
                    if current.suffix.lower() == "." + kind
                    else current.name + "." + kind
                )
                import shutil

                shutil.copyfile(current, target)
                engine._intake_origins[str(target.resolve())] = {
                    **context,
                    "Intake Original Filename": current.name,
                }
                engine._process_message_folder(
                    str(folder),
                    "." + kind,
                    parse_eml_message if kind == "eml" else parse_msg_message,
                    kind.upper(),
                )
            else:
                request = RunRequest(
                    "mixed-archive",
                    datetime.now(timezone.utc),
                    date.today(),
                    (SourceSelection("pst", str(current)),),
                    freeze_json({"as_of": date.today().isoformat()}),
                )
                collect_engine(engine, request, pff_module=pff_module)
        except Exception as error:
            engine.log_error("Intake " + item["name"], error)
        finally:
            engine._intake_current = {}
