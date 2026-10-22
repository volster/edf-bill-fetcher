"""CLI entry points — argparse subcommands for headless extraction and report generation."""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import pickle
import sys
import traceback
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace
from typing import Any, cast

from edf_bill_fetcher.domain.analysis import AnalysisSnapshot
from edf_bill_fetcher.helpers.artifacts import atomic_output, json_safe
from edf_bill_fetcher.models.config import ConfigDict
from edf_bill_fetcher.processors.preparation import prepare_evidence

try:
    import pypff

    HAS_PYPFF = pypff is not None
except ImportError:
    HAS_PYPFF = False

HAS_QT = importlib.util.find_spec("PySide6") is not None


class _RestrictedUnpickler(pickle.Unpickler):
    """Unpickler that only allows known-safe types.

    Permits: built-in scalars, dicts, lists, tuples, sets, frozensets,
    bytes/bytearray, and the project's own ``EvidenceEngine``.  Everything
    else triggers ``pickle.UnpicklingError`` so a crafted pickle can never
    import and call arbitrary code.
    """

    _SAFE_CLASSES: dict[str, set[str] | None] = {
        "builtins": {
            "dict",
            "list",
            "tuple",
            "set",
            "frozenset",
            "int",
            "float",
            "str",
            "bool",
            "bytes",
            "bytearray",
            "NoneType",
            "type",
            "slice",
        },
        "collections": {"OrderedDict", "defaultdict", "Counter", "deque"},
        "collections.__init__": {"OrderedDict", "defaultdict", "Counter", "deque"},
        "pandas.core.series": {"Series"},
        "pandas.core.frame": {"DataFrame"},
        "pandas": {"DataFrame", "Series", "Index", "StringDtype", "RangeIndex"},
        "pandas.arrays": {"ArrowStringArray"},
        "pyarrow.lib": {"DataType", "Table", "Array", "ChunkedArray", "StringArray"},
        "pandas.core.internals.managers": {"BlockManager"},
        "pandas._libs.internals": {"_unpickle_block"},
        "numpy.core.numeric": {"_frombuffer"},
        "numpy._core.numeric": {"_frombuffer"},
        "numpy.dtype": {"dtype"},
        "numpy": {"ndarray", "dtype"},
        "numpy.ndarray": {"ndarray"},
        "numpy.core.multiarray": {"_reconstruct"},
        "numpy._core.multiarray": {"_reconstruct"},
        "pandas.core.indexes.base": {"_new_Index", "Index"},
        "pandas.core.indexes.range": {"RangeIndex"},
        "edf_bill_fetcher.collectors.engine": {"EvidenceEngine"},
    }

    def find_class(self, module: str, name: str) -> type:
        _SENTINEL = object()
        allowed = self._SAFE_CLASSES.get(module, _SENTINEL)
        if allowed is _SENTINEL:
            raise pickle.UnpicklingError(
                f"Blocked unsafe class {module!r}.{name!r} in pickle stream"
            )
        allow_everything = allowed is None
        if allow_everything or (isinstance(allowed, set) and name in allowed):
            if module == "edf_bill_fetcher.collectors.engine":
                import importlib

                mod = importlib.import_module("edf_bill_fetcher.collectors.engine")
                cls = getattr(mod, name)
                if not isinstance(cls, type):
                    raise pickle.UnpicklingError(
                        f"Resolved edf_bill_fetcher.collectors.engine attribute {name!r} is not a class"
                    )
                return cls
            return cast(type, super().find_class(module, name))
        raise pickle.UnpicklingError(f"Blocked unsafe class {module!r}.{name!r} in pickle stream")


def _safe_pickle_load(path: str) -> Any:
    """Load engine state from a versioned JSON or legacy pickle file.

    Dispatch on the file extension: ``.json`` is read through plain
    ``json.load`` (the versioned state written by ``_safe_json_dump``),
    ``.pkl`` through the restricted unpickler, and anything else raises
    ``ValueError`` so a typo'd extension never silently falls back.

    Usage:  state = _safe_pickle_load("engine.json")
    Raises ValueError for an unrecognised extension and
    pickle.UnpicklingError for disallowed types in a pickle stream.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".json":
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    if ext == ".pkl":
        with open(path, "rb") as f:
            return _RestrictedUnpickler(f).load()
    raise ValueError(
        f"Unsupported engine state file extension {ext or '(none)'!r} for {path!r}; "
        "use .json (versioned JSON) or .pkl (legacy pickle)"
    )


@atomic_output("path")
def _safe_json_dump(state: dict[str, Any], path: str) -> None:
    """Write preserved engine state as version 2 JSON.

    The payload carries ``config``, ``filtered_records`` and
    ``error_log`` — the same serialization shape the ``--records-json``
    export uses, so both outputs stay interoperable.  ``default=str``
    mirrors that handler and keeps non-primitive values serialisable.

    Usage:  _safe_json_dump(state, "engine.json")
    """
    payload = dict(state)
    payload["version"] = 2
    payload.setdefault("extracted_at", datetime.now(timezone.utc).isoformat())
    prepared = prepare_evidence(
        payload.get("original_records", payload.get("records", [])), payload.get("config", {})
    )
    payload["config"] = prepared.config
    payload["records"] = prepared.records
    payload["duplicate_records"] = prepared.duplicates
    payload["status"] = "review_required" if payload.get("error_log") else "complete"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(json_safe(payload), f, indent=2, default=str, allow_nan=False)


def _load_report_inputs(parsed):
    """Restore records, configuration, source counts and diagnostics from JSON."""
    with open(parsed.records, encoding="utf-8") as handle:
        loaded = json.load(handle)
    metadata = loaded if isinstance(loaded, dict) else {}
    if metadata.get("version", 1) not in (1, 2):
        raise ValueError("Unsupported evidence JSON version")
    records = metadata.get("records", []) if metadata else loaded
    if not isinstance(records, list) or any(not isinstance(row, dict) for row in records):
        raise ValueError("Records must be a JSON list of objects")
    if parsed.engine_data:
        if os.path.splitext(parsed.engine_data)[1].lower() != ".json":
            raise ValueError(
                "Reports accept JSON engine data; legacy pickle is not a safe interchange format"
            )
        state = _safe_pickle_load(parsed.engine_data)
        if not isinstance(state, dict) or state.get("version", 1) not in (1, 2):
            raise ValueError("Unsupported engine JSON state")
        metadata = {**metadata, **state}
    config = dict(metadata.get("config") or {})
    if parsed.config:
        with open(parsed.config, encoding="utf-8") as handle:
            overrides = json.load(handle)
        if not isinstance(overrides, dict):
            raise ValueError("Configuration must be a JSON object")
        config.update(overrides)
    engine = SimpleNamespace(
        **{
            **{
                "pdf_count": 0,
                "email_count": 0,
                "error_log": [],
                "filtered_records": [],
                "source_manifest": {},
            },
            **metadata,
        }
    )
    return records, config, engine, engine.filtered_records


def run_cli_extract(args: list[str]) -> None:
    """Run extraction from command line (headless mode)."""
    parser = argparse.ArgumentParser(
        description=(
            "Extract EDF billing data from PST/OST, PDF folder, HTM export, or EML/MSG folders"
        ),
        prog="edf-collector --extract",
    )
    parser.add_argument("--pst", help="Path to PST/OST file")
    parser.add_argument("--pdf-dir", help="Path to directory containing PDF bills")
    parser.add_argument("--htm", help="Path to HTM account history export")
    parser.add_argument("--eml-dir", help="Path to directory containing .eml message files")
    parser.add_argument("--msg-dir", help="Path to directory containing .msg message files")
    parser.add_argument("--output", "-o", required=True, help="Output Excel file path")
    parser.add_argument(
        "--records-json", help="Also save extracted records as v2 compatibility JSON"
    )
    parser.add_argument(
        "--save-engine-json",
        help="Save engine state as v2 compatibility JSON (v3 result is always saved in the source bundle)",
    )
    parser.add_argument("--config", "-c", help="Path to config JSON file (optional)")
    parser.add_argument("--acc-filter", help="Filter by account number (e.g., A-12345678)")
    parser.add_argument(
        "--domain-filter",
        default=None,
        help="Comma-separated sender domains for PST filtering",
    )
    parser.add_argument("--min-amount", type=float, default=None, help="Minimum amount threshold")
    parser.add_argument("--no-dedup", action="store_true", help="Disable deduplication")
    parser.add_argument("--no-anchors", action="store_true", help="Disable smart context search")
    parser.add_argument("--no-large", action="store_true", help="Disable large amount fallback")
    parser.add_argument(
        "--no-reading-class", action="store_true", help="Disable reading classification"
    )
    parser.add_argument(
        "--no-pdf-fields", action="store_true", help="Disable deep PDF field extraction"
    )
    parser.add_argument(
        "--no-filter-below", action="store_true", help="Don't filter records below minimum amount"
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit with code 1 when no billing records are found (for CI use). "
        "Without this flag a clean-empty run exits 0.",
    )
    parsed = parser.parse_args(args)

    # Check at least one source
    if not any([parsed.pst, parsed.pdf_dir, parsed.htm, parsed.eml_dir, parsed.msg_dir]):
        sys.stderr.write(
            "ERROR: At least one source required "
            "(--pst, --pdf-dir, --htm, --eml-dir, or --msg-dir)\n"
        )
        sys.exit(1)

    # Load config from file if provided
    config: ConfigDict = {}
    if parsed.config:
        try:
            with open(parsed.config, encoding="utf-8") as f:
                loaded_config = json.load(f)
                if not isinstance(loaded_config, dict):
                    raise ValueError("Configuration must be a JSON object")
                config = cast(ConfigDict, loaded_config)
        except Exception as e:
            sys.stderr.write(f"ERROR: Failed to load config: {e}\n")
            sys.exit(1)

    config = cast(
        ConfigDict,
        {
            "use_acc_filter": False,
            "acc_num": "",
            "use_domain_filter": True,
            "domain_filter": "edfenergy.com",
            "min_amount": 50.0,
            "filter_below": False,
            "use_dedup": True,
            "use_anchors": True,
            "use_large": True,
            "use_reading_classification": True,
            "use_pdf_fields": True,
            "save_filtered": True,
            "save_dups": True,
            **config,
        },
    )
    if parsed.acc_filter is not None:
        config.update({"use_acc_filter": bool(parsed.acc_filter), "acc_num": parsed.acc_filter})
    if parsed.domain_filter is not None:
        config.update({"use_domain_filter": True, "domain_filter": parsed.domain_filter})
    if parsed.min_amount is not None:
        config["min_amount"] = parsed.min_amount
    for flag, key in (
        ("no_dedup", "use_dedup"),
        ("no_anchors", "use_anchors"),
        ("no_large", "use_large"),
        ("no_reading_class", "use_reading_classification"),
        ("no_pdf_fields", "use_pdf_fields"),
        ("no_filter_below", "filter_below"),
    ):
        if getattr(parsed, flag):
            cast(dict[str, Any], config)[key] = False
    try:
        config = prepare_evidence([], config).config
    except ValueError as error:
        parser.error(str(error))

    for source in (parsed.pst, parsed.htm, parsed.pdf_dir, parsed.eml_dir, parsed.msg_dir):
        if source and not os.path.exists(source):
            parser.error(f"Source does not exist: {source}")

    # Check PST dependency
    if parsed.pst and not HAS_PYPFF:
        sys.stderr.write(
            "ERROR: PST/OST support requires 'libpff-python'. Install with: pip install libpff-python\n"
        )
        sys.exit(1)

    from edf_bill_fetcher.collectors.engine import EvidenceEngine  # noqa: F402,E402

    engine = EvidenceEngine(config, print, None, None)

    import tempfile
    from pathlib import Path
    from uuid import uuid4

    from edf_bill_fetcher.application.run_service import bundle_engine_sources, collect_engine, run
    from edf_bill_fetcher.domain.json_values import freeze_json
    from edf_bill_fetcher.domain.requests import RunRequest, SourceSelection

    sources = tuple(
        SourceSelection(kind, path)
        for kind, path in (
            ("pst", parsed.pst),
            ("htm", parsed.htm),
            ("pdf_dir", parsed.pdf_dir),
            ("eml_dir", parsed.eml_dir),
            ("msg_dir", parsed.msg_dir),
        )
        if path
    )
    request = RunRequest(
        uuid4().hex,
        datetime.now(timezone.utc),
        datetime.fromisoformat(config["as_of"]).date(),
        sources,
        freeze_json(dict(config)),
    )
    coordinator_started = False
    try:
        output_parent = Path(parsed.output).absolute().parent
        bundle_dir = Path(
            tempfile.mkdtemp(prefix=Path(parsed.output).stem + "_", dir=output_parent)
        )
        legacy_paths = tuple(
            dict.fromkeys(
                str(path)
                for path in (
                    bundle_dir / "records.json",
                    parsed.records_json,
                    parsed.save_engine_json,
                )
                if path
            )
        )
        destinations = [Path(path).resolve() for path in (*legacy_paths, parsed.output)]
        if len(set(destinations)) != len(destinations) or any(
            Path(source.path).resolve() in destinations for source in sources
        ):
            raise ValueError(
                "Output paths must be distinct and must not overwrite selected sources"
            )
        coordinator_started = True
        outcome = run(
            request,
            {"xlsx": parsed.output},
            result_path=str(bundle_dir / "result.json"),
            manifest_path=str(bundle_dir / "run.json"),
            collect=lambda: collect_engine(
                engine, request, pff_module=pypff if HAS_PYPFF else None
            ),
            copy_sources=lambda prepared: bundle_engine_sources(
                engine, prepared, str(bundle_dir), legacy_paths=legacy_paths
            ),
            cleanup=engine.cleanup,
            progress=lambda event: print(event.message),
        )
        if outcome.status == "empty":
            sys.stderr.write("WARNING: No billing records found\n")
            sys.exit(1 if parsed.strict else 0)
        if outcome.status not in {"complete", "review_required"}:
            raise RuntimeError(outcome.error or outcome.status)
        print(f"Extraction complete! Status: {outcome.status}. Workbook: {parsed.output}")
        if engine.error_log:
            print(f"  Parse errors:   {len(engine.error_log)}")
        print(f"Saved analysis: {bundle_dir / 'result.json'}")
        for path in legacy_paths:
            print(f"Compatibility records: {path}")
        if parsed.records_json:
            print(f"Records saved as JSON: {parsed.records_json}")
        if parsed.save_engine_json:
            print(f"Engine state saved as JSON: {parsed.save_engine_json}")
        print(
            f"PDFs processed: {engine.pdf_count}; emails matched: {engine.email_count}; records: {len(engine.records)}"
        )
    except Exception as e:
        sys.stderr.write(f"ERROR: {e}\n")
        traceback.print_exc()
        if not coordinator_started:
            engine.cleanup()
        sys.exit(1)


def _diff_amount_text(value: Any) -> str:
    """Render an amount cell as ``£x.xx`` (2dp) with a dash for missing values."""
    if value is None or (isinstance(value, float) and value != value):
        return "-"
    try:
        return f"£{float(value):.2f}"
    except (TypeError, ValueError):
        return str(value)


def _diff_field_text(field: str, value: Any) -> str:
    """Render one changed-field value, prefixing amount-shaped fields with £."""
    if field == "Amount (£)":
        return _diff_amount_text(value)
    return "-" if value is None else str(value)


def run_cli_diff(args: list[str]) -> None:
    """Compare two records JSON files and print/report the delta.

    Usage:  edf-collector --diff OLD.json NEW.json [--diff-output out.xlsx]
    """
    parser = argparse.ArgumentParser(
        description="Compare two records JSON exports (added / removed / changed)",
        prog="edf-collector --diff",
    )
    parser.add_argument("old", help="Path to the OLD records JSON file")
    parser.add_argument("new", help="Path to the NEW records JSON file")
    parser.add_argument(
        "--diff-output",
        help="Optional path for an Excel workbook with added/removed/changed sheets",
    )
    parsed = parser.parse_args(args)

    try:
        with open(parsed.old, encoding="utf-8") as f:
            old_loaded = json.load(f)
        with open(parsed.new, encoding="utf-8") as f:
            new_loaded = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        sys.stderr.write(f"ERROR: {e}\n")
        sys.exit(1)

    old_records = (
        old_loaded["records"]
        if isinstance(old_loaded, dict) and "records" in old_loaded
        else old_loaded
    )
    new_records = (
        new_loaded["records"]
        if isinstance(new_loaded, dict) and "records" in new_loaded
        else new_loaded
    )
    if not isinstance(old_records, list) or not isinstance(new_records, list):
        sys.stderr.write("ERROR: both diff inputs must be records.json lists\n")
        sys.exit(1)

    from edf_bill_fetcher.processors.run_diff import diff_records

    diff = diff_records(old_records, new_records)

    print(f"Diff: {parsed.old} -> {parsed.new}")
    print(f"Added:   {len(diff['added'])}")
    print(f"Removed: {len(diff['removed'])}")
    print(f"Changed: {len(diff['changed'])}")
    for row in diff["added"]:
        print(
            f"+ ADDED {row.get('Date', '-')} {_diff_amount_text(row.get('Amount (£)'))} {row.get('Source', '-')}"
        )
    for row in diff["removed"]:
        print(
            f"- REMOVED {row.get('Date', '-')} {_diff_amount_text(row.get('Amount (£)'))} {row.get('Source', '-')}"
        )
    for entry in diff["changed"]:
        new_row = entry["new_row"]
        details = ", ".join(
            f"{field}: {_diff_field_text(field, old)} → {_diff_field_text(field, new)}"
            for field, old, new in zip(
                entry["changed_fields"], entry["old_values"], entry["new_values"], strict=True
            )
        )
        print(
            f"~ CHANGED {new_row.get('Date', '-')} "
            f"{_diff_amount_text(new_row.get('Amount (£)'))} "
            f"{new_row.get('Source', '-')} [{details}]"
        )

    if parsed.diff_output:
        from edf_bill_fetcher.io.writers.diff import write_diff_workbook

        write_diff_workbook(diff, parsed.diff_output)
        print(f"Diff workbook written: {parsed.diff_output}")


def _render_cli_report(parsed: argparse.Namespace, fmt: str) -> None:
    """Restore accepted analysis or explicitly create a revised result for overrides."""
    from dataclasses import replace
    from pathlib import Path
    from uuid import uuid4

    from edf_bill_fetcher.application.run_service import check_source_destinations, run
    from edf_bill_fetcher.domain.json_values import freeze_json, object_value, strict_loads
    from edf_bill_fetcher.domain.requests import RunRequest, SourceSelection
    from edf_bill_fetcher.io.persistence.result_codec import decode_result, migrate_legacy
    from edf_bill_fetcher.io.writers.evidence_bundle import resolve_bundle_paths
    from edf_bill_fetcher.models.report_view import RenderSelection

    payload = strict_loads(Path(parsed.records).read_text(encoding="utf-8"))
    overrides = (
        strict_loads(Path(parsed.config).read_text(encoding="utf-8")) if parsed.config else {}
    )
    if not isinstance(overrides, dict):
        raise ValueError("Configuration must be a JSON object")
    snapshot: AnalysisSnapshot | None = None
    if isinstance(payload, dict) and "schema_version" in payload:
        snapshot = decode_result(payload)
        if parsed.engine_data:
            raise ValueError("A v3 result already contains its engine metadata; omit --engine-data")
        cfg = object_value(snapshot.prepared.effective_config)
        changes = {key: value for key, value in overrides.items() if cfg.get(key) != value}
        cfg.update(overrides)
        if set(changes) <= {"report_sections", "report_account_ref"}:
            snapshot = replace(
                snapshot, prepared=replace(snapshot.prepared, effective_config=freeze_json(cfg))
            )
        else:
            state = payload["evidence"]["legacy_state"]
            payload = {
                **state,
                "version": 2,
                "records": payload["evidence"]["prepared_records"],
                "original_records": [row["raw"] for row in payload["evidence"]["observations"]],
            }
            snapshot = None
    else:
        snapshot = None
        if isinstance(payload, list):
            payload = {"records": payload}
        if not isinstance(payload, dict):
            raise ValueError("Records must be an evidence object or array")
        if parsed.engine_data:
            if Path(parsed.engine_data).suffix.lower() != ".json":
                raise ValueError(
                    "Reports accept JSON engine data; pickle is not an interchange format"
                )
            state = strict_loads(Path(parsed.engine_data).read_text(encoding="utf-8"))
            if not isinstance(state, dict) or state.get("version", 1) not in (1, 2):
                raise ValueError("Unsupported engine JSON state")
            for key in ("records", "original_records"):
                if key in state and key in payload and state[key] != payload[key]:
                    raise ValueError("Conflicting evidence and engine records")
            payload = {**payload, **state}
        cfg = dict(payload.get("config") or {})
        cfg.update(overrides)
    if snapshot is None:
        cfg.setdefault("as_of", datetime.now(timezone.utc).date().isoformat())
        captured = RunRequest(
            uuid4().hex,
            datetime.now(timezone.utc),
            datetime.fromisoformat(cfg["as_of"]).date(),
            (SourceSelection("records_json", str(Path(parsed.records).absolute())),),
            freeze_json(cfg),
        )
        snapshot = migrate_legacy(payload, captured)
    request = snapshot.prepared.extraction.request
    selection = RenderSelection(
        (fmt,), tuple(cfg["report_sections"]) if "report_sections" in cfg else None
    )
    result_path, manifest_path = parsed.output + ".result.json", parsed.output + ".run.json"
    protected = {
        Path(path).resolve() for path in (parsed.records, parsed.engine_data, parsed.config) if path
    }
    if any(
        Path(path).resolve() in protected for path in (parsed.output, result_path, manifest_path)
    ):
        raise ValueError("Report output would overwrite an input file")
    state = object_value(snapshot.prepared.extraction.legacy_state)
    bundle_root = Path(parsed.records).resolve().parent
    paths, warnings = resolve_bundle_paths(state, str(bundle_root))
    check_source_destinations(
        request,
        [Path(path).resolve() for path in (parsed.output, result_path, manifest_path)],
        {
            **state,
            "source_bundle_path": str(bundle_root),
            "saved_evidence_paths": {**state.get("saved_evidence_paths", {}), **paths},
        },
    )
    outcome = run(
        request,
        {fmt: parsed.output},
        result_path=result_path,
        manifest_path=manifest_path,
        accepted_result=snapshot,
        selection=selection,
        warnings=warnings,
    )
    if outcome.status not in {"complete", "review_required"}:
        raise RuntimeError(outcome.error or "No records to report on")
    sys.stdout.write(
        f"{fmt.upper()} report generated: {parsed.output}\nSaved analysis: {result_path}\n"
    )
    for warning in warnings:
        sys.stderr.write(f"WARNING: {warning}\n")


def run_cli_pdf_report(args: list[str]) -> None:
    """Run PDF report generation from command line."""
    parser = argparse.ArgumentParser(
        description="Generate PDF report from extracted records",
        prog="edf-collector --pdf-report",
    )
    parser.add_argument(
        "--records",
        "-i",
        required=True,
        help="Path to extracted records JSON file (exported from GUI or script)",
    )
    parser.add_argument("--output", "-o", required=True, help="Output PDF file path")
    parser.add_argument("--config", "-c", help="Path to config JSON file (optional)")
    parser.add_argument(
        "--engine-data",
        "-e",
        help="Path to versioned engine state JSON",
    )
    parsed = parser.parse_args(args)

    try:
        _render_cli_report(parsed, "pdf")
        sys.exit(0)

    except Exception as e:
        sys.stderr.write(f"ERROR: {e}\n")
        sys.exit(1)


def run_cli_docx_report(args: list[str]) -> None:
    """Run DOCX report generation from command line."""
    parser = argparse.ArgumentParser(
        description="Generate DOCX report from extracted records",
        prog="edf-collector --docx-report",
    )
    parser.add_argument(
        "--records",
        "-i",
        required=True,
        help="Path to extracted records JSON file (exported from GUI or script)",
    )
    parser.add_argument("--output", "-o", required=True, help="Output DOCX file path")
    parser.add_argument("--config", "-c", help="Path to config JSON file (optional)")
    parser.add_argument(
        "--engine-data",
        "-e",
        help="Path to versioned engine state JSON",
    )
    parsed = parser.parse_args(args)

    try:
        _render_cli_report(parsed, "docx")
        sys.exit(0)

    except Exception as e:
        sys.stderr.write(f"ERROR: {e}\n")
        sys.exit(1)


def run_cli_html_report(args: list[str]) -> None:
    """Run HTML report generation from command line."""
    parser = argparse.ArgumentParser(
        description="Generate HTML report from extracted records",
        prog="edf-collector --html-report",
    )
    parser.add_argument(
        "--records",
        "-i",
        required=True,
        help="Path to extracted records JSON file (exported from GUI or script)",
    )
    parser.add_argument("--output", "-o", required=True, help="Output HTML file path")
    parser.add_argument("--config", "-c", help="Path to config JSON file (optional)")
    parser.add_argument(
        "--engine-data",
        "-e",
        help="Path to versioned engine state JSON",
    )
    parsed = parser.parse_args(args)

    try:
        _render_cli_report(parsed, "html")
        sys.exit(0)

    except Exception as e:
        sys.stderr.write(f"ERROR: {e}\n")
        sys.exit(1)


def run_cli_review(args: list[str]) -> None:
    """Regenerate a fresh evidence pack from an amended review workbook."""
    from edf_bill_fetcher.application.workbook_review import regenerate_workbook

    parser = argparse.ArgumentParser(prog="edf-collector --review-workbook")
    parser.add_argument("--workbook", "-i", required=True, help="Amended EDF review .xlsx")
    parser.add_argument(
        "--output-dir", "-o", required=True, help="Parent folder for a new reviewed pack"
    )
    parser.add_argument(
        "--formats",
        default="xlsx,pdf,docx,html",
        help="Comma-separated formats; an updated xlsx is always included",
    )
    parsed = parser.parse_args(args)
    try:
        directory, outcome = regenerate_workbook(
            parsed.workbook, parsed.output_dir, formats=tuple(parsed.formats.split(","))
        )
        print(f"Reviewed pack: {directory}")
        print(f"Status: {outcome.status}")
        if outcome.error:
            raise ValueError(outcome.error)
    except Exception as error:
        sys.stderr.write(f"ERROR: {error}\n")
        sys.exit(1)


def main() -> None:
    """Entry point for the EDF Evidence Collector CLI."""
    if len(sys.argv) > 1 and sys.argv[1] == "--doctor":
        from edf_bill_fetcher.helpers.capabilities import main as doctor

        sys.exit(doctor(sys.argv[2:]))

    if len(sys.argv) > 1:
        if sys.argv[1] == "--review-workbook":
            run_cli_review(sys.argv[2:])
            return
        elif sys.argv[1] in ("--pdf-report", "--report", "-r"):
            run_cli_pdf_report(sys.argv[2:])
            return
        elif sys.argv[1] in ("--docx-report", "--word-report", "-w"):
            run_cli_docx_report(sys.argv[2:])
            return
        elif sys.argv[1] in ("--html-report",):
            run_cli_html_report(sys.argv[2:])
            return
        elif sys.argv[1] in ("--extract", "-e"):
            run_cli_extract(sys.argv[2:])
            return
        elif sys.argv[1] in ("--diff", "-d"):
            run_cli_diff(sys.argv[2:])
            return

    if not HAS_QT:
        sys.stderr.write(
            "ERROR: PySide6 is not installed. "
            "Launch a CLI command instead (e.g. --extract, --pdf-report, "
            "--docx-report, --html-report, --diff) or install the project's runtime dependencies."
        )
        sys.stderr.write("\n")
        sys.exit(2)

    if len(sys.argv) > 1 and sys.argv[1] == "--check-gui":
        # Build-only health probe: the default app still opens normally.
        parser = argparse.ArgumentParser(prog="edf-collector --check-gui")
        parser.add_argument("--report", type=Path, required=True)
        options = parser.parse_args(sys.argv[2:])
        from edf_bill_fetcher.ui.qt_smoke import check_gui

        sys.exit(check_gui(options.report, platform="cocoa" if sys.platform == "darwin" else None))

    from edf_bill_fetcher.ui.qt_app import launch

    sys.exit(launch())


if __name__ == "__main__":
    main()
