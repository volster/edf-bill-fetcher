"""CLI entry points — argparse subcommands for headless extraction and report generation."""

from __future__ import annotations

import argparse
import datetime
import importlib.util
import json
import os
import pickle
import sys
import traceback
from typing import Any, cast

from edf_bill_fetcher.models.config import ConfigDict

try:
    import pypff

    HAS_PYPFF = True
except ImportError:
    HAS_PYPFF = False

HAS_TK = importlib.util.find_spec("tkinter") is not None


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
        "pyarrow.lib": None,
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


def _safe_json_dump(state: dict[str, Any], path: str) -> None:
    """Write engine state as versioned JSON: ``{"version": 1, ...}``.

    The payload carries ``config``, ``filtered_records`` and
    ``error_log`` — the same serialization shape the ``--records-json``
    export uses, so both outputs stay interoperable.  ``default=str``
    mirrors that handler and keeps non-primitive values serialisable.

    Usage:  _safe_json_dump(state, "engine.json")
    """
    payload = {
        "version": 1,
        "config": state.get("config", {}),
        "filtered_records": state.get("filtered_records", []),
        "error_log": state.get("error_log", []),
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, default=str)


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
    parser.add_argument("--records-json", help="Also save extracted records as JSON")
    parser.add_argument(
        "--save-engine-json",
        help="Save engine state (config, filtered records, error log) as versioned JSON",
    )
    parser.add_argument("--config", "-c", help="Path to config JSON file (optional)")
    parser.add_argument("--acc-filter", help="Filter by account number (e.g., A-12345678)")
    parser.add_argument(
        "--domain-filter",
        default="edfenergy.com",
        help="Comma-separated sender domains for PST filtering",
    )
    parser.add_argument("--min-amount", type=float, default=50.0, help="Minimum amount threshold")
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
                config = cast(ConfigDict, json.load(f))
        except Exception as e:
            sys.stderr.write(f"ERROR: Failed to load config: {e}\n")
            sys.exit(1)

    # Override with CLI args
    config.update(
        {
            "use_acc_filter": bool(parsed.acc_filter),
            "acc_num": parsed.acc_filter or "",
            "use_domain_filter": True,
            "domain_filter": parsed.domain_filter,
            "min_amount": parsed.min_amount,
            "filter_below": not parsed.no_filter_below,
            "use_dedup": not parsed.no_dedup,
            "use_anchors": not parsed.no_anchors,
            "use_large": not parsed.no_large,
            "use_reading_classification": not parsed.no_reading_class,
            "use_pdf_fields": not parsed.no_pdf_fields,
            "save_filtered": True,
            "save_dups": True,
        }
    )

    # Check PST dependency
    if parsed.pst and not HAS_PYPFF:
        sys.stderr.write(
            "ERROR: PST/OST support requires 'libpff-python'. Install with: pip install libpff-python\n"
        )
        sys.exit(1)

    from edf_bill_fetcher.collectors.engine import EvidenceEngine  # noqa: F402,E402

    engine = EvidenceEngine(config, print, None, None)

    try:
        if parsed.pst and os.path.exists(parsed.pst):
            print(f"Scanning PST/OST: {parsed.pst}")
            try:
                pff = pypff.file()
            except AttributeError:
                pff = getattr(pypff, "File", None)
                if pff is None:
                    raise AttributeError("pypff module has no 'file' or 'File' attribute") from None
                pff = pff()
            pff.open(os.path.abspath(parsed.pst))
            try:
                engine.crawl_pst(pff.get_root_folder())
            finally:
                pff.close()

        if parsed.htm and os.path.exists(parsed.htm):
            print(f"Parsing HTM: {parsed.htm}")
            engine.process_htm_file(parsed.htm)

        if parsed.pdf_dir and os.path.exists(parsed.pdf_dir):
            print(f"Scanning PDF folder: {parsed.pdf_dir}")
            engine.crawl_local_pdfs(parsed.pdf_dir)

        if parsed.eml_dir and os.path.exists(parsed.eml_dir):
            print(f"Scanning EML folder: {parsed.eml_dir}")
            engine.process_eml_folder(parsed.eml_dir)

        if parsed.msg_dir and os.path.exists(parsed.msg_dir):
            print(f"Scanning MSG folder: {parsed.msg_dir}")
            engine.process_msg_folder(parsed.msg_dir)

        if not engine.records:
            sys.stderr.write("WARNING: No billing records found\n")
            sys.exit(1 if parsed.strict else 0)

        # Export to Excel
        print(f"Writing Excel report: {parsed.output}")

        from edf_bill_fetcher.io.writers.export import export_to_excel

        export_to_excel(
            engine.records,
            parsed.output,
            engine.error_log,
            config,
            filtered=engine.filtered_records,
            sap_rows={
                "contract": engine.sap_contract_rows,
                "meter": engine.sap_meter_rows,
                "financial": engine.sap_financial_rows,
            },
        )

        # Optionally save records as JSON
        if parsed.records_json:
            output_data = {
                "extracted_at": datetime.datetime.now().isoformat(),
                "config": config,
                "records": engine.records,
                "filtered_records": engine.filtered_records,
                "error_log": engine.error_log,
            }
            with open(parsed.records_json, "w", encoding="utf-8") as f:
                json.dump(output_data, f, indent=2, default=str)
            print(f"Records saved as JSON: {parsed.records_json}")

        # Optionally save engine state as versioned JSON
        if parsed.save_engine_json:
            state = {
                "config": config,
                "filtered_records": engine.filtered_records,
                "error_log": engine.error_log,
            }
            _safe_json_dump(state, parsed.save_engine_json)
            print(f"Engine state saved as JSON: {parsed.save_engine_json}")

        print("Extraction complete!")
        print(f"  PDFs processed: {engine.pdf_count}")
        print(f"  Emails matched: {engine.email_count}")
        print(f"  Records found:  {len(engine.records)}")
        if engine.error_log:
            print(f"  Parse errors:   {len(engine.error_log)}")

    except Exception as e:
        sys.stderr.write(f"ERROR: {e}\n")
        traceback.print_exc()
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
        help="Path to engine state file (.json preferred, .pkl legacy)",
    )
    parsed = parser.parse_args(args)

    try:
        with open(parsed.records, encoding="utf-8") as f:
            loaded = json.load(f)

        # Accept either a bare list of records (preferred) or the wrapper
        # object emitted by ``--extract --records-json``.  The wrapper
        # shape is ``{"records": [...], ...meta}`` — unwrap it so both
        # CLI entry points behave identically.
        if isinstance(loaded, dict) and "records" in loaded:
            records = loaded["records"]
        else:
            records = loaded

        config: ConfigDict = {}
        if parsed.config:
            with open(parsed.config, encoding="utf-8") as f:
                config = cast(ConfigDict, json.load(f))

        engine = None
        filtered = None
        if parsed.engine_data:
            # Use the restricted unpickler to prevent arbitrary code
            # execution from crafted pickle files (see C1 fix).  A .json
            # path yields the versioned JSON state dict instead of an
            # engine instance; its filtered_records ride along the same way.
            engine = _safe_pickle_load(parsed.engine_data)
            if isinstance(engine, dict):
                filtered = engine.get("filtered_records")
            else:
                filtered = getattr(engine, "filtered_records", None)

        from edf_bill_fetcher.io.reporters.pdf_report import generate_pdf_from_gui

        success, msg = generate_pdf_from_gui(
            records=records,
            output_path=parsed.output,
            config=config,
            engine=engine,
            filtered=filtered,
        )
        if success:
            sys.stdout.write(msg + "\n")
            sys.exit(0)
        else:
            sys.stderr.write(f"ERROR: {msg}\n")
            sys.exit(1)

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
        help="Path to engine state file (.json preferred, .pkl legacy)",
    )
    parsed = parser.parse_args(args)

    try:
        with open(parsed.records, encoding="utf-8") as f:
            loaded = json.load(f)

        # Accept either a bare list of records (preferred) or the wrapper
        # object emitted by ``--extract --records-json``.  Mirrors the
        # PDF CLI loader so both formats round-trip without extra steps.
        if isinstance(loaded, dict) and "records" in loaded:
            records = loaded["records"]
        else:
            records = loaded

        config: ConfigDict = {}
        if parsed.config:
            with open(parsed.config, encoding="utf-8") as f:
                config = cast(ConfigDict, json.load(f))

        engine = None
        filtered = None
        if parsed.engine_data:
            # Use the restricted unpickler to prevent arbitrary code
            # execution from crafted pickle files (see C1 fix).  A .json
            # path yields the versioned JSON state dict instead of an
            # engine instance; its filtered_records ride along the same way.
            engine = _safe_pickle_load(parsed.engine_data)
            if isinstance(engine, dict):
                filtered = engine.get("filtered_records")
            else:
                filtered = getattr(engine, "filtered_records", None)

        from edf_bill_fetcher.io.reporters.docx_report import generate_docx_from_gui

        success, msg = generate_docx_from_gui(
            records=records,
            output_path=parsed.output,
            config=config,
            engine=engine,
            filtered=filtered,
        )
        if success:
            sys.stdout.write(msg + "\n")
            sys.exit(0)
        else:
            sys.stderr.write(f"ERROR: {msg}\n")
            sys.exit(1)
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
        help="Path to engine state file (.json preferred, .pkl legacy)",
    )
    parsed = parser.parse_args(args)

    try:
        with open(parsed.records, encoding="utf-8") as f:
            loaded = json.load(f)

        # Accept either a bare list of records (preferred) or the wrapper
        # object emitted by ``--extract --records-json``.  Mirrors the
        # PDF/DOCX CLI loaders so all three formats round-trip identically.
        if isinstance(loaded, dict) and "records" in loaded:
            records = loaded["records"]
        else:
            records = loaded

        config: ConfigDict = {}
        if parsed.config:
            with open(parsed.config, encoding="utf-8") as f:
                config = cast(ConfigDict, json.load(f))

        engine = None
        filtered = None
        if parsed.engine_data:
            # Use the restricted unpickler to prevent arbitrary code
            # execution from crafted pickle files (see C1 fix).  A .json
            # path yields the versioned JSON state dict instead of an
            # engine instance; its filtered_records ride along the same way.
            engine = _safe_pickle_load(parsed.engine_data)
            if isinstance(engine, dict):
                filtered = engine.get("filtered_records")
            else:
                filtered = getattr(engine, "filtered_records", None)

        from edf_bill_fetcher.io.reporters.html_report import generate_html_from_gui

        success, msg = generate_html_from_gui(
            records=records,
            output_path=parsed.output,
            config=config,
            engine=engine,
            filtered=filtered,
        )
        if success:
            sys.stdout.write(msg + "\n")
            sys.exit(0)
        else:
            sys.stderr.write(f"ERROR: {msg}\n")
            sys.exit(1)
    except Exception as e:
        sys.stderr.write(f"ERROR: {e}\n")
        sys.exit(1)


def main() -> None:
    """Entry point for the EDF Evidence Collector CLI."""
    if len(sys.argv) > 1:
        if sys.argv[1] in ("--pdf-report", "--report", "-r"):
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

    if not HAS_TK:
        sys.stderr.write(
            "ERROR: tkinter is not available in this Python build. "
            "Launch a CLI command instead (e.g. --extract, --pdf-report, "
            "--docx-report, --html-report, --diff) or run on a system with Tk installed."
        )
        sys.stderr.write("\n")
        sys.exit(2)

    import tkinter as tk  # noqa: F401

    root = tk.Tk()
    from edf_bill_fetcher.ui.app import App  # noqa: F402,E402

    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
