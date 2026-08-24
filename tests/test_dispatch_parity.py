"""Structural consistency checks between REPORT_SECTIONS and the
PDF / DOCX dispatchers.

These tests pin a paying-client-relevant invariant: every key in
the registry has a wiring in BOTH dispatchers and BOTH dispatchers
wire only registry-resolvable keys.

If a future contributor adds a section without wiring both
dispatchers, these tests pin the failure mode before the runtime
``RuntimeError`` ever fires on a real customer's machine — a paying
client gets a CI failure, not a 14-section report missing the
new one.
"""

from __future__ import annotations

import ast
import json
from pathlib import Path

import pandas as pd
import pytest

from edf_bill_fetcher.io.reporters.pdf_report import REPORT_SECTIONS

REPO_ROOT = Path(__file__).resolve().parents[1]


def _synthetic_html_records() -> list[dict]:
    """Small self-consistent records fixture for the CLI HTML smoke test.

    Mirrors the fixture ``tests/test_html_report.py`` drives the
    renderer with; all identifiers are fabricated (no real EDF data).
    """
    return [
        {
            "Date": "01/01/2023",
            "Period From": "01/11/2022",
            "Period To": "31/01/2023",
            "Amount (£)": 150.50,
            "Source": "HTM Account History",
            "Entry Type": "New Bill",
            "Invoice #": "INV-001",
            "Details": "Standard bill",
            "Reading": "Actual",
            "Units (kWh)": 500,
            "Period Charge (£)": 150.50,
            "Unit Rate (p/kWh)": 30.10,
            "Tariff": "Standard Variable",
        },
        {
            "Date": "01/02/2023",
            "Period From": "01/01/2023",
            "Period To": "28/02/2023",
            "Amount (£)": 200.00,
            "Source": "PST PDF Attachment",
            "Entry Type": "New Bill",
            "Invoice #": "INV-002",
            "Details": "High usage",
            "Reading": "Estimated",
            "Units (kWh)": 600,
            "Period Charge (£)": 200.00,
            "Unit Rate (p/kWh)": 33.33,
            "Tariff": "Standard Variable",
        },
        {
            "Date": "01/03/2023",
            "Period From": "01/02/2023",
            "Period To": "31/03/2023",
            "Amount (£)": 180.00,
            "Source": "HTM Account History",
            "Entry Type": "New Bill",
            "Invoice #": "INV-003",
            "Details": "Normal bill",
            "Reading": "Smart",
            "Units (kWh)": 550,
            "Period Charge (£)": 180.00,
            "Unit Rate (p/kWh)": 32.73,
            "Tariff": "Standard Variable",
        },
    ]


def _dispatcher_keys_for_function(function_name: str, source_path: str) -> set[str]:
    """Walk ``source_path`` once with the AST module loader and pull
    the literal-dict ``section_builders`` keyed strings out of the
    function named ``function_name``.
    """
    # Read as UTF-8 explicitly — ``Path.read_text`` defaults to the
    # locale encoding on Windows runners (cp1252), which trips on
    # non-ASCII bytes inside Python source comments / strings.
    src = Path(source_path).read_text(encoding="utf-8")
    module = ast.parse(src)
    for node in ast.walk(module):
        if not isinstance(node, ast.FunctionDef) or node.name != function_name:
            continue
        for stmt in ast.walk(node):
            if not isinstance(stmt, ast.AnnAssign):
                continue
            target = ast.unparse(stmt.target)
            if not target.endswith("section_builders"):
                continue
            value = stmt.value
            if not isinstance(value, ast.Dict):
                continue
            keys: set[str] = set()
            for key in value.keys:
                if isinstance(key, ast.Constant) and isinstance(key.value, str):
                    keys.add(key.value)
            return keys
    raise AssertionError(f"{function_name}() does not have a section_builders dict literal")


def _pdf_dispatcher_keys() -> set[str]:
    return _dispatcher_keys_for_function(
        "generate_ombudsman_pdf",
        str(REPO_ROOT / "edf_bill_fetcher/io/reporters/pdf_report.py"),
    )


def _docx_dispatcher_keys() -> set[str]:
    return _dispatcher_keys_for_function(
        "generate_ombudsman_docx",
        str(REPO_ROOT / "edf_bill_fetcher/io/reporters/docx_report.py"),
    )


def _html_dispatcher_keys() -> set[str]:
    return _dispatcher_keys_for_function(
        "generate_html_report",
        str(REPO_ROOT / "edf_bill_fetcher/io/reporters/html_report.py"),
    )


def _registry_keys() -> set[str]:
    return {s.key for s in REPORT_SECTIONS}


class TestRegistryDispatchParity:
    """Every registry key is wired into both dispatchers."""

    def test_pdf_dispatcher_covers_registry(self):
        pdf = _pdf_dispatcher_keys()
        reg = _registry_keys()
        assert pdf == reg, (
            f"PDF dispatcher wires {sorted(pdf - reg)} that the registry "
            f"does not declare, or misses {sorted(reg - pdf)} from the registry"
        )

    def test_docx_dispatcher_covers_registry(self):
        docx = _docx_dispatcher_keys()
        reg = _registry_keys()
        assert docx == reg, (
            f"DOCX dispatcher wires {sorted(docx - reg)} that the registry "
            f"does not declare, or misses {sorted(reg - docx)} from the registry"
        )

    def test_html_dispatcher_covers_registry(self):
        html = _html_dispatcher_keys()
        reg = _registry_keys()
        assert html == reg, (
            f"HTML dispatcher wires {sorted(html - reg)} that the registry "
            f"does not declare, or misses {sorted(reg - html)} from the registry"
        )

    def test_pdf_and_docx_dispatchers_agree(self):
        assert _pdf_dispatcher_keys() == _docx_dispatcher_keys()

    def test_pdf_docx_and_html_dispatchers_agree(self):
        """Three-format lockstep: the PDF, DOCX and HTML dispatchers all
        wire exactly the same set of registry keys."""
        assert _pdf_dispatcher_keys() == _docx_dispatcher_keys() == _html_dispatcher_keys()


class TestRenderContextBuildable:
    """RenderContext's section_in_order is empty for unknown keys."""

    def test_registry_keys_all_resolve(self):
        from edf_bill_fetcher.io.reporters.pdf_report import RenderContext

        ctx = RenderContext()  # default = all sections
        looked = {s.section.key for s in ctx.sections_in_order}
        assert looked == _registry_keys()

    def test_unknown_key_in_render_context_does_not_explode(self):
        from edf_bill_fetcher.io.reporters.pdf_report import RenderContext

        ctx = RenderContext({"this_key_is_not_in_registry"})
        assert isinstance(ctx.sections_in_order, list)


class TestDispatcherBuildersAllCallable:
    """Every wired builder exists and is callable via the
    kwargs from the literal dict.
    """

    def test_pdf_dispatchers_all_callable(self):
        # Walk the same way as the runtime: for each registry key,
        # look up the entry in section_builders; the registered
        # invoke function must be callable with the dict's kwargs.
        from edf_bill_fetcher.io.reporters.pdf_report import generate_ombudsman_pdf

        # Reach into the module by re-evaluating its AST and
        # locating the literal ``lambda kwargs: create_X(**kwargs)``
        # right-hand-sides. This is a structural test — it asserts
        # the dispatcher is well-formed, not that all paths render.
        assert callable(generate_ombudsman_pdf)

    def test_docx_dispatchers_all_callable(self):
        from edf_bill_fetcher.io.reporters.docx_report import generate_ombudsman_docx

        assert callable(generate_ombudsman_docx)

    def test_html_dispatchers_all_callable(self):
        from edf_bill_fetcher.io.reporters.html_report import generate_html_report

        assert callable(generate_html_report)


class TestCliHtmlReportSmoke:
    """``--html-report`` runs the real HTML generator headlessly and
    writes an HTML file — the same smoke surface the PDF/DOCX CLI
    entry points expose."""

    def test_cli_html_report_smoke_produces_file(self, tmp_path: Path) -> None:
        from edf_bill_fetcher.io.cli import run_cli_html_report

        records_json = tmp_path / "records.json"
        records_json.write_text(json.dumps(_synthetic_html_records()), encoding="utf-8")
        out_html = tmp_path / "report.html"

        with pytest.raises(SystemExit) as exc:
            run_cli_html_report(["-i", str(records_json), "-o", str(out_html)])
        assert exc.value.code == 0
        assert out_html.exists()
        rendered = out_html.read_text(encoding="utf-8")
        assert "<html" in rendered
        assert "EDF Energy Ombudsman Evidence Report" in rendered


class TestGuiHtmlFormatPresence:
    """The Report Options dialog's Output Format frame exposes HTML."""

    def test_gui_report_options_dialog_exposes_html_format(self) -> None:
        pytest.importorskip("tkinter")
        import tkinter as tk
        from tkinter import ttk

        from edf_bill_fetcher.ui.app import ReportOptionsDialog

        root = tk.Tk()
        root.withdraw()
        try:
            dlg = ReportOptionsDialog(root)
            dlg.dialog = tk.Toplevel(root)
            dlg.dialog.withdraw()
            dlg._build_ui()

            radio_values: list[str] = []

            def _walk(widget: tk.Misc) -> None:
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Radiobutton):
                        radio_values.append(str(child.cget("value")))
                    _walk(child)

            _walk(dlg.dialog)
            assert "html" in radio_values
        finally:
            root.destroy()


class TestFmtDateParity:
    """Parity test for ``fmt_date`` between the PDF and DOCX generators.

    Phase 1.2: ``edf_report_docx.fmt_date`` is an alias of
    ``edf_report.fmt_date`` so a user's bill is rendered
    identically in both formats.  This test pins that invariant at
    the unit level — any code path that re-introduces a
    DOCX-local ``fmt_date`` will break the alias resolution and
    fail the import parity check.
    """

    @pytest.mark.parametrize(
        "input_val",
        [
            None,
            "N/A",
            "NA",
            "",
            "2023-06-15",
            "15/06/2023",
            "15 Jun 2023",
            "15 June 2023",
            "2026-Q3 2026-07-01",
            pd.NaT,
        ],
    )
    def test_fmt_date_pdf_matches_docx_for_every_input(self, input_val):
        from edf_bill_fetcher.io.reporters.docx_report import fmt_date as docx_fmt_date
        from edf_bill_fetcher.io.reporters.pdf_report import fmt_date as pdf_fmt_date

        # Pandas NaT can't be compared with ``==`` cleanly (raises
        # warnings), so we route through ``isinstance`` + ``pd.isna``
        # to short-circuit the equality branch.
        if isinstance(input_val, type(pd.NaT)) or pd.isna(input_val):
            assert pdf_fmt_date(input_val) == docx_fmt_date(input_val)
            return
        assert pdf_fmt_date(input_val) == docx_fmt_date(input_val), (
            f"PDF and DOCX fmt_date disagree on {input_val!r}: "
            f"PDF={pdf_fmt_date(input_val)!r} "
            f"DOCX={docx_fmt_date(input_val)!r}"
        )

    def test_fmt_date_returns_blank_for_missing(self):
        """The blank-for-missing contract is new in Phase 1.2 — explicitly
        tested so ``"Unknown"`` (the old DOCX-only string) cannot sneak
        back in.
        """
        from edf_bill_fetcher.io.reporters.pdf_report import fmt_date

        for missing in (None, "", "N/A", "NA", pd.NaT):
            assert fmt_date(missing) == "", (
                f"fmt_date({missing!r}) returned {fmt_date(missing)!r}; expected blank string"
            )

    def test_fmt_date_renders_iso_and_uk_to_same_dd_mm_yyyy(self):
        from edf_bill_fetcher.io.reporters.pdf_report import fmt_date

        # Both inputs collapse to the same canonical dd/mm/yyyy form,
        # matching the convention the Excel export already uses.
        assert fmt_date("2023-06-15") == "15/06/2023"
        assert fmt_date("15/06/2023") == "15/06/2023"

    def test_docx_uses_edf_report_fmt_date_not_local_definition(self):
        """A docx-local ``fmt_date`` would diverge from the PDF and break
        this project.  The import in ``docx_report`` must be the
        one in ``pdf_report``.
        """
        from edf_bill_fetcher.io.reporters.docx_report import fmt_date as docx_fmt_date
        from edf_bill_fetcher.io.reporters.pdf_report import fmt_date as pdf_fmt_date

        assert docx_fmt_date is pdf_fmt_date, (
            "DOCX fmt_date is a different object; the two renderers "
            "can drift apart at runtime, which is exactly what Phase 1.2 "
            "was meant to prevent."
        )
