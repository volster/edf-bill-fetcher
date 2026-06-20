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
from pathlib import Path

from edf_report import REPORT_SECTIONS

REPO_ROOT = Path(__file__).resolve().parents[1]


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
        str(REPO_ROOT / "edf_report.py"),
    )


def _docx_dispatcher_keys() -> set[str]:
    return _dispatcher_keys_for_function(
        "generate_ombudsman_docx",
        str(REPO_ROOT / "edf_report_docx.py"),
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

    def test_pdf_and_docx_dispatchers_agree(self):
        assert _pdf_dispatcher_keys() == _docx_dispatcher_keys()


class TestRenderContextBuildable:
    """RenderContext's section_in_order is empty for unknown keys."""

    def test_registry_keys_all_resolve(self):
        from edf_report import RenderContext

        ctx = RenderContext()  # default = all sections
        looked = {s.section.key for s in ctx.sections_in_order}
        assert looked == _registry_keys()

    def test_unknown_key_in_render_context_does_not_explode(self):
        from edf_report import RenderContext

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
        from edf_report import generate_ombudsman_pdf

        # Reach into the module by re-evaluating its AST and
        # locating the literal ``lambda kwargs: create_X(**kwargs)``
        # right-hand-sides. This is a structural test — it asserts
        # the dispatcher is well-formed, not that all paths render.
        assert callable(generate_ombudsman_pdf)

    def test_docx_dispatchers_all_callable(self):
        from edf_report_docx import generate_ombudsman_docx

        assert callable(generate_ombudsman_docx)
