"""Coverage tests for the io.reporters eager-re-export shim layer.

Closes the missed-line gap (10 stmts, 10 missed across 3 files) by
exercising all re-export paths from:

  * edf_bill_fetcher/io/reporters/__init__.py
  * edf_bill_fetcher/io/reporters/pdf_report.py
  * edf_bill_fetcher/io/reporters/docx_report.py

Each shim is a plain ``from X import (Y, Z, ...)`` eager re-export
(NOT PEP 562). The shims exist to give callers a package-namespace
import path while the canonical homes (``edf_report`` /
``edf_report_docx``) remain at the top level until the next major
strip-the-compat-layer refactor window. These tests assert
twin-identity: each re-exported name IS the same object as the
canonical source symbol.
"""

from __future__ import annotations

import edf_report
import edf_report_docx
from edf_bill_fetcher.io import reporters
from edf_bill_fetcher.io.reporters import docx_report as _docx_shim
from edf_bill_fetcher.io.reporters import pdf_report as _pdf_shim

# ---------- __init__.py package-level re-exports ----------

def test_reporters_init_re_exports_pdf_report_names_are_twin_identical() -> None:
    """Each name re-exported via reporters.__init__ is the same object as edf_report."""
    assert reporters.REPORT_SECTIONS is edf_report.REPORT_SECTIONS
    assert reporters.RenderContext is edf_report.RenderContext
    assert reporters.fmt_date is edf_report.fmt_date
    assert reporters.fmt_money is edf_report.fmt_money
    assert reporters.fmt_number is edf_report.fmt_number
    assert reporters.fmt_pct is edf_report.fmt_pct
    assert reporters.generate_ombudsman_pdf is edf_report.generate_ombudsman_pdf
    assert reporters.generate_pdf_from_gui is edf_report.generate_pdf_from_gui


def test_reporters_init_re_exports_docx_report_names_are_twin_identical() -> None:
    """Each name re-exported via reporters.__init__ is the same object as edf_report_docx."""
    assert reporters.generate_docx_from_gui is edf_report_docx.generate_docx_from_gui
    assert reporters.generate_ombudsman_docx is edf_report_docx.generate_ombudsman_docx


def test_reporters_all_lists_canonical_names() -> None:
    """__all__ matches the expected re-export surface."""
    expected = {
        "REPORT_SECTIONS",
        "RenderContext",
        "fmt_date",
        "fmt_money",
        "fmt_number",
        "fmt_pct",
        "generate_docx_from_gui",
        "generate_ombudsman_docx",
        "generate_ombudsman_pdf",
        "generate_pdf_from_gui",
    }
    assert set(reporters.__all__) == expected


# ---------- pdf_report.py submodule re-exports ----------

def test_pdf_report_shim_re_exports_are_twin_identical() -> None:
    """io.reporters.pdf_report re-exports are the same objects as canonical edf_report."""
    assert _pdf_shim.REPORT_SECTIONS is edf_report.REPORT_SECTIONS
    assert _pdf_shim.RenderContext is edf_report.RenderContext
    assert _pdf_shim.fmt_date is edf_report.fmt_date
    assert _pdf_shim.fmt_money is edf_report.fmt_money
    assert _pdf_shim.fmt_number is edf_report.fmt_number
    assert _pdf_shim.fmt_pct is edf_report.fmt_pct
    assert _pdf_shim.generate_ombudsman_pdf is edf_report.generate_ombudsman_pdf
    assert _pdf_shim.generate_pdf_from_gui is edf_report.generate_pdf_from_gui


def test_pdf_report_shim_all_lists_canonical_names() -> None:
    """pdf_report shim __all__ matches the expected re-export surface."""
    expected = {
        "REPORT_SECTIONS",
        "RenderContext",
        "fmt_date",
        "fmt_money",
        "fmt_number",
        "fmt_pct",
        "generate_ombudsman_pdf",
        "generate_pdf_from_gui",
    }
    assert set(_pdf_shim.__all__) == expected


# ---------- docx_report.py submodule re-exports ----------

def test_docx_report_shim_re_exports_are_twin_identical() -> None:
    """io.reporters.docx_report re-exports are the same objects as canonical edf_report_docx."""
    assert _docx_shim.fmt_money is edf_report_docx.fmt_money
    assert _docx_shim.fmt_number is edf_report_docx.fmt_number
    assert _docx_shim.generate_docx_from_gui is edf_report_docx.generate_docx_from_gui
    assert _docx_shim.generate_ombudsman_docx is edf_report_docx.generate_ombudsman_docx


def test_docx_report_shim_all_lists_canonical_names() -> None:
    """docx_report shim __all__ matches the expected re-export surface."""
    expected = {
        "fmt_money",
        "fmt_number",
        "generate_docx_from_gui",
        "generate_ombudsman_docx",
    }
    assert set(_docx_shim.__all__) == expected
