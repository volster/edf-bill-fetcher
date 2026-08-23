"""Tests for the MSG adapter — ``edf_bill_fetcher.io.adapters.msg``.

Pins the same 5-key dict contract as the EML adapter
(``sender`` / ``subject`` / ``date_str`` / ``body_html`` /
``body_text``) and the optional-dependency guard: the module must import
cleanly when ``extract-msg`` is absent, and ``parse_msg_message`` must
raise an informative ``ImportError`` in that environment.
"""

from __future__ import annotations

import importlib
import importlib.util
import sys
from pathlib import Path

import pytest

from edf_bill_fetcher.io.adapters import msg as msg_module

HTML_BODY = "<html><body><h1>Your EDF bill</h1></body></html>"


# ---------------------------------------------------------------------------
# Optional-dependency guard (extract-msg absent)
# ---------------------------------------------------------------------------


def test_has_extract_msg_flag_matches_environment() -> None:
    """HAS_EXTRACT_MSG reflects whether the extract_msg module is importable."""
    expected = importlib.util.find_spec("extract_msg") is not None
    assert msg_module.HAS_EXTRACT_MSG is expected


def test_module_imports_cleanly_without_extract_msg(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Reloading with ``extract_msg`` unimportable sets HAS_EXTRACT_MSG = False.

    Mirrors the HAS_PYPFF branch test: a ``None`` entry under the module
    key makes ``import extract_msg`` raise ImportError on reload.
    """
    monkeypatch.setitem(sys.modules, "extract_msg", None)
    reloaded = importlib.reload(msg_module)
    flag_after_reload = reloaded.HAS_EXTRACT_MSG
    sys.modules.pop("extract_msg", None)
    importlib.reload(msg_module)
    assert flag_after_reload is False


def test_parse_raises_informative_import_error_when_lib_absent(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """parse_msg_message raises ImportError with an install hint when absent."""
    monkeypatch.setitem(sys.modules, "extract_msg", None)
    importlib.reload(msg_module)
    try:
        with pytest.raises(ImportError, match="extract-msg"):
            msg_module.parse_msg_message("bill.msg")
    finally:
        sys.modules.pop("extract_msg", None)
        importlib.reload(msg_module)


# ---------------------------------------------------------------------------
# Parse behaviour (extract-msg installed)
# ---------------------------------------------------------------------------


def test_msg_full_parse(msg_path: Path) -> None:
    """Parse a synthetic .msg into the exact 5-key record dict."""
    result = msg_module.parse_msg_message(msg_path)
    assert result == {
        "sender": "EDF Billing <billing@edfenergy.com>",
        "subject": "Your EDF bill is ready",
        "date_str": "15/01/2024",
        "body_html": HTML_BODY,
        "body_text": "Your EDF bill for January is ready.",
    }


def test_msg_missing_date_and_bodies_return_empty_strings(msg_empty_path: Path) -> None:
    """Missing date and body content yield empty strings, not exceptions."""
    result = msg_module.parse_msg_message(msg_empty_path)
    assert result == {
        "sender": "EDF Billing <billing@edfenergy.com>",
        "subject": "Your EDF bill is ready",
        "date_str": "",
        "body_html": "",
        "body_text": "",
    }
