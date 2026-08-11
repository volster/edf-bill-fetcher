"""Branch/error-path coverage for ``EvidenceEngine`` and its module helpers.

Targets the lines missed by the existing engine test suite (Wave 2 —
``collectors/engine.py`` was at 54% before this file). Coverage focuses on:

* module-level guards: the ``pypff`` ImportError fallback
* the PST attachment filename walker's malformed-record-set tolerance
* sender-email extraction failure paths
* domain-filter full-address matching
* new-invoice / new-credit fallback chains (inv_num / period / charge)
* ``process_text`` amount-parse and large-amount-fallback edge cases
* ``_classify_entry_type`` heuristic branches
* PDF file processing: cancel, empty PDF, page-extraction failure,
  SAP financial dumps, reconciliation statements
* HTM / PST / OST processing error paths
* the recursive PST crawl and the local-PDF folder crawl

All tests are pure unit tests: pypff / pdfplumber are faked, no real
mailboxes or PDFs required.
"""

from __future__ import annotations

import builtins
import importlib
import sys
import threading
import types
from collections.abc import Callable
from datetime import datetime
from pathlib import Path
from typing import Any, cast

import pdfplumber
import pytest

from edf_bill_fetcher.collectors import engine as engine_mod
from edf_bill_fetcher.collectors.engine import EvidenceEngine
from edf_bill_fetcher.models.config import ConfigDict

BASE_CONFIG: ConfigDict = {
    "use_anchors": True,
    "use_large": True,
    "use_reading_classification": True,
    "use_pdf_fields": True,
    "use_acc_filter": False,
    "acc_num": "",
    "min_amount": 500.0,
    "filter_below": True,
    "use_domain_filter": True,
    "domain_filter": "edfenergy.com",
}


def _make_engine(
    config: dict | None = None,
    progress_cb: Callable[[int, int, str], None] | None = None,
    cancel_event: threading.Event | None = None,
) -> EvidenceEngine:
    return EvidenceEngine(
        config=cast(ConfigDict, config or dict(BASE_CONFIG)),
        update_ui_cb=lambda *_a, **_kw: None,
        progress_cb=progress_cb,
        cancel_event=cancel_event,
    )


# ===========================================================================
# Module-level import guard
# ===========================================================================


def test_pypff_import_guard_falls_back_to_false(monkeypatch: pytest.MonkeyPatch) -> None:
    """Reload the module with ``import pypff`` blocked — HAS_PYPFF must be False."""
    original_cls = engine_mod.EvidenceEngine
    real_import = builtins.__import__

    def _block_pypff(name: str, *args: Any, **kwargs: Any) -> Any:
        if name == "pypff":
            raise ImportError("blocked for test")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", _block_pypff)
    sys.modules.pop("pypff", None)
    importlib.reload(engine_mod)
    assert engine_mod.HAS_PYPFF is False

    # Reload again with the real import to restore runtime state, then
    # re-bind the original class: the reload re-executes the module body and
    # would otherwise create a NEW EvidenceEngine object, breaking
    # identity-sensitive consumers (e.g. the pickle round-trip test).
    monkeypatch.setattr(builtins, "__import__", real_import)
    sys.modules.pop("pypff", None)
    importlib.reload(engine_mod)
    engine_mod.__dict__["EvidenceEngine"] = original_cls
    assert isinstance(engine_mod.HAS_PYPFF, bool)


# ===========================================================================
# _pst_attachment_filename — malformed record-set tolerance
# ===========================================================================


class _FakeEntry:
    def __init__(
        self,
        entry_type: object = 0x3707,
        as_string: object = None,
        data: object = None,
        raise_as_string: Exception | None = None,
        raise_data: Exception | None = None,
    ):
        self._entry_type = entry_type
        self._as_string = as_string
        self._data = data
        self._raise_as_string = raise_as_string
        self._raise_data = raise_data

    @property
    def entry_type(self) -> object:
        return self._entry_type

    def get_data_as_string(self) -> object:
        if self._raise_as_string is not None:
            raise self._raise_as_string
        return self._as_string

    def get_data(self) -> object:
        if self._raise_data is not None:
            raise self._raise_data
        return self._data


class _FakeRecordSet:
    def __init__(
        self,
        entries: list | None = None,
        raise_entries: Exception | None = None,
        raise_get_entry: Exception | None = None,
        no_entries_getter: bool = False,
    ):
        self._entries = entries or []
        self._raise_entries = raise_entries
        self._raise_get_entry = raise_get_entry
        self._no_entries_getter = no_entries_getter

    def get_number_of_entries(self):
        if self._raise_entries is not None:
            raise self._raise_entries
        return len(self._entries)

    def get_entry(self, j):
        if self._raise_get_entry is not None:
            raise self._raise_get_entry
        return self._entries[j]


class _FakePstAtt:
    def __init__(
        self,
        count: int = 0,
        raise_count: Exception | None = None,
        raise_get_record_set: Exception | None = None,
        record_sets: list | None = None,
        no_count_getter: bool = False,
    ):
        self._count = count
        self._raise_count = raise_count
        self._raise_get_record_set = raise_get_record_set
        self._record_sets = record_sets or []
        self._no_count_getter = no_count_getter

    def get_number_of_record_sets(self):
        if self._raise_count is not None:
            raise self._raise_count
        return self._count

    def get_record_set(self, i):
        if self._raise_get_record_set is not None:
            raise self._raise_get_record_set
        return self._record_sets[i]


class _NoEntriesRecordSet:
    """Record set lacking ``get_number_of_entries`` entirely."""


def test_pst_att_filename_no_record_sets_getter() -> None:
    assert engine_mod._pst_attachment_filename(object()) is None


def test_pst_att_filename_count_raises() -> None:
    att = _FakePstAtt(count=1, raise_count=RuntimeError("count boom"))
    assert engine_mod._pst_attachment_filename(att) is None


def test_pst_att_filename_get_record_set_raises() -> None:
    att = _FakePstAtt(count=1, raise_get_record_set=RuntimeError("rs boom"))
    assert engine_mod._pst_attachment_filename(att) is None


def test_pst_att_filename_no_entries_getter() -> None:
    att = _FakePstAtt(count=1, record_sets=[_NoEntriesRecordSet()])
    assert engine_mod._pst_attachment_filename(att) is None


def test_pst_att_filename_entries_count_raises() -> None:
    rs = _FakeRecordSet(raise_entries=RuntimeError("entries boom"))
    att = _FakePstAtt(count=1, record_sets=[rs])
    assert engine_mod._pst_attachment_filename(att) is None


def test_pst_att_filename_get_entry_raises() -> None:
    rs = _FakeRecordSet(entries=[object()], raise_get_entry=RuntimeError("entry boom"))
    att = _FakePstAtt(count=1, record_sets=[rs])
    assert engine_mod._pst_attachment_filename(att) is None


def test_pst_att_filename_entry_type_not_intable() -> None:
    entry = _FakeEntry(entry_type="abc")  # int("abc") raises ValueError
    rs = _FakeRecordSet(entries=[entry])
    att = _FakePstAtt(count=1, record_sets=[rs])
    assert engine_mod._pst_attachment_filename(att) is None


def test_pst_att_filename_wrong_entry_type_skipped() -> None:
    entry = _FakeEntry(entry_type=0x1111, as_string="skipped.txt")
    rs = _FakeRecordSet(entries=[entry])
    att = _FakePstAtt(count=1, record_sets=[rs])
    assert engine_mod._pst_attachment_filename(att) is None


def test_pst_att_filename_as_string_raises() -> None:
    entry = _FakeEntry(raise_as_string=RuntimeError("str boom"))
    rs = _FakeRecordSet(entries=[entry])
    att = _FakePstAtt(count=1, record_sets=[rs])
    assert engine_mod._pst_attachment_filename(att) is None


def test_pst_att_filename_happy_str_path() -> None:
    entry = _FakeEntry(as_string="bill.pdf")
    rs = _FakeRecordSet(entries=[entry])
    att = _FakePstAtt(count=1, record_sets=[rs])
    assert engine_mod._pst_attachment_filename(att) == "bill.pdf"


def test_pst_att_filename_raw_bytes_get_data_raises() -> None:
    entry = _FakeEntry(as_string=b"raw", raise_data=RuntimeError("data boom"))
    rs = _FakeRecordSet(entries=[entry])
    att = _FakePstAtt(count=1, record_sets=[rs])
    assert engine_mod._pst_attachment_filename(att) is None


def test_pst_att_filename_raw_bytes_decoded() -> None:
    entry = _FakeEntry(as_string=b"raw", data="bill.pdf".encode("utf-16-le"))
    rs = _FakeRecordSet(entries=[entry])
    att = _FakePstAtt(count=1, record_sets=[rs])
    assert engine_mod._pst_attachment_filename(att) == "bill.pdf"


# ===========================================================================
# _extract_sender_email — failure paths
# ===========================================================================


class _RaisingMsg:
    def __init__(
        self, raise_headers: Exception | None = None, raise_sender: Exception | None = None
    ):
        self._raise_headers = raise_headers
        self._raise_sender = raise_sender

    def get_transport_headers(self):
        if self._raise_headers is not None:
            raise self._raise_headers
        return None

    def get_sender_name(self):
        if self._raise_sender is not None:
            raise self._raise_sender
        return None


def test_extract_sender_email_both_raise() -> None:
    msg = _RaisingMsg(raise_headers=RuntimeError("h"), raise_sender=RuntimeError("s"))
    assert engine_mod._extract_sender_email(msg) == ""


def test_extract_sender_email_headers_raise_sender_works() -> None:
    class _SenderOnly:
        def get_transport_headers(self):
            raise RuntimeError("h")

        def get_sender_name(self):
            return "billing@edfenergy.com"

    assert engine_mod._extract_sender_email(_SenderOnly()) == "billing@edfenergy.com"


# ===========================================================================
# _matches_domain_filter — full-address branches
# ===========================================================================


def test_domain_filter_full_address_match() -> None:
    assert engine_mod._matches_domain_filter("billing@edf.com", "billing@edf.com") is True


def test_domain_filter_full_address_no_match_falls_through() -> None:
    # "@"-containing pattern that does not equal the sender: the loop
    # must fall through instead of raising or returning True.
    assert engine_mod._matches_domain_filter("other@edf.com", "billing@edf.com") is False


def test_domain_filter_domain_match() -> None:
    assert engine_mod._matches_domain_filter("billing@edfenergy.com", "edfenergy.com") is True


# ===========================================================================
# _fallback_amount — credit-total branch
# ===========================================================================


def test_fallback_amount_credit_total() -> None:
    assert engine_mod._fallback_amount("Total credits for this bill £250.00") == (
        250.0,
        "_CREDIT_TOTAL_RE",
    )


# ===========================================================================
# _process_new_invoice — fallback chains
# ===========================================================================


def test_new_invoice_no_amount_returns_false() -> None:
    engine = _make_engine()
    text = "Invoice number: KI-12345678\nAccount number: A-12345678\nYour charges: 01 Jan 2024 - 31 Jan 2024"
    assert engine._process_new_invoice(text, "test", "test.pdf", "2024-01-01") is False
    assert engine.records == []


def test_new_invoice_account_filter_rejects() -> None:
    cfg = dict(BASE_CONFIG, use_acc_filter=True, acc_num="A-99999999")
    engine = _make_engine(cfg)
    text = (
        "Invoice number: KI-12345678\nAccount number: A-12345678\nCurrent balance £1,234.56 debit"
    )
    assert engine._process_new_invoice(text, "test", "test.pdf", "2024-01-01") is False
    assert engine.records == []


def test_new_invoice_fallback_chain_recovers_fields() -> None:
    """Non-KI invoice number + cover-block period + loose amount all fall back."""
    cfg = dict(BASE_CONFIG, filter_below=False)
    engine = _make_engine(cfg)
    text = (
        "Current balance £100.00 debit\n"
        "Invoice number: T1234567-001\n"
        "for the period 01 Jan 2024 - 31 Jan 2024\n"
        "Subtotal £50.00"
    )
    ok = engine._process_new_invoice(text, "test", "test.pdf", "2024-01-01")
    assert ok is True
    assert len(engine.records) == 1
    row = engine.records[0]
    assert row["Invoice #"] == "T1234567-001"
    assert row["Period From"] == "01 Jan 2024"
    assert row["Period To"] == "31 Jan 2024"
    assert row["Period Charge (£)"] == 100.0
    trace = row["_regex_trace"]
    assert "inv_num via _COVER_BLOCK_INV_RE" in trace
    assert "period_from via _COVER_BLOCK_PERIOD_RE" in trace
    assert "period_to via _COVER_BLOCK_PERIOD_RE" in trace
    assert "period_charge via _POUND_AMOUNT_FALLBACK_RE" in trace


def test_new_invoice_fallback_period_charge_via_loose_amount() -> None:
    cfg = dict(BASE_CONFIG, filter_below=False)
    engine = _make_engine(cfg)
    text = (
        "Current balance £100.00 debit\n"
        "Invoice number: KI-12345678\n"
        "Your charges: 01 Jan 2024 - 31 Jan 2024\n"
        "Subtotal £50.00"
    )
    assert engine._process_new_invoice(text, "test", "test.pdf", "2024-01-01") is True
    assert any(
        "period_charge via _POUND_AMOUNT_FALLBACK_RE" in r["_regex_trace"] for r in engine.records
    )


# ===========================================================================
# _process_new_credit — fallback chains
# ===========================================================================


def test_new_credit_no_amount_returns_false() -> None:
    engine = _make_engine()
    text = "Credit note number: KCR-12345678\nAccount number: A-12345678"
    assert engine._process_new_credit(text, "test", "test.pdf", "2024-01-01") is False


def test_new_credit_account_filter_rejects() -> None:
    cfg = dict(BASE_CONFIG, use_acc_filter=True, acc_num="A-99999999")
    engine = _make_engine(cfg)
    text = "Credit note number: KCR-12345678\nTotal credits for this bill £250.00"
    assert engine._process_new_credit(text, "test", "test.pdf", "2024-01-01") is False


def test_new_credit_fallback_inv_num_and_periods() -> None:
    cfg = dict(BASE_CONFIG, filter_below=False)
    engine = _make_engine(cfg)
    text = (
        "Total credits for this bill £250.00\n"
        "Credit note: T1234567-001\n"
        "for the period 01 Jan 2024 - 31 Jan 2024"
    )
    assert engine._process_new_credit(text, "test", "test.pdf", "2024-01-01") is True
    assert len(engine.records) == 1
    row = engine.records[0]
    assert row["Invoice #"] == "T1234567-001"
    assert row["Period From"] == "01 Jan 2024"
    assert row["Period To"] == "31 Jan 2024"
    trace = row["_regex_trace"]
    assert "inv_num via _FALLBACK_INV_RE" in trace
    assert "period_from via _COVER_BLOCK_PERIOD_RE" in trace
    assert "period_to via _COVER_BLOCK_PERIOD_RE" in trace


# ===========================================================================
# process_text — amount-parse and fallback edges
# ===========================================================================


class _BadAmountMatch:
    def group(self, n):
        return "not-a-number"


class _BadAmountPattern:
    def search(self, text):
        return _BadAmountMatch()


def test_process_text_float_parse_error_continues(monkeypatch: pytest.MonkeyPatch) -> None:
    """A pattern whose captured group is not floatable must not abort the loop."""
    monkeypatch.setattr(engine_mod, "AMOUNT_PATTERNS", [("bad", _BadAmountPattern())])
    engine = _make_engine()
    engine.process_text("Invoice total £900.00", "test", "detail", "2024-01-01", sender="s@x.com")
    assert len(engine.records) == 1
    assert engine.records[0]["Amount (£)"] == 900.0
    assert engine.records[0]["Logic Used"] == "Large Amount Fallback"


def test_process_text_large_fallback_all_below_min() -> None:
    """Amounts below min_amount produce no 'highs' and no record."""
    engine = _make_engine()
    engine.process_text(
        "Your bill total is £100.00", "test", "detail", "2024-01-01", sender="s@x.com"
    )
    assert engine.records == []
    assert engine.filtered_records == []


def test_process_text_old_pdf_date_extraction() -> None:
    engine = _make_engine()
    engine.process_text(
        "Bill date: 15 Jan 2024 Your new account balance £600.00",
        "Old PDF",
        "detail",
        "2024-01-01",
    )
    assert len(engine.records) == 1
    assert engine.records[0]["Date"] == "15/01/2024"


def test_process_text_old_pdf_invoice_number() -> None:
    engine = _make_engine()
    engine.process_text(
        "Invoice number: T1234567 Your new account balance £600.00",
        "Old PDF",
        "detail",
        "2024-01-01",
    )
    assert len(engine.records) == 1
    assert engine.records[0]["Invoice #"] == "T1234567"


def test_process_text_old_pdf_period_charge() -> None:
    engine = _make_engine()
    engine.process_text(
        "total charges for this bill £120.00 Your new account balance £600.00",
        "Old PDF",
        "detail",
        "2024-01-01",
    )
    assert len(engine.records) == 1
    assert engine.records[0]["Period Charge (£)"] == 120.0


# ===========================================================================
# _classify_entry_type — heuristic branches
# ===========================================================================


def test_classify_period_and_bill_markers() -> None:
    engine = _make_engine()
    cls = engine._classify_entry_type
    assert (
        cls(
            "Invoice number: T1 total charges",
            None,
            "01/01/2024",
            "31/01/2024",
            "Smart Context",
        )
        == "New Bill"
    )


def test_classify_ongoing_pattern_name() -> None:
    engine = _make_engine()
    assert (
        engine._classify_entry_type(
            "text", "your_new_account_balance", "N/A", "N/A", "Smart Context"
        )
        == "Ongoing Balance"
    )


def test_classify_account_balance_language() -> None:
    engine = _make_engine()
    assert (
        engine._classify_entry_type(
            "account balance £100", "unknown", "N/A", "N/A", "Smart Context"
        )
        == "Ongoing Balance"
    )


def test_classify_period_only() -> None:
    engine = _make_engine()
    assert (
        engine._classify_entry_type(
            "just some text", None, "01/01/2024", "31/01/2024", "Smart Context"
        )
        == "New Bill"
    )


def test_classify_bill_indicators() -> None:
    engine = _make_engine()
    assert (
        engine._classify_entry_type("kWh 100 standing charge", None, "N/A", "N/A", "Smart Context")
        == "New Bill"
    )


def test_classify_default_ongoing_balance() -> None:
    engine = _make_engine()
    assert (
        engine._classify_entry_type("plain text", None, "N/A", "N/A", "Smart Context")
        == "Ongoing Balance"
    )


# ===========================================================================
# process_pdf_file — cancel / empty / page error / SAP / reconciliation
# ===========================================================================


class _FakePage:
    def __init__(self, text: str):
        self._text = text

    def extract_text(self):
        return self._text


class _RaisingPage:
    def extract_text(self):
        raise ValueError("boom")


class _FakePDF:
    def __init__(self, pages):
        self.pages = pages

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


class _FakePdfModule:
    utils = types.SimpleNamespace(
        exceptions=types.SimpleNamespace(
            PdfminerException=pdfplumber.utils.exceptions.PdfminerException
        )
    )

    def __init__(self, pdf: _FakePDF):
        self._pdf = pdf

    def open(self, *args, **kwargs):
        return self._pdf


@pytest.fixture
def dummy_pdf_path(tmp_path: Path) -> str:
    p = tmp_path / "dummy.pdf"
    p.write_bytes(b"not a real pdf")
    return str(p)


def test_process_pdf_file_cancelled(monkeypatch: pytest.MonkeyPatch, dummy_pdf_path: Path) -> None:
    import threading

    engine = _make_engine(cancel_event=threading.Event())
    engine.cancel_event.set()
    monkeypatch.setattr(engine_mod, "pdfplumber", _FakePdfModule(_FakePDF([])))
    engine.process_pdf_file(dummy_pdf_path, "test", "dummy.pdf", "2024-01-01")
    assert engine.error_log == []


def test_process_pdf_file_no_pages(monkeypatch: pytest.MonkeyPatch, dummy_pdf_path: Path) -> None:
    engine = _make_engine()
    monkeypatch.setattr(engine_mod, "pdfplumber", _FakePdfModule(_FakePDF([])))
    engine.process_pdf_file(dummy_pdf_path, "test", "dummy.pdf", "2024-01-01")
    assert engine.error_log
    assert "has no pages" in engine.error_log[0]


def test_process_pdf_file_page_extraction_error(
    monkeypatch: pytest.MonkeyPatch, dummy_pdf_path: Path
) -> None:
    engine = _make_engine()
    monkeypatch.setattr(engine_mod, "pdfplumber", _FakePdfModule(_FakePDF([_RaisingPage()])))
    engine.process_pdf_file(dummy_pdf_path, "test", "dummy.pdf", "2024-01-01")
    assert engine.error_log
    assert "Page extraction failed" in engine.error_log[0]


_SAP_FINANCIAL_TEXT = (
    '"Kraken ID","SAP Account Number","Name"\n'
    '"Main Transactions","Sub Transactions","Transaction Text"\n'
    '"MT1","ST1","a transaction"'
)


def test_process_pdf_file_sap_financial_dump(
    monkeypatch: pytest.MonkeyPatch, dummy_pdf_path: Path
) -> None:
    engine = _make_engine()
    monkeypatch.setattr(
        engine_mod, "pdfplumber", _FakePdfModule(_FakePDF([_FakePage(_SAP_FINANCIAL_TEXT)]))
    )
    engine.process_pdf_file(dummy_pdf_path, "test", "dummy.pdf", "2024-01-01")
    # No slice error — the financial dump must be routed to sap_financial_rows.
    assert engine.error_log == []


_RECON_TEXT = (
    "Bill reference: 12345678 (31/01/2024)\n"
    "Account number: A-12345678\n"
    "Electricity 01 Jan 2024 - 31 Jan 2024 £123.45\n"
    "Your new balance £500.00 debit"
)


def test_process_pdf_file_reconciliation_statement(
    monkeypatch: pytest.MonkeyPatch, dummy_pdf_path: Path
) -> None:
    engine = _make_engine()
    monkeypatch.setattr(
        engine_mod, "pdfplumber", _FakePdfModule(_FakePDF([_FakePage(_RECON_TEXT)]))
    )
    engine.process_pdf_file(dummy_pdf_path, "test", "dummy.pdf", "2024-01-01")
    assert engine.error_log == []
    assert len(engine.filtered_records) == 1  # £123.45 < min_amount


# ===========================================================================
# process_htm_file — UTF-8 fallback
# ===========================================================================


def test_process_htm_file_utf8_fallback(tmp_path: Path) -> None:
    engine = _make_engine()
    p = tmp_path / "bad.htm"
    p.write_bytes(b"<html><body>\xff\xfe</body></html>")
    engine.process_htm_file(str(p))
    assert engine.error_log
    assert "UTF-8 decode error" in engine.error_log[0]


def test_process_htm_file_failure_warns_user() -> None:
    """A failing HTM file must surface a user-facing warning through the
    update_ui callback, not just sit silently in the error log (M-13).
    """
    messages: list[str] = []
    engine = EvidenceEngine(
        config=cast(ConfigDict, dict(BASE_CONFIG)),
        update_ui_cb=messages.append,
    )
    engine.process_htm_file("/nonexistent/path/does-not-exist.htm")
    assert any(m.startswith("Warning: failed to process HTM file") for m in messages)
    assert len(engine.error_log) == 1


def test_process_htm_file_failure_stderr_fallback(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    """If the update_ui callback itself raises, the HTM failure warning
    must still reach the user via stderr (M-13 fallback) — the bare
    except must never swallow the failure entirely.
    """
    monkeypatch.setattr(engine_mod.sys, "stderr", sys.stderr)
    engine = EvidenceEngine(
        config=cast(ConfigDict, dict(BASE_CONFIG)),
        update_ui_cb=lambda _m: (_ for _ in ()).throw(RuntimeError("callback broken")),
    )
    engine.process_htm_file("/nonexistent/path/does-not-exist.htm")
    captured = capsys.readouterr()
    assert "Warning: failed to process HTM file" in captured.err
    assert len(engine.error_log) == 1


# ===========================================================================
# process_pst_file / process_ost_file — pypff guards and fake PST
# ===========================================================================


class _FakeFolder:
    def __init__(self, messages: list | None = None, subfolders: list | None = None):
        self._messages = messages or []
        self._subfolders = subfolders or []

    def get_number_of_sub_messages(self):
        return len(self._messages)

    def get_sub_message(self, i):
        return self._messages[i]

    def get_number_of_sub_folders(self):
        return len(self._subfolders)

    def get_sub_folder(self, j):
        return self._subfolders[j]


class _FakePST:
    def __init__(self, root: _FakeFolder, raise_close: bool = False):
        self._root = root
        self.opened = False
        self._raise_close = raise_close

    def open(self, path):
        self.opened = True

    def get_root_folder(self):
        return self._root

    def close(self):
        if self._raise_close:
            raise RuntimeError("close boom")


class _FakePypffModule:
    def __init__(self, pst: _FakePST):
        self._pst = pst

    def file(self):
        return self._pst


def test_process_pst_file_pypff_missing(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(engine_mod, "HAS_PYPFF", False)
    engine = _make_engine()
    engine.process_pst_file("fake.pst")
    assert engine.error_log
    assert "pypff not installed" in engine.error_log[0]


def test_process_ost_file_delegates(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(engine_mod, "HAS_PYPFF", False)
    engine = _make_engine()
    engine.process_ost_file("fake.ost")
    assert engine.error_log
    assert "pypff not installed" in engine.error_log[0]


def test_process_pst_file_open_crawl_close(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(engine_mod, "HAS_PYPFF", True)
    pst = _FakePST(root=_FakeFolder(), raise_close=True)
    monkeypatch.setattr(engine_mod, "pypff", _FakePypffModule(pst))
    engine = _make_engine()
    engine.process_pst_file("fake.pst")
    assert pst.opened is True
    # close() raising must be swallowed, not propagated
    assert engine.error_log == []


# ===========================================================================
# crawl_pst — the recursive folder walker
# ===========================================================================


class _FakeMsg:
    def __init__(
        self,
        subject: str = "",
        dtime: datetime | None = None,
        html: str | None = None,
        plain: bytes | None = None,
        rtf: bytes | None = None,
        headers: str | None = None,
        sender_name: str = "",
        attachments: list | None = None,
        raise_subject: Exception | None = None,
        raise_headers: Exception | None = None,
        raise_sender: Exception | None = None,
        raise_rtf: Exception | None = None,
        on_get_subject: Any | None = None,
    ) -> None:
        self._subject = subject
        self._dtime = dtime
        self._html = html
        self._plain = plain
        self._rtf = rtf
        self._headers = headers
        self._sender_name = sender_name
        self._attachments = attachments or []
        self._raise_subject = raise_subject
        self._raise_headers = raise_headers
        self._raise_sender = raise_sender
        self._raise_rtf = raise_rtf
        self._on_get_subject = on_get_subject

    def get_subject(self):
        if self._on_get_subject is not None:
            self._on_get_subject()
        if self._raise_subject is not None:
            raise self._raise_subject
        return self._subject

    def get_delivery_time(self):
        return self._dtime

    def get_transport_headers(self):
        if self._raise_headers is not None:
            raise self._raise_headers
        return self._headers

    def get_sender_name(self):
        if self._raise_sender is not None:
            raise self._raise_sender
        return self._sender_name

    def get_html_body(self):
        return self._html

    def get_plain_text_body(self):
        return self._plain

    def get_rtf_body(self):
        if self._raise_rtf is not None:
            raise self._raise_rtf
        return self._rtf

    def get_number_of_attachments(self):
        return len(self._attachments)

    def get_attachment(self, i):
        return self._attachments[i]


class _FakeAtt:
    def __init__(
        self,
        size: int = 0,
        buf: bytes = b"",
        raise_size: Exception | None = None,
        raise_read: Exception | None = None,
        on_get_size: Any | None = None,
    ) -> None:
        self._size = size
        self._buf = buf
        self._raise_size = raise_size
        self._raise_read = raise_read
        self._on_get_size = on_get_size

    def get_size(self):
        if self._on_get_size is not None:
            self._on_get_size()
        if self._raise_size is not None:
            raise self._raise_size
        return self._size

    def read_buffer(self, size):
        if self._raise_read is not None:
            raise self._raise_read
        return self._buf


def test_crawl_pst_pypff_missing(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(engine_mod, "HAS_PYPFF", False)
    engine = _make_engine()
    engine.crawl_pst(_FakeFolder())
    assert engine.error_log
    assert "pypff not installed" in engine.error_log[0]


def test_crawl_pst_cancelled(monkeypatch: pytest.MonkeyPatch) -> None:
    import threading

    monkeypatch.setattr(engine_mod, "HAS_PYPFF", True)
    cancel = threading.Event()
    cancel.set()
    engine = _make_engine(cancel_event=cancel)
    engine.crawl_pst(_FakeFolder())
    assert engine.email_count == 0
    assert engine.error_log == []


def test_crawl_pst_html_body_domain_filter(monkeypatch: pytest.MonkeyPatch) -> None:
    """HTML body + domain-filter match → email processed."""
    monkeypatch.setattr(engine_mod, "HAS_PYPFF", True)
    progress = []
    engine = _make_engine(progress_cb=lambda *a: progress.append(a))
    msg = _FakeMsg(
        subject="Your EDF bill",
        dtime=datetime(2024, 1, 15),
        html="<p>Your new account balance £600.00</p>",
        headers="From: billing@edfenergy.com\nTo: x",
    )
    folder = _FakeFolder(messages=[msg])
    engine.crawl_pst(folder)
    assert engine.email_count == 1
    assert len(engine.records) == 1
    assert engine.records[0]["Amount (£)"] == 600.0
    assert progress  # i % 100 == 0 fires for message 0


def test_crawl_pst_domain_filter_no_match_skips() -> None:
    monkeypatch_true = pytest.MonkeyPatch()
    monkeypatch_true.setattr(engine_mod, "HAS_PYPFF", True)
    try:
        engine = _make_engine()
        msg = _FakeMsg(
            subject="Spam",
            html="<p>Your new account balance £600.00</p>",
            headers="From: spam@x.com\nTo: x",
        )
        engine.crawl_pst(_FakeFolder(messages=[msg]))
        assert engine.email_count == 0
        assert engine.records == []
    finally:
        monkeypatch_true.undo()


def test_crawl_pst_subject_keyword_path() -> None:
    monkeypatch_true = pytest.MonkeyPatch()
    monkeypatch_true.setattr(engine_mod, "HAS_PYPFF", True)
    try:
        cfg = dict(BASE_CONFIG, use_domain_filter=False)
        engine = _make_engine(cfg)
        msg = _FakeMsg(
            subject="EDF STATEMENT JANUARY",
            plain="Total charges for this period £700.00 debit".encode(),
        )
        engine.crawl_pst(_FakeFolder(messages=[msg]))
        assert engine.email_count == 1
        assert len(engine.records) == 1
        assert engine.records[0]["Amount (£)"] == 700.0
    finally:
        monkeypatch_true.undo()


def test_crawl_pst_rtf_body(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(engine_mod, "HAS_PYPFF", True)
    engine = _make_engine()
    rtf = b"{\\rtf1\\ansi " + "Your new account balance £700.00".encode() + b"}"
    msg = _FakeMsg(
        subject="EDF INVOICE",
        headers="From: billing@edfenergy.com\nTo: x",
        rtf=rtf,
    )
    engine.crawl_pst(_FakeFolder(messages=[msg]))
    assert engine.email_count == 1
    assert len(engine.records) == 1
    assert engine.records[0]["Source"] == "Email Body (RTF)"


def test_crawl_pst_rtf_getter_raises_then_no_body(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(engine_mod, "HAS_PYPFF", True)
    engine = _make_engine()
    msg = _FakeMsg(
        subject="EDF",
        headers="From: billing@edfenergy.com\nTo: x",
        raise_rtf=RuntimeError("rtf boom"),
    )
    engine.crawl_pst(_FakeFolder(messages=[msg]))
    assert engine.email_count == 1
    assert engine.records == []
    assert any("No readable body" in e for e in engine.error_log)


def test_crawl_pst_no_body(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(engine_mod, "HAS_PYPFF", True)
    engine = _make_engine()
    msg = _FakeMsg(subject="EDF", headers="From: billing@edfenergy.com\nTo: x")
    engine.crawl_pst(_FakeFolder(messages=[msg]))
    assert engine.email_count == 1
    assert any("No readable body" in e for e in engine.error_log)


def test_crawl_pst_attachments(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """PDF attachment processed; small / non-PDF attachments skipped."""
    monkeypatch.setattr(engine_mod, "HAS_PYPFF", True)
    engine = _make_engine()
    pdf_att = _FakeAtt(size=10, buf=b"%PDF-1.4 fake pdf bytes")
    small_att = _FakeAtt(size=2, buf=b"xx")
    non_pdf_att = _FakeAtt(size=10, buf=b"NOTPDF")
    failing_att = _FakeAtt(raise_size=RuntimeError("size boom"))
    msg = _FakeMsg(
        subject="EDF BILL",
        headers="From: billing@edfenergy.com\nTo: x",
        html="<p>Your new account balance £600.00</p>",
        attachments=[pdf_att, small_att, non_pdf_att, failing_att],
    )
    engine.crawl_pst(_FakeFolder(messages=[msg]))
    assert engine.pdf_count == 1
    assert any('Attachment in "EDF BILL"' in e for e in engine.error_log)


def test_crawl_pst_attachment_loop_cancel(monkeypatch: pytest.MonkeyPatch) -> None:
    import threading

    monkeypatch.setattr(engine_mod, "HAS_PYPFF", True)
    cancel = threading.Event()
    engine = _make_engine(cancel_event=cancel)

    def _set_cancel():
        cancel.set()

    pdf_att = _FakeAtt(size=10, buf=b"%PDF-1.4 fake", on_get_size=_set_cancel)
    msg = _FakeMsg(
        subject="EDF BILL",
        headers="From: billing@edfenergy.com\nTo: x",
        html="<p>Your new account balance £600.00</p>",
        attachments=[pdf_att],
    )
    engine.crawl_pst(_FakeFolder(messages=[msg]))
    # Cancel fires mid-attachment-loop → crawl returns early.
    assert engine.email_count == 1


def test_crawl_pst_message_loop_cancel(monkeypatch: pytest.MonkeyPatch) -> None:
    import threading

    monkeypatch.setattr(engine_mod, "HAS_PYPFF", True)
    cancel = threading.Event()
    engine = _make_engine(cancel_event=cancel)

    def _set_cancel():
        cancel.set()

    msg0 = _FakeMsg(
        subject="EDF BILL",
        headers="From: billing@edfenergy.com\nTo: x",
        html="<p>Your new account balance £600.00</p>",
        on_get_subject=_set_cancel,
    )
    msg1 = _FakeMsg(
        subject="EDF BILL 2",
        headers="From: billing@edfenergy.com\nTo: x",
        html="<p>Your new account balance £700.00</p>",
    )
    engine.crawl_pst(_FakeFolder(messages=[msg0, msg1]))
    assert engine.email_count == 1  # msg1 never processed


def test_crawl_pst_message_exception_logged(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(engine_mod, "HAS_PYPFF", True)
    engine = _make_engine()
    msg = _FakeMsg(subject="EDF", raise_subject=RuntimeError("subject boom"))
    engine.crawl_pst(_FakeFolder(messages=[msg]))
    assert engine.email_count == 0
    assert any("PST message index 0" in e for e in engine.error_log)


def test_crawl_pst_subfolder_recursion(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(engine_mod, "HAS_PYPFF", True)
    engine = _make_engine()
    child = _FakeMsg(
        subject="EDF STATEMENT",
        headers="From: billing@edfenergy.com\nTo: x",
        plain="Your new account balance £700.00".encode(),
    )
    sub = _FakeFolder(messages=[child])
    parent_msg = _FakeMsg(
        subject="EDF BILL",
        headers="From: billing@edfenergy.com\nTo: x",
        html="<p>Your new account balance £600.00</p>",
    )
    engine.crawl_pst(_FakeFolder(messages=[parent_msg], subfolders=[sub]))
    assert engine.email_count == 2
    assert len(engine.records) == 2


def test_crawl_pst_subfolder_cancel(monkeypatch: pytest.MonkeyPatch) -> None:
    import threading

    monkeypatch.setattr(engine_mod, "HAS_PYPFF", True)
    cancel = threading.Event()
    engine = _make_engine(cancel_event=cancel)

    def _set_cancel():
        cancel.set()

    msg = _FakeMsg(
        subject="EDF BILL",
        headers="From: billing@edfenergy.com\nTo: x",
        html="<p>Your new account balance £600.00</p>",
        on_get_subject=_set_cancel,
    )
    sub = _FakeFolder(
        messages=[_FakeMsg(subject="EDF 2", headers="From: billing@edfenergy.com\nTo: x")]
    )
    engine.crawl_pst(_FakeFolder(messages=[msg], subfolders=[sub]))
    assert engine.email_count == 1  # subfolder never crawled


# ===========================================================================
# crawl_local_pdfs
# ===========================================================================


def test_crawl_local_pdfs_missing_path() -> None:
    engine = _make_engine()
    engine.crawl_local_pdfs("/nonexistent/path")
    assert engine.pdf_count == 0


def test_crawl_local_pdfs_mixed_files(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setattr(
        engine_mod, "pdfplumber", _FakePdfModule(_FakePDF([_FakePage("no amount text")]))
    )
    (tmp_path / "readme.txt").write_text("not a pdf")
    (tmp_path / "bill.pdf").write_bytes(b"%PDF-1.4 fake")
    progress = []
    engine = _make_engine(progress_cb=lambda *a: progress.append(a))
    engine.crawl_local_pdfs(str(tmp_path))
    assert engine.pdf_count == 1  # only the .pdf file counted
    assert progress  # update_progress fired per file


def test_crawl_local_pdfs_cancelled(tmp_path: Path) -> None:
    import threading

    cancel = threading.Event()
    cancel.set()
    engine = _make_engine(cancel_event=cancel)
    (tmp_path / "bill.pdf").write_bytes(b"%PDF-1.4 fake")
    engine.crawl_local_pdfs(str(tmp_path))
    assert engine.pdf_count == 0  # _process_one returns before counting
