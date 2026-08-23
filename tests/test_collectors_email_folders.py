"""Tests for EvidenceEngine EML/MSG folder ingestion (wave 6c task 5).

``process_eml_folder`` / ``process_msg_folder`` list a local folder,
filter by ``.eml`` / ``.msg`` extension, and replicate ``crawl_pst``'s
per-message pipeline for each file: sender extraction (from the
adapter's 5-key dict) → domain-filter / subject-keyword gate →
``process_text`` preferring the html body over the plain one.  A
malformed file logs to ``error_log`` and never aborts the folder.

Also covers the CLI ``--eml-dir`` end-to-end run on a synthetic folder
and the GUI "EML/MSG Folder" source selector appearing in Section 1.
"""

from __future__ import annotations

import json
import struct
import sys
import tkinter as tk
from collections.abc import Iterator
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import pytest

from edf_bill_fetcher.collectors.engine import EvidenceEngine
from edf_bill_fetcher.models.config import ConfigDict

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_engine(
    use_domain_filter: bool = True, domain_filter: str = "edfenergy.com"
) -> EvidenceEngine:
    """Build an engine with crawl_pst-compatible config defaults."""
    config: ConfigDict = {
        "use_anchors": True,
        "use_large": True,
        "use_reading_classification": True,
        "use_pdf_fields": True,
        "use_acc_filter": False,
        "acc_num": "",
        "min_amount": 100.0,
        "analysis_min": 100.0,
        "filter_below": True,
        "save_filtered": True,
        "use_dedup": True,
        "save_dups": True,
        "use_domain_filter": use_domain_filter,
        "domain_filter": domain_filter,
    }
    return EvidenceEngine(config, lambda x: None)


def _write_eml(
    folder: Path,
    name: str,
    sender: str,
    subject: str,
    body: str,
    html: str | None = None,
) -> Path:
    """Write a synthetic ``.eml`` message into ``folder`` and return its path."""
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = "customer@example.com"
    msg["Subject"] = subject
    msg["Date"] = "Tue, 15 Jan 2024 10:30:00 +0000"
    if html is not None:
        msg.set_content(body)
        msg.add_alternative(html, subtype="html")
    else:
        msg.set_content(body)
    path = folder / name
    path.write_bytes(msg.as_bytes())
    return path


def _write_msg(folder: Path, name: str, sender_email: str, subject: str, body: str) -> Path:
    """Write a synthetic ``.msg`` (OLE2) message into ``folder``.

    Mirrors the conftest ``msg_path`` fixture, but with a caller-chosen
    body so the folder ingestion has an amount to extract.
    """
    from extract_msg.ole_writer import OleWriter

    header = bytes(8) + struct.pack("<IIII", 0, 0, 0, 0) + bytes(8)
    delta = datetime(2024, 1, 15, 12, 0, 0) - datetime(1601, 1, 1)
    filetime = int(delta.total_seconds() * 10_000_000)
    streams = {
        "__properties_version1.0": header + struct.pack("<IIQ", 0x00390040, 0, filetime),
        "__substg1.0_001A001F": "IPM.Note".encode("utf-16-le"),
        "__substg1.0_0C1A001F": "EDF Billing".encode("utf-16-le"),
        "__substg1.0_5D01001F": sender_email.encode("utf-16-le"),
        "__substg1.0_0037001F": subject.encode("utf-16-le"),
        "__substg1.0_1000001F": body.encode("utf-16-le"),
    }
    writer = OleWriter()
    for stream_name, data in streams.items():
        writer.addEntry(stream_name, data)
    path = folder / name
    writer.write(path)
    return path


def _make_eml_folder(tmp_path: Path) -> Path:
    """A folder of 2 synthetic .eml files: one domain-match, one keyword-gate."""
    folder = tmp_path / "emails"
    folder.mkdir()
    _write_eml(
        folder,
        "a_domain.eml",
        "billing@edfenergy.com",
        "Message from your supplier",  # no EDF/BILL/STATEMENT/ACCOUNT/INVOICE keyword
        "Total charges for this period £120.00 in debit\n",
    )
    _write_eml(
        folder,
        "b_keyword.eml",
        "accounts@somewhere-else.com",  # not on the domain filter
        "Your EDF statement is ready",  # keyword-gate subject
        "Total charges for this period £150.00 in debit\n",
    )
    return folder


# ---------------------------------------------------------------------------
# Engine folder methods
# ---------------------------------------------------------------------------


class TestProcessEmlFolder:
    """``EvidenceEngine.process_eml_folder`` per-message pipeline."""

    def test_domain_match_and_subject_keyword_gates(self, tmp_path: Path) -> None:
        """The domain-filter run keeps only the edfenergy.com file; the
        keyword-gate run keeps only the keyword-subject file."""
        folder = _make_eml_folder(tmp_path)

        domain_engine = _make_engine(use_domain_filter=True)
        domain_engine.process_eml_folder(str(folder))
        assert len(domain_engine.records) == 1
        assert domain_engine.records[0]["Sender"] == "billing@edfenergy.com"
        assert domain_engine.records[0]["Amount (£)"] == 120.0
        assert domain_engine.email_count == 1

        keyword_engine = _make_engine(use_domain_filter=False)
        keyword_engine.process_eml_folder(str(folder))
        assert len(keyword_engine.records) == 1
        assert keyword_engine.records[0]["Sender"] == "accounts@somewhere-else.com"
        assert keyword_engine.records[0]["Amount (£)"] == 150.0

    def test_two_matching_emls_yield_two_records(self, tmp_path: Path) -> None:
        """A folder of 2 domain-matching .eml files yields 2 records."""
        folder = tmp_path / "emails"
        folder.mkdir()
        _write_eml(
            folder,
            "jan.eml",
            "billing@edfenergy.com",
            "Your EDF bill is ready",
            "Total charges for this period £120.00 in debit\n",
        )
        _write_eml(
            folder,
            "feb.eml",
            "billing@edfenergy.com",
            "Your EDF bill is ready",
            "Total charges for this period £150.00 in debit\n",
        )

        engine = _make_engine()
        engine.process_eml_folder(str(folder))

        assert len(engine.records) == 2
        assert engine.email_count == 2
        assert engine.error_log == []

    def test_malformed_eml_logs_error_and_continues(self, tmp_path: Path) -> None:
        """A garbage .eml logs to error_log without aborting the folder."""
        folder = tmp_path / "emails"
        folder.mkdir()
        _write_eml(
            folder,
            "good.eml",
            "billing@edfenergy.com",
            "Your EDF bill is ready",
            "Total charges for this period £120.00 in debit\n",
        )
        (folder / "broken.eml").mkdir()  # unreadable entry → parser raises

        engine = _make_engine()
        engine.process_eml_folder(str(folder))

        assert len(engine.records) == 1, "the good file must still be processed"
        assert len(engine.error_log) == 1
        assert "broken.eml" in engine.error_log[0]

    def test_prefers_html_body_over_plain(self, tmp_path: Path) -> None:
        """The html body wins when both html and plain parts are present."""
        folder = tmp_path / "emails"
        folder.mkdir()
        _write_eml(
            folder,
            "multipart.eml",
            "billing@edfenergy.com",
            "Your EDF bill is ready",
            "Total charges for this period £50.00 in debit\n",
            html="<html><body><p>Total charges for this period £120.00 in debit</p></body></html>",
        )

        engine = _make_engine()
        engine.process_eml_folder(str(folder))

        assert len(engine.records) == 1
        assert engine.records[0]["Amount (£)"] == 120.0

    def test_missing_folder_returns_without_error(self, tmp_path: Path) -> None:
        """A nonexistent folder is a silent no-op (mirrors crawl_local_pdfs)."""
        engine = _make_engine()
        engine.process_eml_folder(str(tmp_path / "does-not-exist"))
        assert engine.records == []
        assert engine.error_log == []


class TestProcessMsgFolder:
    """``EvidenceEngine.process_msg_folder`` per-message pipeline."""

    def test_extracts_records_from_msg_folder(self, tmp_path: Path) -> None:
        """A folder holding one .msg with an amount in the body yields a record."""
        pytest.importorskip("extract_msg")
        folder = tmp_path / "msgs"
        folder.mkdir()
        _write_msg(
            folder,
            "bill.msg",
            "billing@edfenergy.com",
            "Your EDF bill is ready",
            "Total charges for this period £120.00 in debit",
        )

        engine = _make_engine()
        engine.process_msg_folder(str(folder))

        assert len(engine.records) == 1
        assert engine.records[0]["Amount (£)"] == 120.0
        assert engine.email_count == 1

    def test_missing_extract_msg_logs_error_once(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """Without extract-msg the folder is skipped with a single log entry."""
        import edf_bill_fetcher.collectors.engine as engine_mod

        monkeypatch.setattr(engine_mod, "HAS_EXTRACT_MSG", False)
        folder = tmp_path / "msgs"
        folder.mkdir()
        (folder / "bill.msg").write_bytes(b"junk")

        engine = _make_engine()
        engine.process_msg_folder(str(folder))

        assert engine.records == []
        assert len(engine.error_log) == 1
        assert "extract-msg" in engine.error_log[0]


# ---------------------------------------------------------------------------
# CLI end-to-end
# ---------------------------------------------------------------------------


class TestCliEmlDirEndToEnd:
    """``run_cli_extract --eml-dir`` drives the real engine over a folder."""

    def test_cli_eml_dir_runs_end_to_end(self, tmp_path: Path) -> None:
        """Two synthetic .eml files produce an xlsx and 2 records in JSON."""
        from edf_bill_fetcher.io.cli import run_cli_extract

        folder = tmp_path / "emails"
        folder.mkdir()
        _write_eml(
            folder,
            "jan.eml",
            "billing@edfenergy.com",
            "Your EDF bill is ready",
            "Total charges for this period £120.00 in debit\n",
        )
        _write_eml(
            folder,
            "feb.eml",
            "billing@edfenergy.com",
            "Your EDF bill is ready",
            "Total charges for this period £150.00 in debit\n",
        )
        out_xlsx = tmp_path / "out.xlsx"
        records_json = tmp_path / "records.json"

        run_cli_extract(
            ["--eml-dir", str(folder), "-o", str(out_xlsx), "--records-json", str(records_json)]
        )

        assert out_xlsx.exists()
        data = json.loads(records_json.read_text(encoding="utf-8"))
        assert len(data["records"]) == 2


# ---------------------------------------------------------------------------
# GUI source selector
# ---------------------------------------------------------------------------

pytestmark_gui = pytest.mark.skipif(
    sys.platform == "win32",
    reason=(
        "Windows CI intermittently fails with _tkinter.TclError: "
        "invalid command name 'tcl_findLibrary'"
    ),
)


def _walk_children(widget: tk.Misc) -> Iterator[tk.Misc]:
    for child in widget.winfo_children():
        yield child
        yield from _walk_children(child)


def _all_widget_text(widget: tk.Misc) -> list[str]:
    texts: list[str] = []
    for child in _walk_children(widget):
        try:
            cls = child.winfo_class()
            if cls in ("Label", "Button", "Checkbutton", "TLabel", "TButton"):
                t = child.cget("text")
                if t and isinstance(t, str):
                    texts.append(t)
        except tk.TclError:
            pass
    return texts


@pytestmark_gui
def test_gui_eml_msg_folder_source_appears(monkeypatch: pytest.MonkeyPatch) -> None:
    """Section 1 exposes an "EML/MSG Folder" selector backed by a StringVar."""
    from edf_bill_fetcher.ui.app import App

    monkeypatch.setattr(App, "_load_config", lambda self: None)
    root = tk.Tk()
    root.withdraw()
    try:
        app = App(root)
        texts = _all_widget_text(app.root)
        assert any("EML/MSG Folder:" in t for t in texts)
        assert hasattr(app, "eml_msg_dir"), "App must carry an eml_msg_dir StringVar"
    finally:
        root.destroy()
