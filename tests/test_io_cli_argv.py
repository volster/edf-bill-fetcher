"""Tests for edf_bill_fetcher.io.cli — argv-driven CLI entry points.

Covers the headless CLI surface that ``main()`` dispatches to:
``run_cli_extract``, ``run_cli_pdf_report``, ``run_cli_docx_report``,
plus the restricted-pickle ``find_class`` branches and the
``HAS_PYPFF`` import-failure module-level branch.

Every test drives the public CLI functions through ``sys.argv``
or a direct call with a synthetic ``args`` list, stubbing
``EvidenceEngine`` and the report exporters so no real file I/O
or PDF/DOCX rendering happens.  Observable behaviour (stdout/stderr
content, exit codes, files written, config passed to stubs) is
asserted — never call counts.
"""

from __future__ import annotations

import importlib
import json
import pickle
import sys
from pathlib import Path
from typing import Any

import pytest

from edf_bill_fetcher.io import cli as cli_module
from edf_bill_fetcher.io.cli import (
    _RestrictedUnpickler,
    main,
    run_cli_docx_report,
    run_cli_extract,
    run_cli_pdf_report,
)

# ---------------------------------------------------------------------------
# Synthetic records + engine stub
# ---------------------------------------------------------------------------


def _synthetic_records() -> list[dict[str, Any]]:
    """Return deterministic synthetic EDF bill records for stubbing.

    All identifiers are fabricated; no real customer data.
    """
    return [
        {
            "Source": "Local PDF",
            "Sender": "",
            "Date": "15/01/2026",
            "Period From": "01/12/2025",
            "Period To": "31/12/2025",
            "Invoice #": "KI-0000000-0001",
            "Amount (£)": 240.50,
            "Period Charge (£)": 240.50,
            "Unit Rate (p/kWh)": 32.10,
            "% Change": "",
            "Entry Type": "New Bill",
            "Reading": "Smart",
            "Units (kWh)": 750.0,
            "Standing Chg (p/day)": 53.68,
            "Attachment Name": "Jan 2026 bill.pdf",
            "Details": "Your charges: 1 December 2025 - 31 December 2025",
            "Logic Used": "New Invoice Format",
        },
    ]


class _StubEngine:
    """Minimal stand-in for ``EvidenceEngine`` with pre-computed records.

    Exposes every attribute ``run_cli_extract`` reads from the engine
    (``records``, ``filtered_records``, ``error_log``, ``pdf_count``,
    ``email_count``, ``sap_*_rows``) plus the methods it calls
    (``crawl_pst``, ``crawl_local_pdfs``, ``process_htm_file``).
    No real I/O or parsing happens — the stub records what was called
    so tests can assert observable dispatch behaviour.
    """

    def __init__(self, records: list[dict[str, Any]] | None = None) -> None:
        self.records: list[dict[str, Any]] = (
            records if records is not None else _synthetic_records()
        )
        self.filtered_records: list[dict[str, Any]] = []
        self.error_log: list[str] = []
        self.pdf_count: int = 1
        self.email_count: int = 0
        self.sap_contract_rows: list[dict[str, Any]] = []
        self.sap_meter_rows: list[dict[str, Any]] = []
        self.sap_financial_rows: list[dict[str, Any]] = []
        self.crawl_pst_calls: list[Any] = []
        self.crawl_local_pdfs_calls: list[str] = []
        self.process_htm_file_calls: list[str] = []

    def crawl_pst(self, root_folder: Any) -> None:
        """Record the PST root folder handed to the engine."""
        self.crawl_pst_calls.append(root_folder)

    def crawl_local_pdfs(self, path: str) -> None:
        """Record the PDF directory handed to the engine."""
        self.crawl_local_pdfs_calls.append(path)

    def process_htm_file(self, path: str) -> None:
        """Record the HTM file path handed to the engine."""
        self.process_htm_file_calls.append(path)


def _install_stub_engine(monkeypatch: pytest.MonkeyPatch, engine: _StubEngine) -> None:
    """Replace ``EvidenceEngine`` in the cli module's import scope.

    ``run_cli_extract`` imports ``EvidenceEngine`` lazily inside its
    body, so we patch the attribute on the source module
    (``edf_bill_fetcher.collectors.engine``) which is what the
    ``from ... import EvidenceEngine`` line resolves to at call time.
    """
    import edf_bill_fetcher.collectors.engine as engine_mod

    monkeypatch.setattr(engine_mod, "EvidenceEngine", lambda *a, **kw: engine)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


@pytest.fixture
def stub_engine(monkeypatch: pytest.MonkeyPatch) -> _StubEngine:
    """A ``_StubEngine`` installed as the cli's ``EvidenceEngine``."""
    engine = _StubEngine()
    _install_stub_engine(monkeypatch, engine)
    return engine


@pytest.fixture
def stub_export_to_excel(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> dict[str, list[Any]]:
    """Stub ``export_to_excel`` to record calls without writing a workbook.

    ``run_cli_extract`` imports ``export_to_excel`` lazily from
    ``edf_bill_fetcher.io.writers.export``; we patch it on that module.
    """
    import edf_bill_fetcher.io.writers.export as export_mod

    calls: dict[str, list[Any]] = {"calls": []}

    def _fake_export(data, output_path, error_log, config, filtered=None, sap_rows=None):
        """Record the call and touch the output path so callers see a file."""
        calls["calls"].append(
            {
                "data": data,
                "output_path": output_path,
                "error_log": error_log,
                "config": config,
                "filtered": filtered,
                "sap_rows": sap_rows,
            }
        )
        Path(output_path).touch()
        return None

    monkeypatch.setattr(export_mod, "export_to_excel", _fake_export)
    return calls


# ---------------------------------------------------------------------------
# _RestrictedUnpickler.find_class branches (lines 79, 90)
# ---------------------------------------------------------------------------


class TestRestrictedUnpicklerFindClass:
    """Cover the two ``find_class`` rejection branches."""

    def test_find_class_raises_when_module_not_whitelisted(self) -> None:
        """An unknown module name raises ``pickle.UnpicklingError``.

        Covers line 79 — the ``allowed is _SENTINEL`` branch where the
        module is entirely absent from ``_SAFE_CLASSES``.  We build a
        real pickle whose ``__reduce__`` references ``subprocess.check_output``
        (a module not on the whitelist) so the restricted unpickler's
        ``find_class`` rejects it.
        """
        import io

        class _Evil:
            """Object whose reduce references a non-whitelisted callable."""

            def __reduce__(self) -> tuple:
                """Return a reduce tuple pointing at ``subprocess.check_output``."""
                import subprocess

                return (subprocess.check_output, (["echo", "pwned"],))

        payload = pickle.dumps(_Evil(), protocol=pickle.HIGHEST_PROTOCOL)
        with pytest.raises(pickle.UnpicklingError, match="Blocked unsafe class"):
            _RestrictedUnpickler(io.BytesIO(payload)).load()

    def test_find_class_raises_when_resolved_attr_is_not_a_class(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """A whitelisted module+name that resolves to a non-type raises.

        Covers line 90 — the ``not isinstance(cls, type)`` branch inside
        the ``edf_bill_fetcher.collectors.engine`` special-case.  The
        only name in the engine's allowed set is ``EvidenceEngine``, so
        we temporarily replace the real ``EvidenceEngine`` class with a
        non-class callable (a function), build a pickle that references
        it, and confirm the restricted loader rejects it.  The real
        class is restored in the ``finally`` block.
        """
        import io

        import edf_bill_fetcher.collectors.engine as engine_mod

        real_evidence_engine = engine_mod.EvidenceEngine

        def _fake_evidence_engine(*args: Any, **kwargs: Any) -> None:
            """Non-class stand-in injected so ``isinstance(cls, type)`` is False."""
            return None

        _fake_evidence_engine.__module__ = "edf_bill_fetcher.collectors.engine"
        _fake_evidence_engine.__qualname__ = "EvidenceEngine"
        monkeypatch.setattr(engine_mod, "EvidenceEngine", _fake_evidence_engine)
        try:
            payload = pickle.dumps(_fake_evidence_engine, protocol=pickle.HIGHEST_PROTOCOL)
            with pytest.raises(pickle.UnpicklingError, match="is not a class"):
                _RestrictedUnpickler(io.BytesIO(payload)).load()
        finally:
            engine_mod.__dict__["EvidenceEngine"] = real_evidence_engine


# ---------------------------------------------------------------------------
# HAS_PYPFF ImportError branch (lines 19-20)
# ---------------------------------------------------------------------------


class TestHasPypffImportBranch:
    """Cover the ``except ImportError`` branch at module load time."""

    def test_has_pypff_false_branch_executes_when_import_fails(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Reloading cli with ``pypff`` unimportable sets ``HAS_PYPFF = False``.

        Covers lines 19-20 — the ``except ImportError: HAS_PYPFF = False``
        branch.  We hide any cached ``pypff`` module and inject a
        ``None`` entry under that key so the ``import pypff`` statement
        raises ``ImportError`` on reload.
        """
        monkeypatch.setitem(sys.modules, "pypff", None)
        reloaded = importlib.reload(cli_module)
        has_pypff_after_reload = reloaded.HAS_PYPFF
        # Restore the module's real HAS_PYPFF state by reloading
        # it once more under the true import environment.
        sys.modules.pop("pypff", None)
        importlib.reload(cli_module)
        assert has_pypff_after_reload is False


# ---------------------------------------------------------------------------
# run_cli_extract (lines 110-255)
# ---------------------------------------------------------------------------


class TestRunCliExtract:
    """Cover the ``run_cli_extract`` headless extraction entry point."""

    def test_no_source_required_exits_1(self, capsys: pytest.CaptureFixture[str]) -> None:
        """Calling extract with no source flag writes an error and exits 1.

        Covers lines 110-144 — argparse setup, the
        ``not any([parsed.pst, parsed.pdf_dir, parsed.htm])`` guard,
        the stderr write, and ``sys.exit(1)``.
        """
        with pytest.raises(SystemExit) as exc:
            run_cli_extract(["-o", "out.xlsx"])
        assert exc.value.code == 1
        assert "At least one source required" in capsys.readouterr().err

    def test_pdf_dir_happy_path_writes_xlsx_and_prints_summary(
        self,
        stub_engine: _StubEngine,
        stub_export_to_excel: dict[str, list[Any]],
        tmp_path: Path,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """A ``--pdf-dir`` run exports to xlsx and prints the completion summary.

        Covers lines 145-251 — config load/override, the
        ``EvidenceEngine`` instantiation, the ``--pdf-dir`` crawl
        branch, the ``export_to_excel`` call, and the stdout summary
        lines (PDFs processed, Emails matched, Records found, Parse
        errors).
        """
        out_xlsx = tmp_path / "out.xlsx"
        pdf_dir = tmp_path / "pdfs"
        pdf_dir.mkdir()

        run_cli_extract(["--pdf-dir", str(pdf_dir), "-o", str(out_xlsx)])

        # The stub engine was driven.
        assert stub_engine.crawl_local_pdfs_calls == [str(pdf_dir)]
        # export_to_excel was called with the stub records + config.
        assert len(stub_export_to_excel["calls"]) == 1
        call = stub_export_to_excel["calls"][0]
        assert call["output_path"] == str(out_xlsx)
        assert call["data"] == stub_engine.records
        assert call["filtered"] == stub_engine.filtered_records
        assert call["sap_rows"]["contract"] == stub_engine.sap_contract_rows
        # The output file was touched by the stub.
        assert out_xlsx.exists()
        # Stdout carries the summary narrative.
        out = capsys.readouterr().out
        assert "Writing Excel report" in out
        assert "Extraction complete" in out
        assert "PDFs processed: 1" in out
        assert "Records found:  1" in out

    def test_records_json_flag_writes_wrapper_json(
        self,
        stub_engine: _StubEngine,
        stub_export_to_excel: dict[str, list[Any]],
        tmp_path: Path,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """The ``--records-json`` flag writes the wrapper JSON alongside xlsx.

        Covers lines 233-243 — the ``output_data`` dict construction,
        ``json.dump``, and the ``Records saved as JSON`` print.
        """
        out_xlsx = tmp_path / "out.xlsx"
        records_json = tmp_path / "records.json"
        pdf_dir = tmp_path / "pdfs"
        pdf_dir.mkdir()

        run_cli_extract(
            [
                "--pdf-dir",
                str(pdf_dir),
                "-o",
                str(out_xlsx),
                "--records-json",
                str(records_json),
            ]
        )

        assert records_json.exists()
        loaded = json.loads(records_json.read_text(encoding="utf-8"))
        assert loaded["records"] == stub_engine.records
        assert loaded["filtered_records"] == stub_engine.filtered_records
        assert loaded["error_log"] == stub_engine.error_log
        assert "config" in loaded
        assert "extracted_at" in loaded
        out = capsys.readouterr().out
        assert "Records saved as JSON" in out

    def test_config_file_failure_exits_1(
        self,
        stub_engine: _StubEngine,
        tmp_path: Path,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """A config file that fails to load writes an error and exits 1.

        Covers lines 148-154 — the ``try/except`` around config loading.
        """
        bad_config = tmp_path / "bad.json"
        bad_config.write_text("{ not valid json", encoding="utf-8")
        pdf_dir = tmp_path / "pdfs"
        pdf_dir.mkdir()

        with pytest.raises(SystemExit) as exc:
            run_cli_extract(
                ["--pdf-dir", str(pdf_dir), "-o", str(tmp_path / "o.xlsx"), "-c", str(bad_config)]
            )
        assert exc.value.code == 1
        assert "Failed to load config" in capsys.readouterr().err

    def test_no_records_found_exits_1(
        self,
        monkeypatch: pytest.MonkeyPatch,
        stub_export_to_excel: dict[str, list[Any]],
        tmp_path: Path,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """An engine that finds no records writes a warning and exits 1.

        Covers lines 210-212 — the ``if not engine.records`` guard,
        stderr warning, and ``sys.exit(1)``.
        """
        empty_engine = _StubEngine(records=[])
        _install_stub_engine(monkeypatch, empty_engine)
        pdf_dir = tmp_path / "pdfs"
        pdf_dir.mkdir()

        with pytest.raises(SystemExit) as exc:
            run_cli_extract(["--pdf-dir", str(pdf_dir), "-o", str(tmp_path / "o.xlsx")])
        assert exc.value.code == 1
        assert "No billing records found" in capsys.readouterr().err

    def test_error_log_printed_when_nonempty(
        self,
        monkeypatch: pytest.MonkeyPatch,
        stub_export_to_excel: dict[str, list[Any]],
        tmp_path: Path,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """A non-empty ``error_log`` triggers the ``Parse errors`` summary line.

        Covers line 250 — the ``if engine.error_log`` branch.
        """
        engine = _StubEngine()
        engine.error_log = ["PDF: stuck on bill.pdf: no anchor"]
        _install_stub_engine(monkeypatch, engine)
        pdf_dir = tmp_path / "pdfs"
        pdf_dir.mkdir()

        run_cli_extract(["--pdf-dir", str(pdf_dir), "-o", str(tmp_path / "o.xlsx")])
        out = capsys.readouterr().out
        assert "Parse errors:   1" in out

    def test_exception_during_extract_exits_1(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """An exception inside the extract try-block writes an error and exits 1.

        Covers lines 252-255 — the ``except Exception`` handler,
        stderr write, ``traceback.print_exc``, and ``sys.exit(1)``.
        The exception is raised by the stub engine's ``crawl_local_pdfs``
        so it lands inside the ``try`` block (the engine constructor
        runs before the ``try``).
        """

        class _BoomEngine(_StubEngine):
            def crawl_local_pdfs(self, path: str) -> None:
                raise RuntimeError("crawl blew up")

        import edf_bill_fetcher.collectors.engine as engine_mod

        monkeypatch.setattr(engine_mod, "EvidenceEngine", lambda *a, **kw: _BoomEngine())
        pdf_dir = tmp_path / "pdfs"
        pdf_dir.mkdir()

        with pytest.raises(SystemExit) as exc:
            run_cli_extract(["--pdf-dir", str(pdf_dir), "-o", str(tmp_path / "o.xlsx")])
        assert exc.value.code == 1
        err = capsys.readouterr().err
        assert "crawl blew up" in err

    def test_pst_without_pypff_exits_1(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """A ``--pst`` flag when ``HAS_PYPFF`` is False writes an error and exits 1.

        Covers lines 176-180 — the PST dependency check.
        """
        monkeypatch.setattr(cli_module, "HAS_PYPFF", False)
        pst_path = tmp_path / "archive.pst"
        pst_path.touch()

        with pytest.raises(SystemExit) as exc:
            run_cli_extract(["--pst", str(pst_path), "-o", str(tmp_path / "o.xlsx")])
        assert exc.value.code == 1
        assert "libpff-python" in capsys.readouterr().err

    def test_htm_source_drives_process_htm_file(
        self,
        stub_engine: _StubEngine,
        stub_export_to_excel: dict[str, list[Any]],
        tmp_path: Path,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """A ``--htm`` flag drives ``engine.process_htm_file``.

        Covers lines 202-204 — the HTM source branch.
        """
        htm_path = tmp_path / "history.htm"
        htm_path.touch()

        run_cli_extract(["--htm", str(htm_path), "-o", str(tmp_path / "o.xlsx")])
        assert stub_engine.process_htm_file_calls == [str(htm_path)]
        out = capsys.readouterr().out
        assert "Parsing HTM" in out

    def test_config_file_loaded_when_provided(
        self,
        stub_engine: _StubEngine,
        stub_export_to_excel: dict[str, list[Any]],
        tmp_path: Path,
    ) -> None:
        """A config file is loaded and merged into the engine config.

        Covers lines 148-151 — the successful config-load path.
        """
        config_path = tmp_path / "config.json"
        config_path.write_text(json.dumps({"custom_key": "custom_value"}), encoding="utf-8")
        pdf_dir = tmp_path / "pdfs"
        pdf_dir.mkdir()

        run_cli_extract(
            ["--pdf-dir", str(pdf_dir), "-o", str(tmp_path / "o.xlsx"), "-c", str(config_path)]
        )
        call = stub_export_to_excel["calls"][0]
        assert call["config"]["custom_key"] == "custom_value"
        # CLI overrides still apply on top of the loaded config.
        assert call["config"]["use_dedup"] is True
        assert call["config"]["min_amount"] == 500.0

    def test_cli_flags_propagate_into_config(
        self,
        stub_engine: _StubEngine,
        stub_export_to_excel: dict[str, list[Any]],
        tmp_path: Path,
    ) -> None:
        """The ``--no-*`` flags and ``--acc-filter`` flip config booleans.

        Covers lines 156-173 — the ``config.update`` block.
        """
        pdf_dir = tmp_path / "pdfs"
        pdf_dir.mkdir()

        run_cli_extract(
            [
                "--pdf-dir",
                str(pdf_dir),
                "-o",
                str(tmp_path / "o.xlsx"),
                "--acc-filter",
                "A-12345678",
                "--min-amount",
                "100",
                "--no-dedup",
                "--no-anchors",
                "--no-large",
                "--no-reading-class",
                "--no-pdf-fields",
                "--no-filter-below",
                "--domain-filter",
                "example.com",
            ]
        )
        call = stub_export_to_excel["calls"][0]
        cfg = call["config"]
        assert cfg["use_acc_filter"] is True
        assert cfg["acc_num"] == "A-12345678"
        assert cfg["min_amount"] == 100.0
        assert cfg["use_dedup"] is False
        assert cfg["use_anchors"] is False
        assert cfg["use_large"] is False
        assert cfg["use_reading_classification"] is False
        assert cfg["use_pdf_fields"] is False
        assert cfg["filter_below"] is False
        assert cfg["domain_filter"] == "example.com"


# ---------------------------------------------------------------------------
# run_cli_pdf_report (lines 258-323)
# ---------------------------------------------------------------------------


class TestRunCliPdfReport:
    """Cover the ``run_cli_pdf_report`` headless PDF report entry point."""

    def _patch_pdf_generator(
        self, monkeypatch: pytest.MonkeyPatch, returns: tuple[bool, str]
    ) -> dict[str, list[Any]]:
        """Stub ``generate_pdf_from_gui`` to record calls and return ``returns``."""
        import edf_bill_fetcher.io.reporters.pdf_report as pdf_mod

        calls: dict[str, list[Any]] = {"calls": []}

        def _fake_generate(records, output_path, config, engine, filtered=None):
            calls["calls"].append(
                {
                    "records": records,
                    "output_path": output_path,
                    "config": config,
                    "engine": engine,
                    "filtered": filtered,
                }
            )
            return returns

        monkeypatch.setattr(pdf_mod, "generate_pdf_from_gui", _fake_generate)
        return calls

    def test_bare_list_records_success_path(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """A bare-list records JSON succeeds and writes the success message.

        Covers lines 258-316 — argparse, bare-list load (the ``else``
        branch at 290), the success branch (314-316), and ``sys.exit(0)``.
        """
        records = _synthetic_records()
        records_json = tmp_path / "records.json"
        records_json.write_text(json.dumps(records), encoding="utf-8")
        calls = self._patch_pdf_generator(monkeypatch, (True, "Report written"))

        with pytest.raises(SystemExit) as exc:
            run_cli_pdf_report(["-i", str(records_json), "-o", str(tmp_path / "r.pdf")])
        assert exc.value.code == 0
        out = capsys.readouterr().out
        assert "Report written" in out
        assert calls["calls"][0]["records"] == records
        assert calls["calls"][0]["engine"] is None
        assert calls["calls"][0]["filtered"] is None

    def test_wrapper_dict_records_unwrapped(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
    ) -> None:
        """A wrapper ``{"records": [...]}`` JSON is unwrapped into the list.

        Covers line 288 — the ``records = loaded["records"]`` branch.
        """
        records = _synthetic_records()
        wrapper = {"extracted_at": "2026-01-01", "records": records}
        records_json = tmp_path / "records.json"
        records_json.write_text(json.dumps(wrapper), encoding="utf-8")
        calls = self._patch_pdf_generator(monkeypatch, (True, "ok"))

        with pytest.raises(SystemExit):
            run_cli_pdf_report(["-i", str(records_json), "-o", str(tmp_path / "r.pdf")])
        assert calls["calls"][0]["records"] == records

    def test_failure_path_exits_1(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """A ``False`` return from the generator writes an error and exits 1.

        Covers lines 318-319 — the failure branch.
        """
        records = _synthetic_records()
        records_json = tmp_path / "records.json"
        records_json.write_text(json.dumps(records), encoding="utf-8")
        self._patch_pdf_generator(monkeypatch, (False, "rendering failed"))

        with pytest.raises(SystemExit) as exc:
            run_cli_pdf_report(["-i", str(records_json), "-o", str(tmp_path / "r.pdf")])
        assert exc.value.code == 1
        assert "ERROR: rendering failed" in capsys.readouterr().err

    def test_config_file_loaded(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
    ) -> None:
        """A ``--config`` file is loaded and forwarded to the generator.

        Covers lines 293-295 — the config-load branch.
        """
        records = _synthetic_records()
        records_json = tmp_path / "records.json"
        records_json.write_text(json.dumps(records), encoding="utf-8")
        config_path = tmp_path / "config.json"
        config_path.write_text(json.dumps({"report_sections": ["exec_summary"]}), encoding="utf-8")
        calls = self._patch_pdf_generator(monkeypatch, (True, "ok"))

        with pytest.raises(SystemExit):
            run_cli_pdf_report(
                [
                    "-i",
                    str(records_json),
                    "-o",
                    str(tmp_path / "r.pdf"),
                    "-c",
                    str(config_path),
                ]
            )
        assert calls["calls"][0]["config"]["report_sections"] == ["exec_summary"]

    def test_engine_data_pickle_loaded_and_filtered_extracted(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
    ) -> None:
        """A ``--engine-data`` pickle is loaded via the restricted unpickler.

        Covers lines 302-303 — ``_safe_pickle_load`` and the
        ``getattr(engine, "filtered_records", None)`` lookup.
        """
        records = _synthetic_records()
        records_json = tmp_path / "records.json"
        records_json.write_text(json.dumps(records), encoding="utf-8")

        # Build a real pickled engine so _safe_pickle_load succeeds.
        from edf_bill_fetcher.collectors.engine import EvidenceEngine

        def _noop(*a: Any, **kw: Any) -> None:
            return None

        engine = EvidenceEngine({}, _noop, _noop, None)
        engine.filtered_records = [{"filtered": True}]
        pkl_path = tmp_path / "engine.pkl"
        pkl_path.write_bytes(pickle.dumps(engine, protocol=pickle.HIGHEST_PROTOCOL))

        calls = self._patch_pdf_generator(monkeypatch, (True, "ok"))

        with pytest.raises(SystemExit):
            run_cli_pdf_report(
                [
                    "-i",
                    str(records_json),
                    "-o",
                    str(tmp_path / "r.pdf"),
                    "-e",
                    str(pkl_path),
                ]
            )
        call = calls["calls"][0]
        assert isinstance(call["engine"], EvidenceEngine)
        assert call["filtered"] == [{"filtered": True}]

    def test_exception_handler_exits_1(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """An exception inside the report try-block writes an error and exits 1.

        Covers lines 321-323 — the ``except Exception`` handler.
        """
        # A missing records file makes json.load raise FileNotFoundError.
        with pytest.raises(SystemExit) as exc:
            run_cli_pdf_report(
                ["-i", str(tmp_path / "missing.json"), "-o", str(tmp_path / "r.pdf")]
            )
        assert exc.value.code == 1
        assert "ERROR:" in capsys.readouterr().err


# ---------------------------------------------------------------------------
# run_cli_docx_report (lines 326-389)
# ---------------------------------------------------------------------------


class TestRunCliDocxReport:
    """Cover the ``run_cli_docx_report`` headless DOCX report entry point."""

    def _patch_docx_generator(
        self, monkeypatch: pytest.MonkeyPatch, returns: tuple[bool, str]
    ) -> dict[str, list[Any]]:
        """Stub ``generate_docx_from_gui`` to record calls and return ``returns``."""
        import edf_bill_fetcher.io.reporters.docx_report as docx_mod

        calls: dict[str, list[Any]] = {"calls": []}

        def _fake_generate(records, output_path, config, engine, filtered=None):
            calls["calls"].append(
                {
                    "records": records,
                    "output_path": output_path,
                    "config": config,
                    "engine": engine,
                    "filtered": filtered,
                }
            )
            return returns

        monkeypatch.setattr(docx_mod, "generate_docx_from_gui", _fake_generate)
        return calls

    def test_bare_list_success_path(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """A bare-list records JSON succeeds and writes the success message.

        Covers lines 326-383 — argparse, bare-list load, the success
        branch (381-383), and ``sys.exit(0)``.
        """
        records = _synthetic_records()
        records_json = tmp_path / "records.json"
        records_json.write_text(json.dumps(records), encoding="utf-8")
        calls = self._patch_docx_generator(monkeypatch, (True, "DOCX written"))

        with pytest.raises(SystemExit) as exc:
            run_cli_docx_report(["-i", str(records_json), "-o", str(tmp_path / "r.docx")])
        assert exc.value.code == 0
        assert "DOCX written" in capsys.readouterr().out
        assert calls["calls"][0]["records"] == records

    def test_wrapper_dict_unwrapped(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
    ) -> None:
        """A wrapper ``{"records": [...]}`` JSON is unwrapped.

        Covers lines 354-355 — the wrapper unwrap branch.
        """
        records = _synthetic_records()
        wrapper = {"records": records, "meta": "x"}
        records_json = tmp_path / "records.json"
        records_json.write_text(json.dumps(wrapper), encoding="utf-8")
        calls = self._patch_docx_generator(monkeypatch, (True, "ok"))

        with pytest.raises(SystemExit):
            run_cli_docx_report(["-i", str(records_json), "-o", str(tmp_path / "r.docx")])
        assert calls["calls"][0]["records"] == records

    def test_failure_path_exits_1(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """A ``False`` return writes an error and exits 1.

        Covers lines 384-386 — the failure branch.
        """
        records = _synthetic_records()
        records_json = tmp_path / "records.json"
        records_json.write_text(json.dumps(records), encoding="utf-8")
        self._patch_docx_generator(monkeypatch, (False, "docx rendering failed"))

        with pytest.raises(SystemExit) as exc:
            run_cli_docx_report(["-i", str(records_json), "-o", str(tmp_path / "r.docx")])
        assert exc.value.code == 1
        assert "ERROR: docx rendering failed" in capsys.readouterr().err

    def test_config_file_loaded(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
    ) -> None:
        """A ``--config`` file is loaded and forwarded.

        Covers lines 360-362 — the config-load branch.
        """
        records = _synthetic_records()
        records_json = tmp_path / "records.json"
        records_json.write_text(json.dumps(records), encoding="utf-8")
        config_path = tmp_path / "config.json"
        config_path.write_text(json.dumps({"report_sections": ["timeline"]}), encoding="utf-8")
        calls = self._patch_docx_generator(monkeypatch, (True, "ok"))

        with pytest.raises(SystemExit):
            run_cli_docx_report(
                [
                    "-i",
                    str(records_json),
                    "-o",
                    str(tmp_path / "r.docx"),
                    "-c",
                    str(config_path),
                ]
            )
        assert calls["calls"][0]["config"]["report_sections"] == ["timeline"]

    def test_engine_data_pickle_loaded(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
    ) -> None:
        """A ``--engine-data`` pickle is loaded via the restricted unpickler.

        Covers lines 369-370 — ``_safe_pickle_load`` and the
        ``getattr(engine, "filtered_records", None)`` lookup.
        """
        records = _synthetic_records()
        records_json = tmp_path / "records.json"
        records_json.write_text(json.dumps(records), encoding="utf-8")

        from edf_bill_fetcher.collectors.engine import EvidenceEngine

        def _noop(*a: Any, **kw: Any) -> None:
            return None

        engine = EvidenceEngine({}, _noop, _noop, None)
        engine.filtered_records = [{"docx_filtered": True}]
        pkl_path = tmp_path / "engine.pkl"
        pkl_path.write_bytes(pickle.dumps(engine, protocol=pickle.HIGHEST_PROTOCOL))

        calls = self._patch_docx_generator(monkeypatch, (True, "ok"))

        with pytest.raises(SystemExit):
            run_cli_docx_report(
                [
                    "-i",
                    str(records_json),
                    "-o",
                    str(tmp_path / "r.docx"),
                    "-e",
                    str(pkl_path),
                ]
            )
        call = calls["calls"][0]
        assert isinstance(call["engine"], EvidenceEngine)
        assert call["filtered"] == [{"docx_filtered": True}]

    def test_exception_handler_exits_1(
        self,
        tmp_path: Path,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """An exception inside the report try-block writes an error and exits 1.

        Covers lines 387-389 — the ``except Exception`` handler.
        """
        with pytest.raises(SystemExit) as exc:
            run_cli_docx_report(
                ["-i", str(tmp_path / "missing.json"), "-o", str(tmp_path / "r.docx")]
            )
        assert exc.value.code == 1
        assert "ERROR:" in capsys.readouterr().err


# ---------------------------------------------------------------------------
# main() dispatch (lines 392-420)
# ---------------------------------------------------------------------------


class TestMainDispatch:
    """Cover the ``main()`` argv dispatcher and the no-tkinter fallback."""

    def test_pdf_report_dispatch(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
    ) -> None:
        """``main()`` with ``--pdf-report`` dispatches to ``run_cli_pdf_report``.

        Covers lines 394-397 — the ``--pdf-report`` branch.
        """
        records = _synthetic_records()
        records_json = tmp_path / "records.json"
        records_json.write_text(json.dumps(records), encoding="utf-8")

        import edf_bill_fetcher.io.reporters.pdf_report as pdf_mod

        monkeypatch.setattr(
            pdf_mod,
            "generate_pdf_from_gui",
            lambda **kw: (True, "ok"),
        )
        monkeypatch.setattr(
            sys,
            "argv",
            [
                "edf-collector",
                "--pdf-report",
                "-i",
                str(records_json),
                "-o",
                str(tmp_path / "r.pdf"),
            ],
        )

        with pytest.raises(SystemExit):
            main()

    def test_docx_report_dispatch(
        self,
        monkeypatch: pytest.MonkeyPatch,
        tmp_path: Path,
    ) -> None:
        """``main()`` with ``--docx-report`` dispatches to ``run_cli_docx_report``.

        Covers lines 398-400 — the ``--docx-report`` branch.
        """
        records = _synthetic_records()
        records_json = tmp_path / "records.json"
        records_json.write_text(json.dumps(records), encoding="utf-8")

        import edf_bill_fetcher.io.reporters.docx_report as docx_mod

        monkeypatch.setattr(
            docx_mod,
            "generate_docx_from_gui",
            lambda **kw: (True, "ok"),
        )
        monkeypatch.setattr(
            sys,
            "argv",
            [
                "edf-collector",
                "--docx-report",
                "-i",
                str(records_json),
                "-o",
                str(tmp_path / "r.docx"),
            ],
        )

        with pytest.raises(SystemExit):
            main()

    def test_extract_dispatch(
        self,
        monkeypatch: pytest.MonkeyPatch,
        stub_engine: _StubEngine,
        stub_export_to_excel: dict[str, list[Any]],
        tmp_path: Path,
    ) -> None:
        """``main()`` with ``--extract`` dispatches to ``run_cli_extract``.

        Covers lines 401-403 — the ``--extract`` branch.
        """
        pdf_dir = tmp_path / "pdfs"
        pdf_dir.mkdir()
        monkeypatch.setattr(
            sys,
            "argv",
            [
                "edf-collector",
                "--extract",
                "--pdf-dir",
                str(pdf_dir),
                "-o",
                str(tmp_path / "o.xlsx"),
            ],
        )

        main()  # Should complete without raising SystemExit.
        assert len(stub_export_to_excel["calls"]) == 1

    def test_no_tkinter_exits_2(
        self,
        monkeypatch: pytest.MonkeyPatch,
        capsys: pytest.CaptureFixture[str],
    ) -> None:
        """``main()`` with no args and no tkinter writes an error and exits 2.

        Covers lines 405-412 — the ``not HAS_TK`` branch, stderr write,
        and ``sys.exit(2)``.
        """
        monkeypatch.setattr(cli_module, "HAS_TK", False)
        monkeypatch.setattr(sys, "argv", ["edf-collector"])

        with pytest.raises(SystemExit) as exc:
            main()
        assert exc.value.code == 2
        err = capsys.readouterr().err
        assert "tkinter is not available" in err

    def test_no_args_with_tkinter_launches_gui(
        self,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """``main()`` with no args and tkinter present launches the GUI.

        Covers lines 414-420 — the ``import tkinter``, ``tk.Tk()``,
        ``App(root)``, ``root.mainloop()`` path.  We stub all three so
        no real display server is needed.
        """
        monkeypatch.setattr(cli_module, "HAS_TK", True)
        monkeypatch.setattr(sys, "argv", ["edf-collector"])

        mainloop_called: list[bool | str] = []

        class _FakeRoot:
            def mainloop(self) -> None:
                """Record that mainloop was invoked."""
                mainloop_called.append(True)

        class _FakeTkModule:
            @staticmethod
            def Tk() -> _FakeRoot:
                return _FakeRoot()

        # Stub tkinter before main() imports it.
        monkeypatch.setitem(sys.modules, "tkinter", _FakeTkModule)

        # Stub the App import inside main().  main() does
        # ``from edf_bill_fetcher.ui.app import App`` lazily, so we
        # inject a fake ``edf_bill_fetcher.ui.app`` module.
        import types

        fake_app_mod = types.ModuleType("edf_bill_fetcher.ui.app")

        def _fake_app(root: Any) -> None:
            """Record that App was constructed."""
            mainloop_called.append("app_constructed")

        fake_app_mod.App = _fake_app  # type: ignore[attr-defined]
        monkeypatch.setitem(sys.modules, "edf_bill_fetcher.ui.app", fake_app_mod)

        main()
        assert "app_constructed" in mainloop_called
        assert True in mainloop_called
