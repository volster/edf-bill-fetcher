"""Wave 3 coverage: dialog handlers, thread-marshalling, _run, config branches.

Exercises ``edf_bill_fetcher/ui/app.py`` through observable UI and artifact behavior:
modal handlers via boundary mocks (filedialog/messagebox), the EXTRACT
workflow state machine, thread-marshalling branches, the full ``_run``
worker, ``load_spreadsheet_and_report``, ``_resolve_output_path``,
``_run_auto_report``, ``ReportOptionsDialog``, and ``_load_config`` edges.

Boundary-mock strategy: patch ONLY the filedialog/messagebox call, then
assert UI state, preserved evidence and truthful run manifests. Collector and
renderer mocks expose external calls without bypassing the coordinator.
"""

from __future__ import annotations

import json
import os
import tkinter as tk
from collections.abc import Callable, Generator
from datetime import date
from pathlib import Path
from threading import Event
from types import SimpleNamespace
from typing import Any, cast
from unittest.mock import MagicMock, patch

import pandas as pd
import pytest

from edf_bill_fetcher.ui import app as app_module
from edf_bill_fetcher.ui.app import App, ReportOptionsDialog


class _NoThread:
    """Mock threading.Thread that doesn't actually start (for start_thread tests)."""

    def __init__(self, *args: object, **kwargs: object) -> None:
        pass

    def start(self) -> None:
        pass

    daemon = False


def _make_app(root: tk.Tk) -> App:
    """Construct App without invoking the real _load_config (avoids dev config leak).

    Suppresses the constructor's ``_load_config`` call so a developer's real
    ``~/.edf_collector/config.json`` cannot poison the App instance, then
    restores the real method so explicit ``_load_config()`` calls in tests run
    against the temp config path the test sets afterwards.

    Uses setattr/delattr directly instead of ``monkeypatch`` to avoid
    clobbering other patches set by autouse fixtures on the same test
    (e.g. the messagebox suppression in ``TestRunWorker``).
    """
    original_load = App._load_config
    App._load_config = lambda self: None  # type: ignore[method-assign]
    try:
        return App(root)
    finally:
        App._load_config = original_load  # type: ignore[method-assign]


@pytest.fixture
def root() -> Generator[tk.Tk, None, None]:
    r = tk.Tk()
    r.withdraw()
    yield r
    # Drain any pending after-callbacks before destroying to prevent them
    # from firing during a later test's root instance (Tkinter event queues
    # are per-interpreter, not per-root, so stale callbacks can leak).
    try:
        r.update_idletasks()
    except tk.TclError:
        pass
    r.destroy()


@pytest.fixture
def app(root: tk.Tk, tmp_path: Path) -> App:
    instance = _make_app(root)
    instance.output_folder.set(str(tmp_path))
    instance._CONFIG_PATH = str(tmp_path / "settings.json")
    return instance


# ---------------------------------------------------------------------------
# ReportOptionsDialog
# ---------------------------------------------------------------------------


class TestReportOptionsDialog:
    """Cover show/_build_ui/_select_all/_none/_defaults/_generate/_cancel."""

    def test_show_returns_none_on_cancel(self, root: tk.Tk) -> None:
        """_cancel sets result None and destroys dialog; show returns None."""
        dlg = ReportOptionsDialog(root)
        # Build the dialog UI manually without wait_window (which blocks).
        dlg.dialog = tk.Toplevel(root)
        dlg.dialog.withdraw()
        dlg._build_ui()
        dlg._cancel()
        assert dlg.result is None
        assert dlg.dialog is None or not dlg.dialog.winfo_exists()

    def test_generate_with_selections_returns_dict(self, root: tk.Tk) -> None:
        dlg = ReportOptionsDialog(root)
        dlg.dialog = tk.Toplevel(root)
        dlg.dialog.withdraw()
        dlg._build_ui()
        # Defaults have all sections True, so _generate should produce a result.
        dlg._generate()
        assert dlg.result is not None
        assert "format" in dlg.result
        assert "sections" in dlg.result
        assert len(dlg.result["sections"]) > 0

    def test_generate_zero_selected_shows_warning(self, root: tk.Tk) -> None:
        dlg = ReportOptionsDialog(root)
        dlg.dialog = tk.Toplevel(root)
        dlg.dialog.withdraw()
        dlg._build_ui()
        dlg._select_none()
        with patch("tkinter.messagebox.showwarning") as mock_warn:
            dlg._generate()
            mock_warn.assert_called_once()
        assert dlg.result is None

    def test_select_all_then_none_toggles_vars(self, root: tk.Tk) -> None:
        dlg = ReportOptionsDialog(root)
        dlg.dialog = tk.Toplevel(root)
        dlg.dialog.withdraw()
        dlg._build_ui()
        dlg._select_none()
        assert all(not v.get() for v in dlg.section_vars.values())
        dlg._select_all()
        assert all(v.get() for v in dlg.section_vars.values())

    def test_select_defaults_restores_defaults(self, root: tk.Tk) -> None:
        dlg = ReportOptionsDialog(root)
        dlg.dialog = tk.Toplevel(root)
        dlg.dialog.withdraw()
        dlg._build_ui()
        dlg._select_none()
        dlg._select_defaults()
        for key, _, default in ReportOptionsDialog.SECTIONS:
            assert dlg.section_vars[key].get() == default

    def test_format_var_defaults_to_both(self, root: tk.Tk) -> None:
        dlg = ReportOptionsDialog(root)
        dlg.dialog = tk.Toplevel(root)
        dlg.dialog.withdraw()
        dlg._build_ui()
        assert dlg.format_var.get() == "both"

    def test_show_full_flow_via_after_callback(self, root: tk.Tk) -> None:
        """Drive show() end-to-end: schedule _generate via root.after, then update.

        show() calls wait_window which blocks until the Toplevel is destroyed.
        We schedule _generate (which destroys the dialog) via root.after so the
        wait_window returns once the event loop processes it.
        """
        dlg = ReportOptionsDialog(root)
        root.after(50, dlg._generate)
        result = dlg.show()
        assert result is not None
        assert "sections" in result

    def test_mousewheel_and_canvas_configure_callbacks(self, root: tk.Tk) -> None:
        """Exercise the <MouseWheel> and <Configure> bindings via two show() cycles."""
        dlg = ReportOptionsDialog(root)
        root.after(50, dlg._generate)
        dlg.show()
        dlg2 = ReportOptionsDialog(root)
        root.after(50, dlg2._cancel)
        dlg2.show()


# ---------------------------------------------------------------------------
# File-picker handlers (boundary-mock filedialog, assert state mutation)
# ---------------------------------------------------------------------------


class TestPickHandlers:
    def test_pick_pst_sets_var(self, app: App) -> None:
        with patch("tkinter.filedialog.askopenfilename", return_value="/tmp/test.pst"):
            app._pick_pst()
        assert app.pst_path.get() == "/tmp/test.pst"

    def test_pick_pst_empty_no_change(self, app: App) -> None:
        original = app.pst_path.get()
        with patch("tkinter.filedialog.askopenfilename", return_value=""):
            app._pick_pst()
        assert app.pst_path.get() == original

    def test_pick_pdf_dir_sets_var(self, app: App) -> None:
        with patch("tkinter.filedialog.askdirectory", return_value="/tmp/pdfs"):
            app._pick_pdf_dir()
        assert app.pdf_dir.get() == "/tmp/pdfs"

    def test_pick_pdf_dir_empty_no_change(self, app: App) -> None:
        original = app.pdf_dir.get()
        with patch("tkinter.filedialog.askdirectory", return_value=""):
            app._pick_pdf_dir()
        assert app.pdf_dir.get() == original

    def test_pick_output_folder_sets_var_and_saves(
        self, app: App, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        saved: list[int] = []
        monkeypatch.setattr(app, "_save_config", lambda: saved.append(1))
        with patch("tkinter.filedialog.askdirectory", return_value="/tmp/out"):
            app._pick_output_folder()
        assert app.output_folder.get() == "/tmp/out"
        assert saved == [1]

    def test_pick_output_folder_empty_no_save(
        self, app: App, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        saved: list[int] = []
        monkeypatch.setattr(app, "_save_config", lambda: saved.append(1))
        with patch("tkinter.filedialog.askdirectory", return_value=""):
            app._pick_output_folder()
        assert saved == []


# ---------------------------------------------------------------------------
# _open_report_options + build_ui amalgamate callback
# ---------------------------------------------------------------------------


class TestOpenReportOptions:
    def test_open_report_options_persists_on_ok(
        self, app: App, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        fake_dialog = MagicMock()
        fake_dialog.show.return_value = {"format": "pdf", "sections": ["exec_summary"]}
        saved: list[int] = []
        monkeypatch.setattr(app, "_save_config", lambda: saved.append(1))
        with patch("edf_bill_fetcher.ui.app.ReportOptionsDialog", return_value=fake_dialog):
            app._open_report_options()
        assert app._report_options == {"format": "pdf", "sections": ["exec_summary"]}
        assert saved == [1]

    def test_open_report_options_no_save_on_cancel(
        self, app: App, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        fake_dialog = MagicMock()
        fake_dialog.show.return_value = None
        original = app._report_options.copy()
        saved: list[int] = []
        monkeypatch.setattr(app, "_save_config", lambda: saved.append(1))
        with patch("edf_bill_fetcher.ui.app.ReportOptionsDialog", return_value=fake_dialog):
            app._open_report_options()
        assert app._report_options == original
        assert saved == []


class TestBuildUiCallbacks:
    def test_amalgamate_state_callback_via_save_dups_toggle(self, app: App, root: tk.Tk) -> None:
        app.use_dedup.set(True)
        app.save_dups.set(True)
        root.update_idletasks()
        app.save_dups.set(False)
        root.update_idletasks()


# ---------------------------------------------------------------------------
# Thread-marshalling: set_status / set_progress / _show
# ---------------------------------------------------------------------------


class TestThreadMarshalling:
    def test_set_status_main_thread_inline(self, app: App, root: tk.Tk) -> None:
        app.set_status("hello-main")
        root.update_idletasks()
        assert app.status.get() == "hello-main"

    def test_set_status_off_thread_via_after(self, app: App, root: tk.Tk) -> None:
        """Off-thread branch: simulate non-main thread so root.after(0,...) is used.

        Tkinter is not thread-safe, so we do NOT spawn a real thread that calls
        root.after. Instead we patch threading.current_thread() to return a
        non-main thread, which makes set_status take the ``else`` branch and
        schedule via root.after(0, _apply). We then pump the queue on the main
        thread (safe) and assert the state mutation.
        """
        fake_thread = MagicMock()
        with patch("edf_bill_fetcher.ui.app.threading.current_thread", return_value=fake_thread):
            with patch("edf_bill_fetcher.ui.app.threading.main_thread", return_value=object()):
                app.set_status("from-worker")
                app._poll_events()
        root.update()
        assert app.status.get() == "from-worker"

    def test_set_progress_main_thread_clamps_and_sets(self, app: App, root: tk.Tk) -> None:
        app.set_progress(50, 100, text="halfway")
        root.update_idletasks()
        assert app.progress_v.get() == 50.0
        assert app.status.get() == "halfway"

    def test_set_progress_clamps_above_100(self, app: App, root: tk.Tk) -> None:
        app.set_progress(150, 100)
        root.update_idletasks()
        assert app.progress_v.get() == 100.0

    def test_set_progress_clamps_below_0(self, app: App, root: tk.Tk) -> None:
        app.set_progress(-10, 100)
        root.update_idletasks()
        assert app.progress_v.get() == 0.0

    def test_set_progress_total_zero_is_zero(self, app: App, root: tk.Tk) -> None:
        app.set_progress(5, 0)
        root.update_idletasks()
        assert app.progress_v.get() == 0.0

    def test_set_progress_off_thread_via_after(self, app: App, root: tk.Tk) -> None:
        """Off-thread branch for set_progress (see set_status off-thread test)."""
        fake_thread = MagicMock()
        with patch("edf_bill_fetcher.ui.app.threading.current_thread", return_value=fake_thread):
            with patch("edf_bill_fetcher.ui.app.threading.main_thread", return_value=object()):
                app.set_progress(25, 100, text="worker-progress")
                app._poll_events()
        root.update()
        assert app.progress_v.get() == 25.0
        assert app.status.get() == "worker-progress"

    def test_show_info_main_thread(self, app: App) -> None:
        with patch("tkinter.messagebox.showinfo") as mock_info:
            app._show("info", "T", "msg")
            mock_info.assert_called_once_with("T", "msg")

    def test_show_warning_main_thread(self, app: App) -> None:
        with patch("tkinter.messagebox.showwarning") as mock_warn:
            app._show("warning", "T", "msg")
            mock_warn.assert_called_once_with("T", "msg")

    def test_show_error_main_thread(self, app: App) -> None:
        with patch("tkinter.messagebox.showerror") as mock_err:
            app._show("error", "T", "msg")
            mock_err.assert_called_once_with("T", "msg")

    def test_show_unknown_level_defaults_error(self, app: App) -> None:
        with patch("tkinter.messagebox.showerror") as mock_err:
            app._show("other", "T", "msg")
            mock_err.assert_called_once_with("T", "msg")

    def test_show_off_thread_via_after(self, app: App, root: tk.Tk) -> None:
        """Off-thread branch for _show (see set_status off-thread test)."""
        fake_thread = MagicMock()
        with patch("edf_bill_fetcher.ui.app.threading.current_thread", return_value=fake_thread):
            with patch("edf_bill_fetcher.ui.app.threading.main_thread", return_value=object()):
                with patch("tkinter.messagebox.showinfo") as mock_info:
                    app._show("info", "T", "worker-msg")
                    app._poll_events()
                    # Process the queued after(0, _s) callback.
                    root.update_idletasks()
                    root.update()
        mock_info.assert_called_once_with("T", "worker-msg")


# ---------------------------------------------------------------------------
# EXTRACT workflow state machine
# ---------------------------------------------------------------------------


class TestExtractWorkflow:
    def test_idle_state_text(self, app: App) -> None:
        assert app.run_btn["text"] == "EXTRACT TO EXCEL"

    def test_start_thread_flips_to_cancel(self, app: App, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setattr(app_module.threading, "Thread", _NoThread)
        app.pst_path.set("/tmp/nonexistent.pst")
        app.start_thread()
        assert app.run_btn["text"] == "Cancel"

    def test_cancel_flips_to_cancelling(self, app: App) -> None:
        app._cancel()
        assert app.run_btn["text"] == "Cancelling..."
        assert app.cancel_event.is_set()

    def test_finish_via_after_returns_to_idle(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(app_module.threading, "Thread", _NoThread)
        app.pst_path.set("/tmp/nonexistent.pst")
        app.start_thread()
        app._cancel()
        # _finish is normally scheduled via root.after(0, _finish) by _run's
        # finally; simulate that scheduling + pump.
        root.after(0, app._finish)
        root.update()
        assert app.run_btn["text"] == "EXTRACT TO EXCEL"

    def test_finish_no_cancel_sets_ready(self, app: App, root: tk.Tk) -> None:
        app.cancel_event.clear()
        root.after(0, app._finish)
        root.update()
        assert app.status.get() == "Ready."

    def test_finish_after_cancel_sets_cancelled(self, app: App, root: tk.Tk) -> None:
        app.cancel_event.set()
        root.after(0, app._finish)
        root.update()
        assert app.status.get() == "Cancelled."


# ---------------------------------------------------------------------------
# start_thread validation branches
# ---------------------------------------------------------------------------


class TestStartThreadValidation:
    def test_no_sources_shows_error(self, app: App, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setattr(app_module.threading, "Thread", _NoThread)
        with patch("tkinter.messagebox.showerror") as mock_err:
            app.start_thread()
            mock_err.assert_called_once()
        assert app.run_btn["text"] == "EXTRACT TO EXCEL"

    def test_min_amount_invalid_shows_error(
        self, app: App, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        # Tk DoubleVar raises TclError on non-numeric .get() in some builds;
        # force the .get() to raise by patching it.
        monkeypatch.setattr(app_module.threading, "Thread", _NoThread)
        app.pst_path.set("/tmp/x.pst")

        def raise_bad() -> float:
            raise ValueError("bad number")

        monkeypatch.setattr(app.min_amount, "get", raise_bad)
        with patch("tkinter.messagebox.showerror") as mock_err:
            app.start_thread()
            mock_err.assert_called_once()
        assert app.run_btn["text"] == "EXTRACT TO EXCEL"


# ---------------------------------------------------------------------------
# _resolve_output_path branches
# ---------------------------------------------------------------------------


class TestResolveOutputPath:
    def test_batch_n_branch(self, app: App, tmp_path: Path) -> None:
        app.output_folder.set(str(tmp_path))
        path = app._resolve_output_path("stem", "xlsx", batch_n=5)
        assert path.endswith(f"stem_{date.today().isoformat()}_5.xlsx")

    def test_glob_scan_branch_finds_n_plus_1(self, app: App, tmp_path: Path) -> None:
        app.output_folder.set(str(tmp_path))
        ds = date.today().isoformat()
        # Create an existing file with N=3 to force N+1=4.
        existing = os.path.join(tmp_path, f"stem_{ds}_3.xlsx")
        open(existing, "w").close()  # noqa: S108
        path = app._resolve_output_path("stem", "xlsx")
        assert path.endswith(f"stem_{ds}_4.xlsx")

    def test_report_suffix_strip_branch(self, app: App, tmp_path: Path) -> None:
        app.output_folder.set(str(tmp_path))
        ds = date.today().isoformat()
        existing = os.path.join(tmp_path, f"stem_{ds}_2_Report.pdf")
        open(existing, "w").close()  # noqa: S108
        path = app._resolve_output_path("stem", "pdf", is_report=True)
        assert path.endswith(f"stem_{ds}_3_Report.pdf")

    def test_empty_output_folder_falls_back_to_cwd(
        self, app: App, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        app.output_folder.set("")
        monkeypatch.chdir(tmp_path)
        path = app._resolve_output_path("stem", "xlsx", batch_n=1)
        # Assert against the real cwd, not a hardcoded /tmp: macOS resolves
        # /tmp to /private/tmp and Windows has no /tmp at all.
        cwd = os.getcwd()
        assert path.startswith(cwd + os.sep)


# ---------------------------------------------------------------------------
# _run worker: full coverage of all branches
# ---------------------------------------------------------------------------


class _DataEngine:
    """Explicit collector data; runtime primitives never enter saved evidence."""

    def __init__(self, records: list | None = None) -> None:
        self.records = records if records is not None else _bill_records()
        self.error_log: list[str] = []
        self.email_count = 0
        self.pdf_count = 1
        self.filtered_records: list[dict] = []
        self.sap_contract_rows: list[dict] = []
        self.sap_meter_rows: list[dict] = []
        self.sap_financial_rows: list[dict] = []
        self.source_paths: dict = {}
        self.source_manifest: dict = {}
        self.config = {"as_of": "2026-06-01", "analysis_min": 0}
        self.cancel_event = Event()
        self.cleanup = MagicMock()
        self.crawl_pst = MagicMock()
        self.process_htm_file = MagicMock()
        self.crawl_local_pdfs = MagicMock()
        self.process_eml_folder = MagicMock()
        self.process_msg_folder = MagicMock()

    def __getstate__(self) -> dict:
        return {
            key: getattr(self, key)
            for key in (
                "records",
                "config",
                "error_log",
                "filtered_records",
                "email_count",
                "pdf_count",
                "sap_contract_rows",
                "sap_meter_rows",
                "sap_financial_rows",
                "source_paths",
                "source_manifest",
            )
        }


def _bill_records() -> list[dict]:
    return [
        {
            "Date": "31/01/2026",
            "Account": "A-TEST",
            "Fuel": "Electricity",
            "Entry Type": "New Bill",
            "Invoice #": "KI-TEST",
            "Amount (£)": 100,
            "Amount Kind": "Balance",
            "Period Charge (£)": 60,
        }
    ]


def _make_mock_engine(records: list | None = None) -> _DataEngine:
    return _DataEngine(records)


def _run_manifest(app: App) -> dict:
    return cast(
        dict,
        json.loads((Path(app._run_request["run_dir"]) / "run.json").read_text(encoding="utf-8")),
    )


def _run_worker(app: App) -> None:
    app._run_request = app._snapshot_request()
    app._save_config()
    app._run()
    app._poll_events()


def _auto_report(app: App, engine: _DataEngine, stem: str, batch_n: int) -> list[str]:
    app._run_request = app._snapshot_request()
    return cast(list[str], app._run_auto_report(engine, stem, batch_n))


class TestRunWorker:
    """Exercise the actual coordinator, storage and UI state; mock only external I/O."""

    @pytest.fixture(autouse=True)
    def _boundaries(self, monkeypatch: pytest.MonkeyPatch) -> None:
        self.mock_info = MagicMock()
        self.mock_warn = MagicMock()
        self.mock_err = MagicMock()
        for kind, mock in (
            ("info", self.mock_info),
            ("warning", self.mock_warn),
            ("error", self.mock_err),
        ):
            monkeypatch.setattr("tkinter.messagebox.show" + kind, mock)
        self.rendered: dict[str, Any] = {}
        for fmt, module in (
            ("xlsx", "edf_bill_fetcher.io.writers.export"),
            ("pdf", "edf_bill_fetcher.io.reporters.pdf_report"),
            ("docx", "edf_bill_fetcher.io.reporters.docx_report"),
            ("html", "edf_bill_fetcher.io.reporters.html_report"),
        ):

            def render(view, path, fmt=fmt):
                self.rendered[fmt] = view
                Path(path).write_bytes(b"complete renderer output")

            monkeypatch.setattr(module + ".render_" + fmt, render)

    def test_cancel_before_run_shows_cancelled_warning(self, app: App) -> None:
        app.cancel_event.set()
        engine = _make_mock_engine()
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            _run_worker(app)
        assert _run_manifest(app)["status"] == "cancelled"
        assert _run_manifest(app)["result"] is None
        assert self.mock_warn.call_args[0][0] == "Cancelled"
        assert not self.rendered
        engine.cleanup.assert_called_once()

    def test_missing_pypff_records_failed_run(
        self, app: App, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        import sys

        app.pst_path.set("/tmp/fake.pst")
        monkeypatch.setattr(app_module, "HAS_PYPFF", False)
        monkeypatch.setitem(sys.modules, "pypff", None)
        engine = _make_mock_engine(records=[])
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            _run_worker(app)
        assert _run_manifest(app)["status"] == "failed"
        assert "pypff" in _run_manifest(app)["error"]
        assert self.mock_err.call_args[0][0] == "Run Failed"
        engine.cleanup.assert_called_once()

    def test_pypff_File_fallback_closes_archive(
        self, app: App, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        app.pst_path.set("/tmp/fake.pst")
        monkeypatch.setattr(app_module, "HAS_PYPFF", True)
        archive = MagicMock()
        constructor = MagicMock(return_value=archive)
        engine = _make_mock_engine(records=[])
        with patch("edf_bill_fetcher.ui.app.pypff", SimpleNamespace(File=constructor)):
            with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
                _run_worker(app)
        constructor.assert_called_once()
        archive.open.assert_called_once_with(str(Path("/tmp/fake.pst").absolute()))
        engine.crawl_pst.assert_called_once_with(archive.get_root_folder.return_value)
        archive.close.assert_called_once()
        assert _run_manifest(app)["status"] == "empty"

    def test_pypff_missing_constructors_records_failure(
        self, app: App, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        app.pst_path.set("/tmp/fake.pst")
        monkeypatch.setattr(app_module, "HAS_PYPFF", True)
        with patch("edf_bill_fetcher.ui.app.pypff", SimpleNamespace()):
            with patch(
                "edf_bill_fetcher.collectors.engine.EvidenceEngine",
                return_value=_make_mock_engine([]),
            ):
                _run_worker(app)
        assert _run_manifest(app)["status"] == "failed"
        assert "no file constructor" in _run_manifest(app)["error"]
        self.mock_err.assert_called_once()

    def test_empty_run_preserves_result_and_shows_no_data(self, app: App) -> None:
        with patch(
            "edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=_make_mock_engine([])
        ):
            _run_worker(app)
        manifest = _run_manifest(app)
        assert manifest["status"] == "empty"
        assert manifest["result"]["sha256"]
        assert manifest["artifacts"] == {}
        assert self.mock_warn.call_args[0][0] == "No Data"

    def test_records_export_snapshot_and_reset_controls(self, app: App) -> None:
        engine = _make_mock_engine()
        app._set_extract_running()
        app.load_report_btn.config(state="disabled")
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            with patch.object(app, "_save_config") as saved:
                _run_worker(app)
        saved.assert_called_once()
        assert self.mock_info.call_args[0][0] == "Success"
        assert _run_manifest(app)["status"] == "complete"
        assert set(_run_manifest(app)["artifacts"]) == {"xlsx"}
        assert self.rendered["xlsx"].snapshot.summary.charges == 60
        assert str(app.run_btn["text"]) == "EXTRACT TO EXCEL"
        assert str(app.load_report_btn["state"]) == "normal"
        engine.cleanup.assert_called_once()

    def test_collection_errors_require_review_and_are_preserved(self, app: App) -> None:
        engine = _make_mock_engine()
        engine.error_log = ["parse error 1"]
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            _run_worker(app)
        assert self.mock_warn.call_args[0][0] == "Review Required"
        assert _run_manifest(app)["status"] == "review_required"
        saved = json.loads(
            (Path(app._run_request["run_dir"]) / "result.json").read_text(encoding="utf-8")
        )
        assert saved["evidence"]["legacy_state"]["error_log"] == ["parse error 1"]

    def test_bundle_copies_sources_and_builds_index(self, app: App) -> None:
        app.save_evidence_files_var.set(True)
        with patch(
            "edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=_make_mock_engine()
        ):
            with patch(
                "edf_bill_fetcher.io.writers.evidence_bundle.save_evidence_files", return_value={}
            ) as copy:
                with patch(
                    "edf_bill_fetcher.io.writers.evidence_bundle.build_bundle_index"
                ) as index:
                    _run_worker(app)
        copy.assert_called_once()
        index.assert_called_once()
        assert not copy.call_args[0][0].empty
        assert _run_manifest(app)["status"] == "complete"

    def test_bundle_failure_stops_before_analysis(self, app: App) -> None:
        app.save_evidence_files_var.set(True)
        engine = _make_mock_engine()
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            with patch(
                "edf_bill_fetcher.io.writers.evidence_bundle.save_evidence_files",
                side_effect=OSError("bundle full"),
            ):
                _run_worker(app)
        manifest = _run_manifest(app)
        assert manifest["status"] == "failed"
        assert manifest["stage"] == "copying_sources"
        assert manifest["result"] is None
        assert not self.rendered
        engine.cleanup.assert_called_once()

    def test_auto_reports_use_same_saved_analysis(
        self, app: App, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        app.auto_generate_report.set(True)
        app._report_options = {"format": "both", "sections": ["exec_summary"]}
        for flag in ("HAS_PDF_REPORT", "HAS_DOCX_REPORT", "HAS_HTML_REPORT"):
            monkeypatch.setattr(app_module, flag, True)
        with patch(
            "edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=_make_mock_engine()
        ):
            _run_worker(app)
        assert set(self.rendered) == {"xlsx", "pdf", "docx", "html"}
        assert len({id(view.snapshot) for view in self.rendered.values()}) == 1
        assert set(_run_manifest(app)["artifacts"]) == set(self.rendered)

    @pytest.mark.parametrize(
        ("widget", "path", "method"),
        [
            ("htm_path", "/tmp/export.htm", "process_htm_file"),
            ("pdf_dir", "/tmp/pdfs", "crawl_local_pdfs"),
            ("eml_msg_dir", "/tmp/mail", "process_eml_folder"),
        ],
    )
    def test_selected_source_reaches_collector(
        self, app: App, widget: str, path: str, method: str
    ) -> None:
        getattr(app, widget).set(path)
        engine = _make_mock_engine([])
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            _run_worker(app)
        getattr(engine, method).assert_called_once_with(path)
        assert _run_manifest(app)["status"] == "empty"

    def test_output_folder_falls_back_to_source_directory(self, app: App, tmp_path: Path) -> None:
        source = tmp_path / "export.htm"
        source.write_text("")
        app.htm_path.set(str(source))
        app.output_folder.set("")
        with patch(
            "edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=_make_mock_engine()
        ):
            _run_worker(app)
        assert Path(app._run_request["run_dir"]).parent == tmp_path
        assert _run_manifest(app)["status"] == "complete"

    def test_capture_isolated_from_widget_changes_and_repeated_runs(self, app: App) -> None:
        app.output_name.set("MyReport.xlsx")
        app.analysis_min.set(0)
        first = app._snapshot_request()
        app.analysis_min.set(999999)
        app.output_name.set("Changed.xlsx")
        app._run_request = first
        with patch(
            "edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=_make_mock_engine()
        ):
            app._run()
            app._poll_events()
        manifest = _run_manifest(app)
        assert self.rendered["xlsx"].config["analysis_min"] == 0
        assert Path(manifest["artifacts"]["xlsx"]["path"]).name == "MyReport.xlsx"
        previous = {p: p.read_bytes() for p in Path(first["run_dir"]).iterdir() if p.is_file()}
        with patch(
            "edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=_make_mock_engine()
        ):
            _run_worker(app)
        assert app._run_request["run_dir"] != first["run_dir"]
        assert all(p.read_bytes() == content for p, content in previous.items())

    def test_renderer_failure_preserves_analysis_and_resets_controls(
        self, app: App, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        def fail(view, path):
            raise OSError("disk full")

        monkeypatch.setattr("edf_bill_fetcher.io.writers.export.render_xlsx", fail)
        app._set_extract_running()
        engine = _make_mock_engine()
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            _run_worker(app)
        manifest = _run_manifest(app)
        assert manifest["status"] == "failed"
        assert manifest["result"]["sha256"]
        assert manifest["artifacts"] == {}
        assert self.mock_err.call_args[0][0] == "Run Failed"
        assert str(app.run_btn["text"]) == "EXTRACT TO EXCEL"
        engine.cleanup.assert_called_once()

    @pytest.mark.parametrize("cancel", [False, True])
    def test_real_worker_marshals_ui_and_honours_captured_inputs(
        self, app: App, tmp_path: Path, cancel: bool
    ) -> None:
        entered, release = Event(), Event()
        app.pdf_dir.set(str(tmp_path))
        app.analysis_min.set(0)
        app.scan_sap_dumps_var.set(False)
        app.generate_reconciliation_sheet_var.set(False)
        app.save_evidence_files_var.set(False)
        engine = _make_mock_engine()
        engine.cancel_event = app.cancel_event

        def collect(path):
            entered.set()
            assert release.wait(5), "Test collector was not released"

        engine.crawl_local_pdfs.side_effect = collect
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            try:
                app.start_thread()
                assert app._worker is not None
                assert entered.wait(5), "Worker never reached collection"
                first = app._run_request
                app.analysis_min.set(999999)
                app.start_thread()
                assert app._run_request is first, "A second click launched overlapping work"
                if cancel:
                    app._cancel()
                release.set()
                app._worker.join(10)
                assert not app._worker.is_alive(), "Worker did not finish"
                # Worker messages and control reset must wait for main-thread polling.
                assert not self.mock_info.called
                assert not self.mock_warn.called
                assert not self.mock_err.called
                app._poll_events()
            finally:
                release.set()
                if app._worker is not None:
                    app._worker.join(10)
        manifest = _run_manifest(app)
        assert manifest["status"] == ("cancelled" if cancel else "complete")
        assert str(app.run_btn["text"]) == "EXTRACT TO EXCEL"
        assert str(app.load_report_btn["state"]) == "normal"
        engine.cleanup.assert_called_once()
        if cancel:
            assert self.mock_warn.call_args[0][0] == "Cancelled"
            assert manifest["result"] is None
            assert not self.rendered
        else:
            assert self.mock_info.call_args[0][0] == "Success"
            config = self.rendered["xlsx"].config
            assert config["analysis_min"] == 0
            assert config["scan_sap_dumps"] is False
            assert config["generate_reconciliation_sheet"] is False
            assert config["save_evidence_files"] is False

    def test_constructor_failure_resets_controls(self, app: App) -> None:
        app._set_extract_running()
        with patch(
            "edf_bill_fetcher.collectors.engine.EvidenceEngine",
            side_effect=OSError("engine unavailable"),
        ):
            _run_worker(app)
        assert self.mock_err.call_args[0] == ("Run Failed", "engine unavailable")
        assert str(app.run_btn["text"]) == "EXTRACT TO EXCEL"


# ---------------------------------------------------------------------------
# _run_auto_report branches
# ---------------------------------------------------------------------------


class TestRunAutoReport:
    def test_pdf_only(self, app: App, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
        app.output_folder.set(str(tmp_path))
        app._report_options = {"format": "pdf", "sections": ["exec_summary"]}
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", True)
        monkeypatch.setattr(app_module, "HAS_DOCX_REPORT", False)
        engine = _make_mock_engine(records=[{"A": 1}])
        with patch(
            "edf_bill_fetcher.io.reporters.pdf_report.generate_pdf_from_gui",
            return_value=(True, "ok"),
        ):
            with patch(
                "edf_bill_fetcher.io.reporters.docx_report.generate_docx_from_gui"
            ) as mock_docx:
                written = _auto_report(app, engine, "stem", 1)
                assert len(written) == 1
                assert written[0].endswith(".pdf")
                mock_docx.assert_not_called()

    def test_docx_only(self, app: App, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
        app.output_folder.set(str(tmp_path))
        app._report_options = {"format": "docx", "sections": ["exec_summary"]}
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", False)
        monkeypatch.setattr(app_module, "HAS_DOCX_REPORT", True)
        engine = _make_mock_engine(records=[{"A": 1}])
        with patch(
            "edf_bill_fetcher.io.reporters.docx_report.generate_docx_from_gui",
            return_value=(True, "ok"),
        ):
            with patch(
                "edf_bill_fetcher.io.reporters.pdf_report.generate_pdf_from_gui"
            ) as mock_pdf:
                written = _auto_report(app, engine, "stem", 1)
                assert len(written) == 1
                assert written[0].endswith(".docx")
                mock_pdf.assert_not_called()

    def test_both_formats(self, app: App, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
        app.output_folder.set(str(tmp_path))
        app._report_options = {"format": "both", "sections": ["exec_summary"]}
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", True)
        monkeypatch.setattr(app_module, "HAS_DOCX_REPORT", True)
        monkeypatch.setattr(app_module, "HAS_HTML_REPORT", False)
        engine = _make_mock_engine(records=[{"A": 1}])
        with patch(
            "edf_bill_fetcher.io.reporters.pdf_report.generate_pdf_from_gui",
            return_value=(True, "ok"),
        ):
            with patch(
                "edf_bill_fetcher.io.reporters.docx_report.generate_docx_from_gui",
                return_value=(True, "ok"),
            ):
                written = _auto_report(app, engine, "stem", 1)
                assert len(written) == 2

    def test_html_only(self, app: App, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
        app.output_folder.set(str(tmp_path))
        app._report_options = {"format": "html", "sections": ["exec_summary"]}
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", False)
        monkeypatch.setattr(app_module, "HAS_DOCX_REPORT", False)
        monkeypatch.setattr(app_module, "HAS_HTML_REPORT", True)
        engine = _make_mock_engine(records=[{"A": 1}])
        with patch(
            "edf_bill_fetcher.io.reporters.html_report.generate_html_from_gui",
            return_value=(True, "ok"),
        ):
            written = _auto_report(app, engine, "stem", 1)
            assert len(written) == 1
            assert written[0].endswith(".html")

    def test_both_includes_html(
        self, app: App, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        app.output_folder.set(str(tmp_path))
        app._report_options = {"format": "both", "sections": ["exec_summary"]}
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", True)
        monkeypatch.setattr(app_module, "HAS_DOCX_REPORT", True)
        monkeypatch.setattr(app_module, "HAS_HTML_REPORT", True)
        engine = _make_mock_engine(records=[{"A": 1}])
        with patch(
            "edf_bill_fetcher.io.reporters.pdf_report.generate_pdf_from_gui",
            return_value=(True, "ok"),
        ):
            with patch(
                "edf_bill_fetcher.io.reporters.docx_report.generate_docx_from_gui",
                return_value=(True, "ok"),
            ):
                with patch(
                    "edf_bill_fetcher.io.reporters.html_report.generate_html_from_gui",
                    return_value=(True, "ok"),
                ):
                    written = _auto_report(app, engine, "stem", 1)
                    assert len(written) == 3
                    assert sum(1 for p in written if p.endswith(".html")) == 1

    def test_neither_available_returns_empty(
        self, app: App, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        app.output_folder.set(str(tmp_path))
        app._report_options = {"format": "both", "sections": ["exec_summary"]}
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", False)
        monkeypatch.setattr(app_module, "HAS_DOCX_REPORT", False)
        monkeypatch.setattr(app_module, "HAS_HTML_REPORT", False)
        engine = _make_mock_engine(records=[{"A": 1}])
        written = _auto_report(app, engine, "stem", 1)
        assert written == []

    def test_pdf_failure_not_written(
        self, app: App, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        app.output_folder.set(str(tmp_path))
        app._report_options = {"format": "pdf", "sections": ["exec_summary"]}
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", True)
        engine = _make_mock_engine(records=[{"A": 1}])
        with patch(
            "edf_bill_fetcher.io.reporters.pdf_report.generate_pdf_from_gui",
            return_value=(False, "err"),
        ):
            with pytest.raises(RuntimeError, match="err"):
                _auto_report(app, engine, "stem", 1)


# ---------------------------------------------------------------------------
# load_spreadsheet_and_report branches
# ---------------------------------------------------------------------------


class _SyncThread:
    """Mock Thread that runs the target synchronously on the current thread."""

    def __init__(
        self,
        target: Callable[..., Any] | None = None,
        args: tuple = (),
        kwargs: dict | None = None,
        daemon: bool = False,
    ) -> None:
        self._target = target
        self._args = args
        self._kwargs = kwargs or {}
        self.daemon = daemon

    def start(self) -> None:
        if self._target is not None:
            self._target(*self._args, **self._kwargs)

    def join(self, timeout: float | None = None) -> None:
        pass

    def is_alive(self) -> bool:
        return False


class TestLoadSpreadsheetAndReport:
    def test_no_report_libs_shows_error(self, app: App, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", False)
        monkeypatch.setattr(app_module, "HAS_DOCX_REPORT", False)
        monkeypatch.setattr(app_module, "HAS_HTML_REPORT", False)
        with patch("tkinter.messagebox.showerror") as mock_err:
            app.load_spreadsheet_and_report()
            mock_err.assert_called_once()
            assert "Report Unavailable" in mock_err.call_args[0][0]

    def test_no_file_picked_early_return(self, app: App, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", True)
        with patch("tkinter.filedialog.askopenfilename", return_value=""):
            with patch("tkinter.messagebox.showerror") as mock_err:
                app.load_spreadsheet_and_report()
                mock_err.assert_not_called()

    def test_empty_spreadsheet_shows_warning(
        self, app: App, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", True)
        with patch("tkinter.filedialog.askopenfilename", return_value="/tmp/fake.xlsx"):
            with patch(
                "edf_bill_fetcher.io.review_workbook.read_workbook", return_value=([], {}, None)
            ):
                with patch("tkinter.messagebox.showwarning") as mock_warn:
                    app.load_spreadsheet_and_report()
                    mock_warn.assert_called_once()
                    assert mock_warn.call_args[0][0] == "No Data"

    @pytest.mark.parametrize("fmt", ["both", "html"])
    def test_real_workbook_import_preserves_limitations(
        self, app: App, monkeypatch: pytest.MonkeyPatch, tmp_path: Path, fmt: str
    ) -> None:
        for flag in ("HAS_PDF_REPORT", "HAS_DOCX_REPORT", "HAS_HTML_REPORT"):
            monkeypatch.setattr(app_module, flag, True)
        app._report_options = {"format": fmt, "sections": ["exec_summary"]}
        source = tmp_path / "source.xlsx"
        with pd.ExcelWriter(source) as workbook:
            pd.DataFrame(_bill_records()).to_excel(
                workbook, sheet_name="EDF Evidence Report", index=False
            )
            pd.DataFrame([["as_of", "2026-01-31"]]).to_excel(
                workbook, sheet_name="Provenance", header=False, index=False
            )
        original = source.read_bytes()
        rendered = []

        def render(view, path):
            rendered.append(view)
            Path(path).write_bytes(b"complete report")

        for report_fmt in ("pdf", "docx", "html"):
            monkeypatch.setattr(
                f"edf_bill_fetcher.io.reporters.{report_fmt}_report.render_{report_fmt}", render
            )
        with patch.object(app_module.threading, "Thread", _SyncThread):
            with patch("tkinter.filedialog.askopenfilename", return_value=str(source)):
                with patch("tkinter.messagebox.showwarning") as warning:
                    app.load_spreadsheet_and_report()
                    app._poll_events()
        assert warning.call_args[0][0] == "Reports Generated — Workbook Import"
        assert "limited provenance" in warning.call_args[0][1]
        run_dir = next(tmp_path.glob("source_report_*"))
        manifest = json.loads((run_dir / "run.json").read_text(encoding="utf-8"))
        result = json.loads((run_dir / "result.json").read_text(encoding="utf-8"))
        assert manifest["status"] == "review_required"
        assert set(manifest["artifacts"]) == (
            {"html"} if fmt == "html" else {"pdf", "docx", "html"}
        )
        state = result["evidence"]["legacy_state"]
        assert state["provenance"] == "lossy_workbook_import"
        assert state["source_manifest"] == {}
        assert result["request"]["as_of"] == "2026-01-31"
        assert state["error_log"]
        assert len({id(view.snapshot) for view in rendered}) == 1
        assert source.read_bytes() == original
        assert str(app.load_report_btn["state"]) == "normal"

    def test_no_report_paths_resolved_warning(
        self, app: App, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        # Guard passes because HAS_PDF_REPORT=True; but fmt="docx" skips the
        # PDF path, and HAS_DOCX_REPORT=False skips the DOCX path -> no paths.
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", True)
        monkeypatch.setattr(app_module, "HAS_DOCX_REPORT", False)
        app._report_options = {"format": "docx", "sections": ["exec_summary"]}
        fake_xlsx = tmp_path / "source.xlsx"
        fake_xlsx.write_text("")
        with patch("tkinter.filedialog.askopenfilename", return_value=str(fake_xlsx)):
            with patch(
                "edf_bill_fetcher.io.review_workbook.read_workbook",
                return_value=([{"Amount": 100}], {}, None),
            ):
                with patch("tkinter.messagebox.showwarning") as mock_warn:
                    app.load_spreadsheet_and_report()
                    mock_warn.assert_called_once()
                    assert mock_warn.call_args[0][0] == "No Reports"

    def test_review_workbook_rebuilds_pack_in_legacy_desktop(
        self, app: App, tmp_path: Path
    ) -> None:
        from openpyxl import load_workbook

        from edf_bill_fetcher.io.review_workbook import HEADERS, REVIEW_SHEET
        from edf_bill_fetcher.io.writers.export import render_xlsx
        from edf_bill_fetcher.models.report_view import capture_report_view

        source = tmp_path / "review.xlsx"
        view = capture_report_view(_bill_records(), {"as_of": "2026-01-31"}, formats=("xlsx",))
        render_xlsx(view, str(source))
        workbook = load_workbook(source)
        workbook[REVIEW_SHEET].cell(5, HEADERS.index("Period Charge (£)") + 1, 85)
        workbook[REVIEW_SHEET].cell(5, 2, "Verified source")
        workbook.save(source)
        app._report_options = {"format": "html", "sections": ["exec_summary"]}
        app.analysis_min.set(999999)
        with patch.object(app_module.threading, "Thread", _SyncThread):
            with patch("tkinter.filedialog.askopenfilename", return_value=str(source)):
                with patch("tkinter.messagebox.showinfo") as notice:
                    app.load_spreadsheet_and_report()
                    app._poll_events()
        assert notice.call_args[0][0] == "Reviewed Pack Generated"
        directory = next(tmp_path.glob("review_report_*"))
        result = json.loads((directory / "result.json").read_text(encoding="utf-8"))
        manifest = json.loads((directory / "run.json").read_text(encoding="utf-8"))
        assert result["analysis"]["summary"]["charges"] == "85.00"
        assert result["request"]["as_of"] == "2026-01-31"
        assert set(manifest["artifacts"]) == {"html", "xlsx"}
        assert (directory / "evidence_manifest.json").is_file()

    def test_load_error_on_exception(self, app: App, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", True)
        with patch("tkinter.filedialog.askopenfilename", return_value="/tmp/fake.xlsx"):
            with patch(
                "edf_bill_fetcher.io.review_workbook.read_workbook",
                side_effect=RuntimeError("boom"),
            ):
                with patch("tkinter.messagebox.showerror") as mock_err:
                    app.load_spreadsheet_and_report()
                    mock_err.assert_called_once()
                    assert "Load Error" in mock_err.call_args[0][0]


# ---------------------------------------------------------------------------
# _load_config edge branches
# ---------------------------------------------------------------------------


class TestLoadConfigEdges:
    def _make_app_with_config(self, root: tk.Tk, config_path: str) -> App:
        a = _make_app(root)
        a._CONFIG_PATH = config_path
        return a

    def test_missing_file_silent(self, root: tk.Tk, tmp_path: Path) -> None:
        path = str(tmp_path / "missing.json")
        assert not os.path.exists(path)
        app = self._make_app_with_config(root, path)
        app._load_config()
        assert app.output_folder.get() == ""

    def test_malformed_json_silent(self, root: tk.Tk, tmp_path: Path) -> None:
        path = tmp_path / "bad.json"
        path.write_text("{not valid json")
        app = self._make_app_with_config(root, str(path))
        app._load_config()
        assert app.output_folder.get() == ""

    def test_float_cast_value_error_silent(self, root: tk.Tk, tmp_path: Path) -> None:
        path = tmp_path / "bad_float.json"
        payload = {"gui_state": {"min_amount": "not-a-number"}, "report_options": {}}
        path.write_text(json.dumps(payload))
        app = self._make_app_with_config(root, str(path))
        app._load_config()
        # ValueError caught -> min_amount stays at default 50.0
        assert app.min_amount.get() == 50.0

    def test_report_options_loaded(self, root: tk.Tk, tmp_path: Path) -> None:
        path = tmp_path / "with_ro.json"
        payload = {
            "gui_state": {"output_folder": "/tmp/loaded"},
            "report_options": {"format": "pdf", "sections": ["exec_summary"]},
        }
        path.write_text(json.dumps(payload))
        app = self._make_app_with_config(root, str(path))
        app._load_config()
        assert app.output_folder.get() == "/tmp/loaded"
        assert app._report_options == {"format": "pdf", "sections": ["exec_summary"]}

    def test_bool_and_str_keys_loaded(self, root: tk.Tk, tmp_path: Path) -> None:
        path = tmp_path / "full.json"
        payload = {
            "gui_state": {
                "use_anchors": False,
                "acc_num": "A-999",
                "output_name": "Custom.xlsx",
            },
            "report_options": {},
        }
        path.write_text(json.dumps(payload))
        app = self._make_app_with_config(root, str(path))
        app._load_config()
        assert app.use_anchors.get() is False
        assert app.acc_num.get() == "A-999"
        assert app.output_name.get() == "Custom.xlsx"
