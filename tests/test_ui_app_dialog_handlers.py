"""Wave 3 coverage: dialog handlers, thread-marshalling, _run, config branches.

Drives ``edf_bill_fetcher/ui/app.py`` from 43% toward >=95% by exercising
modal handlers via boundary mocks (filedialog/messagebox), the EXTRACT
workflow state machine, thread-marshalling branches, the full ``_run``
worker, ``load_spreadsheet_and_report``, ``_resolve_output_path``,
``_run_auto_report``, ``ReportOptionsDialog``, and ``_load_config`` edges.

Boundary-mock strategy: patch ONLY the filedialog/messagebox call, then
assert STATE mutations on tk vars / button text. Never assert mock
call_count or call_args.
"""

from __future__ import annotations

import json
import os
import tkinter as tk
from collections.abc import Callable, Generator
from datetime import date
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, patch

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
def app(root: tk.Tk) -> App:
    return _make_app(root)


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
        self, app: App, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        app.output_folder.set("")
        monkeypatch.chdir("/tmp")
        path = app._resolve_output_path("stem", "xlsx", batch_n=1)
        assert path.startswith("/tmp" + os.sep)


# ---------------------------------------------------------------------------
# _run worker: full coverage of all branches
# ---------------------------------------------------------------------------


def _make_mock_engine(records: list | None = None) -> MagicMock:
    """Build a MagicMock that quacks like EvidenceEngine with populated attrs."""
    engine = MagicMock()
    engine.records = records if records is not None else [{"Amount": 100}]
    engine.error_log = []
    engine.email_count = 0
    engine.pdf_count = 1
    engine.filtered_records = []
    engine.sap_contract_rows = []
    engine.sap_meter_rows = []
    engine.sap_financial_rows = []
    engine.source_paths = {}
    return engine


class TestRunWorker:
    """Drive _run directly (it's the thread target) with mocked dependencies.

    The autouse ``_suppress_messagebox`` fixture patches all three
    ``messagebox.show*`` calls and exposes them as ``self.mock_info``,
    ``self.mock_warn``, ``self.mock_err`` so no real modal dialog can block
    and individual tests can assert on the recorded calls.
    """

    @pytest.fixture(autouse=True)
    def _suppress_messagebox(self, monkeypatch: pytest.MonkeyPatch) -> None:
        self.mock_info = MagicMock()
        self.mock_warn = MagicMock()
        self.mock_err = MagicMock()
        monkeypatch.setattr("tkinter.messagebox.showinfo", self.mock_info)
        monkeypatch.setattr("tkinter.messagebox.showwarning", self.mock_warn)
        monkeypatch.setattr("tkinter.messagebox.showerror", self.mock_err)

    def test_cancel_before_run_shows_cancelled_warning(self, app: App, root: tk.Tk) -> None:
        app.pst_path.set("/tmp/none.pst")
        app.cancel_event.set()
        with patch(
            "edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=_make_mock_engine()
        ):
            app._run()
            root.update()
            root.update()
        self.mock_warn.assert_called_once()
        assert self.mock_warn.call_args[0][0] == "Cancelled"

    def test_no_pypff_shows_warning(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(os.path, "exists", lambda p: True)
        app.pst_path.set("/tmp/fake.pst")
        monkeypatch.setattr(app_module, "HAS_PYPFF", False)
        engine = _make_mock_engine(records=[])
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            app._run()
            root.update()
            root.update()
        calls = self.mock_warn.call_args_list
        assert any(c[0][0] == "PST" for c in calls)

    def test_pypff_attribute_error_fallback_to_File(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(os.path, "exists", lambda p: True)
        app.pst_path.set("/tmp/fake.pst")
        monkeypatch.setattr(app_module, "HAS_PYPFF", True)

        fake_pff_instance = MagicMock()
        fake_pff_instance.get_root_folder.return_value = MagicMock()

        class _FakePypff:
            def file(self) -> None:
                raise AttributeError("no file attr")

            File = MagicMock(return_value=fake_pff_instance)

        fake_pypff = _FakePypff()
        engine = _make_mock_engine(records=[])
        with patch("edf_bill_fetcher.ui.app.pypff", fake_pypff):
            with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
                app._run()
                root.update()
                root.update()
        fake_pypff.File.assert_called_once()

    def test_pypff_file_and_File_both_missing_raises(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(os.path, "exists", lambda p: True)
        app.pst_path.set("/tmp/fake.pst")
        monkeypatch.setattr(app_module, "HAS_PYPFF", True)

        class _EmptyPypff:
            pass

        engine = _make_mock_engine(records=[])
        with patch("edf_bill_fetcher.ui.app.pypff", _EmptyPypff()):
            with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
                app._run()
                root.update()
                root.update()
        self.mock_err.assert_called_once()
        assert "Error" in self.mock_err.call_args[0][0]

    def test_records_empty_shows_no_data_warning(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        app.pdf_dir.set("/tmp/pdfs")
        monkeypatch.setattr(os.path, "exists", lambda p: True)
        engine = _make_mock_engine(records=[])
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            app._run()
            root.update()
            root.update()
        calls = self.mock_warn.call_args_list
        assert any(c[0][0] == "No Data" for c in calls)

    def test_records_present_exports_excel_and_shows_summary(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        app.output_folder.set(str(tmp_path))
        app.output_name.set("TestEvidence.xlsx")
        engine = _make_mock_engine(records=[{"Amount": 100}, {"Amount": 200}])
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            with patch("edf_bill_fetcher.io.writers.export.export_to_excel") as mock_export:
                with patch.object(app, "_save_config") as mock_save:
                    app._run()
                    root.update()
                    root.update()
                    mock_export.assert_called_once()
                    mock_save.assert_called_once()
        self.mock_info.assert_called_once()
        assert self.mock_info.call_args[0][0] == "Success"

    def test_records_present_with_error_log_in_summary(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        app.output_folder.set(str(tmp_path))
        engine = _make_mock_engine(records=[{"Amount": 100}])
        engine.error_log = ["parse error 1"]
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            with patch("edf_bill_fetcher.io.writers.export.export_to_excel"):
                app._run()
                root.update()
                root.update()
        summary = self.mock_info.call_args[0][1]
        assert "Parse errors: 1" in summary

    def test_save_evidence_files_bundle_path(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        app.output_folder.set(str(tmp_path))
        app.save_evidence_files_var.set(True)
        engine = _make_mock_engine(records=[{"Amount": 100}])
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            with patch("edf_bill_fetcher.io.writers.export.export_to_excel"):
                with patch("edf_bill_fetcher.ui.app.pd") as mock_pd:
                    mock_pd.DataFrame.return_value = MagicMock()
                    with patch(
                        "edf_bill_fetcher.io.writers.evidence_bundle.save_evidence_files", return_value={"a": "b"}
                    ) as mock_save_ev:
                        with patch("edf_bill_fetcher.io.writers.evidence_bundle.build_bundle_index") as mock_build:
                            app._run()
                            root.update()
                            root.update()
                            mock_save_ev.assert_called_once()
                            mock_build.assert_called_once()

    def test_save_evidence_files_bundle_failure_shows_warning(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        app.output_folder.set(str(tmp_path))
        app.save_evidence_files_var.set(True)
        engine = _make_mock_engine(records=[{"Amount": 100}])
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            with patch("edf_bill_fetcher.io.writers.export.export_to_excel"):
                with patch("edf_bill_fetcher.ui.app.pd") as mock_pd:
                    mock_pd.DataFrame.return_value = MagicMock()
                    with patch(
                        "edf_bill_fetcher.io.writers.evidence_bundle.save_evidence_files", side_effect=RuntimeError("boom")
                    ):
                        app._run()
                        root.update()
                        root.update()
        calls = self.mock_warn.call_args_list
        assert any(c[0][0] == "Bundle step failed" for c in calls)

    def test_auto_report_path_when_toggle_on(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        app.output_folder.set(str(tmp_path))
        app.auto_generate_report.set(True)
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", True)
        monkeypatch.setattr(app_module, "HAS_DOCX_REPORT", True)
        engine = _make_mock_engine(records=[{"Amount": 100}])
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            with patch("edf_bill_fetcher.io.writers.export.export_to_excel"):
                with patch.object(
                    app, "_run_auto_report", return_value=["/tmp/r.pdf"]
                ) as mock_auto:
                    app._run()
                    root.update()
                    root.update()
                    mock_auto.assert_called_once_with(engine, "EDF_Dispute_Evidence", 1)
        summary = self.mock_info.call_args[0][1]
        assert "Reports:" in summary

    def test_htm_branch_processes(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(os.path, "exists", lambda p: True)
        app.htm_path.set("/tmp/export.htm")
        engine = _make_mock_engine(records=[])
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            app._run()
            root.update()
            root.update()
            engine.process_htm_file.assert_called_once_with("/tmp/export.htm")

    def test_pdf_branch_crawls(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(os.path, "exists", lambda p: True)
        app.pdf_dir.set("/tmp/pdfs")
        engine = _make_mock_engine(records=[])
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            app._run()
            root.update()
            root.update()
            engine.crawl_local_pdfs.assert_called_once_with("/tmp/pdfs")

    def test_output_folder_fallback_to_source_dir_when_empty(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        pst_file = tmp_path / "archive.pst"
        pst_file.write_text("")
        app.pst_path.set(str(pst_file))
        app.output_folder.set("")
        app.output_name.set("Out.xlsx")
        # Suppress the PST branch so pypff doesn't try to open the fake file.
        monkeypatch.setattr(app_module, "HAS_PYPFF", False)
        engine = _make_mock_engine(records=[{"Amount": 100}])
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            with patch("edf_bill_fetcher.io.writers.export.export_to_excel"):
                app._run()
                root.update()
                root.update()
                assert app.output_folder.get() == str(tmp_path)

    def test_stem_strips_xlsx_suffix(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        app.output_folder.set(str(tmp_path))
        app.output_name.set("MyReport.xlsx")
        engine = _make_mock_engine(records=[{"Amount": 100}])
        with patch("edf_bill_fetcher.collectors.engine.EvidenceEngine", return_value=engine):
            with patch("edf_bill_fetcher.io.writers.export.export_to_excel") as mock_export:
                app._run()
                root.update()
                root.update()
                xlsx_path = mock_export.call_args[0][1]
                assert "MyReport_" in xlsx_path
                assert not xlsx_path.endswith(".xlsx.xlsx")


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
                written = app._run_auto_report(engine, "stem", 1)
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
                written = app._run_auto_report(engine, "stem", 1)
                assert len(written) == 1
                assert written[0].endswith(".docx")
                mock_pdf.assert_not_called()

    def test_both_formats(self, app: App, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
        app.output_folder.set(str(tmp_path))
        app._report_options = {"format": "both", "sections": ["exec_summary"]}
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", True)
        monkeypatch.setattr(app_module, "HAS_DOCX_REPORT", True)
        engine = _make_mock_engine(records=[{"A": 1}])
        with patch(
            "edf_bill_fetcher.io.reporters.pdf_report.generate_pdf_from_gui",
            return_value=(True, "ok"),
        ):
            with patch(
                "edf_bill_fetcher.io.reporters.docx_report.generate_docx_from_gui",
                return_value=(True, "ok"),
            ):
                written = app._run_auto_report(engine, "stem", 1)
                assert len(written) == 2

    def test_neither_available_returns_empty(
        self, app: App, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        app.output_folder.set(str(tmp_path))
        app._report_options = {"format": "both", "sections": ["exec_summary"]}
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", False)
        monkeypatch.setattr(app_module, "HAS_DOCX_REPORT", False)
        engine = _make_mock_engine(records=[{"A": 1}])
        written = app._run_auto_report(engine, "stem", 1)
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
            written = app._run_auto_report(engine, "stem", 1)
            assert written == []


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


class TestLoadSpreadsheetAndReport:
    def test_no_report_libs_shows_error(self, app: App, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", False)
        monkeypatch.setattr(app_module, "HAS_DOCX_REPORT", False)
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
            with patch("edf_bill_fetcher.ui.app.pd") as mock_pd:
                mock_df = MagicMock()
                mock_df.empty = True
                mock_pd.read_excel.return_value = mock_df
                with patch("tkinter.messagebox.showwarning") as mock_warn:
                    app.load_spreadsheet_and_report()
                    mock_warn.assert_called_once()
                    assert mock_warn.call_args[0][0] == "No Data"

    def test_happy_path_generates_reports(
        self, app: App, root: tk.Tk, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", True)
        monkeypatch.setattr(app_module, "HAS_DOCX_REPORT", True)
        app._report_options = {"format": "both", "sections": ["exec_summary"]}
        app.output_folder.set(str(tmp_path))
        fake_xlsx = tmp_path / "source.xlsx"
        fake_xlsx.write_text("")

        with patch.object(app_module.threading, "Thread", _SyncThread):
            with patch("tkinter.filedialog.askopenfilename", return_value=str(fake_xlsx)):
                with patch("edf_bill_fetcher.ui.app.pd") as mock_pd:
                    mock_df = MagicMock()
                    mock_df.empty = False
                    mock_df.to_dict.return_value = [{"Amount": 100}]
                    mock_pd.read_excel.return_value = mock_df
                    with patch(
                        "edf_bill_fetcher.io.reporters.pdf_report.generate_pdf_from_gui",
                        return_value=(True, "PDF done"),
                    ):
                        with patch(
                            "edf_bill_fetcher.io.reporters.docx_report.generate_docx_from_gui",
                            return_value=(True, "DOCX done"),
                        ):
                            with patch("tkinter.messagebox.showinfo") as mock_info:
                                app.load_spreadsheet_and_report()
                                root.update()
                                root.update()
                                assert mock_info.called
                                assert mock_info.call_args[0][0] == "Reports Generated"

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
            with patch("edf_bill_fetcher.ui.app.pd") as mock_pd:
                mock_df = MagicMock()
                mock_df.empty = False
                mock_df.to_dict.return_value = [{"Amount": 100}]
                mock_pd.read_excel.return_value = mock_df
                with patch("tkinter.messagebox.showwarning") as mock_warn:
                    app.load_spreadsheet_and_report()
                    mock_warn.assert_called_once()
                    assert mock_warn.call_args[0][0] == "No Reports"

    def test_load_error_on_exception(self, app: App, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setattr(app_module, "HAS_PDF_REPORT", True)
        with patch("tkinter.filedialog.askopenfilename", return_value="/tmp/fake.xlsx"):
            with patch("edf_bill_fetcher.ui.app.pd") as mock_pd:
                mock_pd.read_excel.side_effect = RuntimeError("boom")
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
