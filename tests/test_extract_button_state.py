"""Three-state EXTRACT button -- spec Design Section 4.

Idle (orange, EXTRACT TO EXCEL) -> Running (navy, Cancel) ->
Cancelling (grey, Cancelling...) -> Idle.
"""

import tkinter as tk

import pytest

from edf_collector import EDF_NAVY, EDF_ORANGE, MEDIUM_GREY, App


class _NoThread:
    """Mock threading.Thread that doesn't actually start."""

    def __init__(self, *a, **kw):
        pass

    def start(self):
        pass

    daemon = False


@pytest.fixture
def app(monkeypatch):
    root = tk.Tk()
    root.withdraw()
    try:
        # Patch threading.Thread so start_thread doesn't actually spawn
        import edf_collector

        monkeypatch.setattr(edf_collector.threading, "Thread", _NoThread)
        yield App(root)
    finally:
        root.destroy()


class TestExtractButtonStates:
    def test_idle_state(self, app):
        assert app.run_btn.cget("text") == "EXTRACT TO EXCEL"
        assert app.run_btn.cget("bg") == EDF_ORANGE

    def test_no_cancel_button_exists(self, app):
        """The separate Cancel button should be gone."""
        # The old attribute should not exist
        assert not hasattr(app, "cancel_btn") or not isinstance(
            getattr(app, "cancel_btn", None), tk.Misc
        )

    def test_start_thread_flips_to_running(self, app):
        app.pst_path.set("/tmp/nonexistent.pst")
        app.cancel_event.clear()
        app.start_thread()
        assert app.run_btn.cget("text") == "Cancel"
        assert app.run_btn.cget("bg") == EDF_NAVY

    def test_cancel_flips_to_cancelling(self, app):
        app.pst_path.set("/tmp/nonexistent.pst")
        app.cancel_event.clear()
        app.start_thread()
        app._cancel()
        assert app.run_btn.cget("text") == "Cancelling..."
        assert app.run_btn.cget("bg") == MEDIUM_GREY

    def test_finish_flips_back_to_idle(self, app):
        app.pst_path.set("/tmp/nonexistent.pst")
        app.cancel_event.clear()
        app.start_thread()
        app._cancel()
        app._finish()
        assert app.run_btn.cget("text") == "EXTRACT TO EXCEL"
        assert app.run_btn.cget("fg") == "white"
        assert app.run_btn.cget("bg") == EDF_ORANGE

    def test_finish_after_no_cancel_stays_idle(self, app):
        """Finish from plain Running state should still go to Idle."""
        app.pst_path.set("/tmp/nonexistent.pst")
        app.cancel_event.clear()
        app.start_thread()
        app._finish()
        assert app.run_btn.cget("text") == "EXTRACT TO EXCEL"

    def test_report_options_button_in_actionbar(self, app):
        assert app.report_options_btn.cget("text") == "Report Options"

    def test_load_report_button_in_actionbar(self, app):
        assert app.load_report_btn.cget("text") == "LOAD & REPORT"
