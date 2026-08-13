"""Config-file persistence for App GUI state.

Config path: ~/.edf_collector/config.json; atomic write (temp+rename);
silent fallback to defaults when file missing/unreadable.
"""

import json
import os
import sys
import tkinter as tk

import pytest

from edf_bill_fetcher.ui.app import App

pytestmark = pytest.mark.skipif(
    sys.platform == "win32",
    reason=(
        "Windows CI intermittently fails with _tkinter.TclError: "
        "Can't find a usable tk.tcl"
    ),
)


@pytest.fixture
def tmp_config_path(tmp_path):
    """Provide a temp config file path for override."""
    return tmp_path / "config.json"


def _make_app(root, config_path, monkeypatch=None):
    """Construct App and override _CONFIG_PATH to the temp path.

    ``App.__init__`` calls ``self._load_config()`` before we get a chance to
    override ``_CONFIG_PATH``, so a developer's real ``~/.edf_collector/config.json``
    (if one exists on the machine running the tests) would leak into the App
    instance and pollute assertions in ``test_load_config_missing_file_silent``
    / ``test_load_config_malformed_json_silent`` (which expect pristine
    hardcoded defaults before they call ``_load_config()`` themselves).

    To prevent that poisoning, we patch ``_load_config`` to a no-op for the
    duration of ``App(root)``. The caller is responsible for invoking
    ``app._load_config()`` explicitly after we've set ``_CONFIG_PATH`` to the
    temp path; that call runs against the temp file as intended.
    """
    if monkeypatch is None:
        # Late VueVue-style fallback for callers that didn't inject monkeypatch.
        # Without it, we can't suppress the constructor's _load_config call.
        # Fall through to legacy behaviour (real config may poison the var) —
        # the caller gets a deprecation-via-assertion hint if this matters.
        app = App(root)
        app._CONFIG_PATH = str(config_path)
        return app

    # Suppress the constructor's _load_config call so the real
    # ~/.edf_collector/config.json cannot poison the App instance before
    # we redirect _CONFIG_PATH to the temp file.
    monkeypatch.setattr(App, "_load_config", lambda self: None)
    app = App(root)
    monkeypatch.undo()  # restore the real _load_config for explicit test calls
    app._CONFIG_PATH = str(config_path)
    return app


class TestConfigLoadSave:
    """Round-trip and edge-case tests for _load_config / _save_config."""

    def test_save_config_writes_valid_json(self, tmp_config_path, monkeypatch):
        root = tk.Tk()
        root.withdraw()
        try:
            app = _make_app(root, tmp_config_path, monkeypatch)
            app.output_folder.set("/tmp/test_output")
            app._save_config()
            assert tmp_config_path.exists()
            data = json.loads(tmp_config_path.read_text())
            assert data["gui_state"]["output_folder"] == "/tmp/test_output"
        finally:
            root.destroy()

    def test_load_config_round_trips_state(self, tmp_config_path, monkeypatch):
        root = tk.Tk()
        root.withdraw()
        try:
            app = _make_app(root, tmp_config_path, monkeypatch)
            app.output_folder.set("/tmp/abc")
            app.amalgamate_duplicates.set(True)
            app.auto_generate_report.set(True)
            app._save_config()

            root2 = tk.Tk()
            root2.withdraw()
            try:
                app2 = _make_app(root2, tmp_config_path, monkeypatch)
                app2._load_config()
                assert app2.output_folder.get() == "/tmp/abc"
                assert app2.amalgamate_duplicates.get() is True
                assert app2.auto_generate_report.get() is True
            finally:
                root2.destroy()
        finally:
            root.destroy()

    def test_load_config_missing_file_silent(self, tmp_config_path, monkeypatch):
        """Missing config file -> defaults preserved, no crash."""
        assert not tmp_config_path.exists()
        root = tk.Tk()
        root.withdraw()
        try:
            app = _make_app(root, tmp_config_path, monkeypatch)
            app._load_config()
            assert app.output_folder.get() == ""
            assert app.amalgamate_duplicates.get() is False
            assert app.auto_generate_report.get() is False
        finally:
            root.destroy()

    def test_load_config_malformed_json_silent(self, tmp_config_path, monkeypatch):
        """Corrupt config file -> defaults preserved, no crash."""
        tmp_config_path.parent.mkdir(parents=True, exist_ok=True)
        tmp_config_path.write_text("{not valid json")
        root = tk.Tk()
        root.withdraw()
        try:
            app = _make_app(root, tmp_config_path, monkeypatch)
            app._load_config()
            assert app.output_folder.get() == ""
            assert app.amalgamate_duplicates.get() is False
        finally:
            root.destroy()

    def test_save_config_atomic_write(self, tmp_config_path, monkeypatch):
        """Config written via temp+rename; no partial file on crash."""
        root = tk.Tk()
        root.withdraw()
        try:
            app = _make_app(root, tmp_config_path, monkeypatch)
            app._save_config()
            # No .tmp file should remain after atomic rename
            assert not (tmp_config_path.parent / "config.json.tmp").exists()
        finally:
            root.destroy()

    def test_save_config_includes_report_options(self, tmp_config_path, monkeypatch):
        root = tk.Tk()
        root.withdraw()
        try:
            app = _make_app(root, tmp_config_path, monkeypatch)
            app._report_options = {"format": "pdf", "sections": ["exec_summary"]}
            app._save_config()
            data = json.loads(tmp_config_path.read_text())
            assert data["report_options"]["format"] == "pdf"
            assert "exec_summary" in data["report_options"]["sections"]
        finally:
            root.destroy()

    @pytest.mark.skipif(os.name == "nt", reason="Unix permission bits not enforced on Windows")
    def test_save_config_file_permissions_0600(self, tmp_config_path, monkeypatch):
        """Config file should have 0o600 permissions (user-only)."""
        root = tk.Tk()
        root.withdraw()
        try:
            app = _make_app(root, tmp_config_path, monkeypatch)
            app._save_config()
            mode = tmp_config_path.stat().st_mode & 0o777
            assert mode == 0o600
        finally:
            root.destroy()
