"""Config-file persistence for App GUI state.

Spec: docs/superpowers/specs/2026-07-10-ui-refresh-design.md Design - Section 3.
Config path: ~/.edf_collector/config.json; atomic write (temp+rename);
silent fallback to defaults when file missing/unreadable.
"""

import json
import tkinter as tk

import pytest

from edf_collector import App


@pytest.fixture
def tmp_config_path(tmp_path):
    """Provide a temp config file path for override."""
    return tmp_path / "config.json"


def _make_app(root, config_path):
    """Construct App and override _CONFIG_PATH to the temp path."""
    app = App(root)
    # Override instance attribute so _load_config / _save_config use temp
    app._CONFIG_PATH = str(config_path)
    return app


class TestConfigLoadSave:
    """Round-trip and edge-case tests for _load_config / _save_config."""

    def test_save_config_writes_valid_json(self, tmp_config_path):
        root = tk.Tk()
        root.withdraw()
        try:
            app = _make_app(root, tmp_config_path)
            app.output_folder.set("/tmp/test_output")
            app._save_config()
            assert tmp_config_path.exists()
            data = json.loads(tmp_config_path.read_text())
            assert data["gui_state"]["output_folder"] == "/tmp/test_output"
        finally:
            root.destroy()

    def test_load_config_round_trips_state(self, tmp_config_path):
        root = tk.Tk()
        root.withdraw()
        try:
            app = _make_app(root, tmp_config_path)
            app.output_folder.set("/tmp/abc")
            app.amalgamate_duplicates.set(True)
            app.auto_generate_report.set(True)
            app._save_config()

            root2 = tk.Tk()
            root2.withdraw()
            try:
                app2 = App(root2)
                app2._CONFIG_PATH = str(tmp_config_path)
                app2._load_config()
                assert app2.output_folder.get() == "/tmp/abc"
                assert app2.amalgamate_duplicates.get() is True
                assert app2.auto_generate_report.get() is True
            finally:
                root2.destroy()
        finally:
            root.destroy()

    def test_load_config_missing_file_silent(self, tmp_config_path):
        """Missing config file -> defaults preserved, no crash."""
        assert not tmp_config_path.exists()
        root = tk.Tk()
        root.withdraw()
        try:
            app = _make_app(root, tmp_config_path)
            app._load_config()
            assert app.output_folder.get() == ""
            assert app.amalgamate_duplicates.get() is False
            assert app.auto_generate_report.get() is False
        finally:
            root.destroy()

    def test_load_config_malformed_json_silent(self, tmp_config_path):
        """Corrupt config file -> defaults preserved, no crash."""
        tmp_config_path.parent.mkdir(parents=True, exist_ok=True)
        tmp_config_path.write_text("{not valid json")
        root = tk.Tk()
        root.withdraw()
        try:
            app = _make_app(root, tmp_config_path)
            app._load_config()
            assert app.output_folder.get() == ""
            assert app.amalgamate_duplicates.get() is False
        finally:
            root.destroy()

    def test_save_config_atomic_write(self, tmp_config_path):
        """Config written via temp+rename; no partial file on crash."""
        root = tk.Tk()
        root.withdraw()
        try:
            app = _make_app(root, tmp_config_path)
            app._save_config()
            # No .tmp file should remain after atomic rename
            assert not (tmp_config_path.parent / "config.json.tmp").exists()
        finally:
            root.destroy()

    def test_save_config_includes_report_options(self, tmp_config_path):
        root = tk.Tk()
        root.withdraw()
        try:
            app = _make_app(root, tmp_config_path)
            app._report_options = {"format": "pdf", "sections": ["exec_summary"]}
            app._save_config()
            data = json.loads(tmp_config_path.read_text())
            assert data["report_options"]["format"] == "pdf"
            assert "exec_summary" in data["report_options"]["sections"]
        finally:
            root.destroy()

    def test_save_config_file_permissions_0600(self, tmp_config_path):
        """Config file should have 0o600 permissions (user-only)."""
        root = tk.Tk()
        root.withdraw()
        try:
            app = _make_app(root, tmp_config_path)
            app._save_config()
            mode = tmp_config_path.stat().st_mode & 0o777
            assert mode == 0o600
        finally:
            root.destroy()
