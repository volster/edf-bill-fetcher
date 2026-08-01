"""ReportOptionsDialog persists format+sections on OK via _open_report_options."""

import tkinter as tk
from unittest.mock import MagicMock, patch

import pytest

from edf_bill_fetcher.ui.app import App


@pytest.fixture
def app(tmp_path):
    root = tk.Tk()
    root.withdraw()
    try:
        app = App(root)
        app._CONFIG_PATH = str(tmp_path / "config.json")
        yield app
    finally:
        root.destroy()


class TestReportDialogPersist:
    def test_open_report_options_persists_on_ok(self, app):
        """_open_report_options should save to config on OK."""
        fake_dialog = MagicMock()
        fake_dialog.show.return_value = {"format": "pdf", "sections": ["exec_summary"]}
        with patch("edf_collector.ReportOptionsDialog", return_value=fake_dialog):
            with patch.object(app, "_save_config") as mock_save:
                app._open_report_options()
                assert app._report_options == {
                    "format": "pdf",
                    "sections": ["exec_summary"],
                }
                mock_save.assert_called_once()

    def test_open_report_options_no_save_on_cancel(self, app):
        """Dialog cancelled: no change to _report_options, no _save_config."""
        fake_dialog = MagicMock()
        fake_dialog.show.return_value = None
        original = app._report_options.copy()
        with patch("edf_collector.ReportOptionsDialog", return_value=fake_dialog):
            with patch.object(app, "_save_config") as mock_save:
                app._open_report_options()
                assert app._report_options == original
                mock_save.assert_not_called()
