"""Sequential output-file naming -- spec Section 2.

Algorithm: ISO YYYY-MM-DD date stamp + shared per-batch counter.
Glob {stem}_{date}_*.{ext} in output_folder, find max N, use N+1.
Pass batch_n= to reuse the same counter across xlsx+pdf+docx.
"""

import os
import tkinter as tk
from unittest.mock import patch

import pytest

from edf_bill_fetcher.ui.app import App


@pytest.fixture
def app_with_tmp_folder(tmp_path):
    root = tk.Tk()
    root.withdraw()
    try:
        app = App(root)
        app._CONFIG_PATH = str(tmp_path / "config.json")
        app.output_folder.set(str(tmp_path))
        yield app, tmp_path
    finally:
        root.destroy()


class TestResolveOutputPath:
    def test_empty_folder_returns_n1(self, app_with_tmp_folder):
        app, tmp = app_with_tmp_folder
        with patch("edf_bill_fetcher.ui.app.date") as mock_date:
            mock_date.today.return_value.isoformat.return_value = "2026-07-10"
            path = app._resolve_output_path("EDF_Dispute_Evidence", "xlsx")
        assert path == str(tmp / "EDF_Dispute_Evidence_2026-07-10_1.xlsx")

    def test_occupied_folder_increments(self, app_with_tmp_folder):
        app, tmp = app_with_tmp_folder
        (tmp / "EDF_Dispute_Evidence_2026-07-10_1.xlsx").touch()
        with patch("edf_bill_fetcher.ui.app.date") as mock_date:
            mock_date.today.return_value.isoformat.return_value = "2026-07-10"
            path = app._resolve_output_path("EDF_Dispute_Evidence", "xlsx")
        assert path == str(tmp / "EDF_Dispute_Evidence_2026-07-10_2.xlsx")

    def test_shared_batch_counter(self, app_with_tmp_folder):
        app, tmp = app_with_tmp_folder
        with patch("edf_bill_fetcher.ui.app.date") as mock_date:
            mock_date.today.return_value.isoformat.return_value = "2026-07-10"
            xlsx = app._resolve_output_path("EDF_Dispute_Evidence", "xlsx")
            pdf = app._resolve_output_path("EDF_Dispute_Evidence", "pdf", batch_n=1, is_report=True)
            docx = app._resolve_output_path(
                "EDF_Dispute_Evidence", "docx", batch_n=1, is_report=True
            )
        assert "EDF_Dispute_Evidence_2026-07-10_1.xlsx" in xlsx
        assert "EDF_Dispute_Evidence_2026-07-10_1_Report.pdf" in pdf
        assert "EDF_Dispute_Evidence_2026-07-10_1_Report.docx" in docx

    def test_report_suffix_appended(self, app_with_tmp_folder):
        app, tmp = app_with_tmp_folder
        with patch("edf_bill_fetcher.ui.app.date") as mock_date:
            mock_date.today.return_value.isoformat.return_value = "2026-07-10"
            path = app._resolve_output_path("EDF_Dispute_Evidence", "pdf", is_report=True)
        assert "_Report.pdf" in path

    def test_empty_output_folder_falls_back(self, app_with_tmp_folder):
        app, tmp = app_with_tmp_folder
        app.output_folder.set("")
        with patch("edf_bill_fetcher.ui.app.date") as mock_date:
            mock_date.today.return_value.isoformat.return_value = "2026-07-10"
            path = app._resolve_output_path("EDF_Dispute_Evidence", "xlsx")
        assert os.path.dirname(path) == os.getcwd()

    def test_counter_resets_per_day(self, app_with_tmp_folder):
        app, tmp = app_with_tmp_folder
        (tmp / "EDF_Dispute_Evidence_2026-07-09_5.xlsx").touch()
        with patch("edf_bill_fetcher.ui.app.date") as mock_date:
            mock_date.today.return_value.isoformat.return_value = "2026-07-10"
            path = app._resolve_output_path("EDF_Dispute_Evidence", "xlsx")
        assert "_2026-07-10_1.xlsx" in path

    def test_non_numeric_suffixes_ignored(self, app_with_tmp_folder):
        app, tmp = app_with_tmp_folder
        (tmp / "EDF_Dispute_Evidence_2026-07-10_abc.xlsx").touch()
        with patch("edf_bill_fetcher.ui.app.date") as mock_date:
            mock_date.today.return_value.isoformat.return_value = "2026-07-10"
            path = app._resolve_output_path("EDF_Dispute_Evidence", "xlsx")
        assert "_1.xlsx" in path

    def test_glob_metachars_in_stem_treated_literally(self, app_with_tmp_folder):
        """A stem containing glob metacharacters (e.g. ``[test]``) must not
        be interpreted as a pattern by the counter's glob scan.

        Without ``glob.escape`` the ``[test]`` character class would match
        unrelated single-char filenames and miss the existing file, so the
        counter would restart at 1 and collide with it.  The escaped stem
        finds the existing ``_1`` file and returns ``_2``.
        """
        app, tmp = app_with_tmp_folder
        (tmp / "EDF_[test]_2026-07-10_1.xlsx").touch()
        with patch("edf_bill_fetcher.ui.app.date") as mock_date:
            mock_date.today.return_value.isoformat.return_value = "2026-07-10"
            path = app._resolve_output_path("EDF_[test]", "xlsx")
        assert path == str(tmp / "EDF_[test]_2026-07-10_2.xlsx")
