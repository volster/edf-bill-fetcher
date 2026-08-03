"""App.__init__ declares output_folder, amalgamate_duplicates, auto_generate_report.

All tests below mock ``App._load_config`` to a no-op so they observe the
HARDCODED defaults in ``App.__init__`` rather than whatever the developer's
local ``~/.edf_collector/config.json`` happens to contain (which would
override via ``.set(bool(...))`` / ``.set(str(...))`` inside the loader).
Without this mock, tests pass on machines where the GUI has never been
launched but fail unpredictably on developer machines that have actually
saved GUI state —_fragility that has bitten us on Windows.
"""

import tkinter as tk

import pytest

from edf_bill_fetcher.ui.app import App


@pytest.fixture(autouse=True)
def _no_load_config(monkeypatch):
    monkeypatch.setattr(App, "_load_config", lambda self: None)


class TestNewVarDeclarations:
    def test_output_folder_var_is_empty_stringvar(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = App(root)
            assert app.output_folder.get() == ""
        finally:
            root.destroy()

    def test_amalgamate_duplicates_var_defaults_false(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = App(root)
            assert app.amalgamate_duplicates.get() is False
        finally:
            root.destroy()

    def test_auto_generate_report_var_defaults_false(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = App(root)
            assert app.auto_generate_report.get() is False
        finally:
            root.destroy()

    def test_report_options_attr_exists(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = App(root)
            assert hasattr(app, "_report_options")
        finally:
            root.destroy()
