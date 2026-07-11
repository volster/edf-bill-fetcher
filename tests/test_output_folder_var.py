"""App.__init__ declares output_folder, amalgamate_duplicates, auto_generate_report."""

import tkinter as tk

from edf_collector import App


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
