"""Section 1 output-folder picker -- spec Design Section 1."""

import tkinter as tk
from collections.abc import Iterator

import pytest

from edf_collector import App


def _walk_children(widget: tk.Misc) -> Iterator[tk.Misc]:
    for child in widget.winfo_children():
        yield child
        yield from _walk_children(child)


def _all_widget_text(widget: tk.Misc) -> list[str]:
    texts: list[str] = []
    for child in _walk_children(widget):
        try:
            cls = child.winfo_class()
            if cls in ("Label", "Button", "Checkbutton", "TLabel", "TButton"):
                t = child.cget("text")
                if t and isinstance(t, str):
                    texts.append(t)
        except tk.TclError:
            pass
    return texts


@pytest.fixture
def app():
    root = tk.Tk()
    root.withdraw()
    try:
        yield App(root)
    finally:
        root.destroy()


class TestOutputFolderPickerUI:
    def test_output_folder_label_exists(self, app):
        texts = _all_widget_text(app.root)
        assert any("Output Folder:" in t for t in texts)

    def test_output_filename_label_exists(self, app):
        # Moved from Section 2 to Section 1
        texts = _all_widget_text(app.root)
        assert any("Output filename:" in t for t in texts)

    def test_output_folder_var_set_get(self, app):
        app.output_folder.set("/tmp/test")
        assert app.output_folder.get() == "/tmp/test"

    def test_output_folder_empty_defaults_to_source_dir(self):
        root = tk.Tk()
        root.withdraw()
        try:
            app = App(root)
            assert app.output_folder.get() == ""
        finally:
            root.destroy()
