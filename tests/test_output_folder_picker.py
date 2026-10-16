"""Section 1 output-folder picker -- spec Design Section 1.

The ``app`` fixture mocks ``App._load_config`` to a no-op so tests observe
the HARDCODED defaults in ``App.__init__`` rather than whatever the
developer's local ``~/.edf_collector/config.json`` happens to contain.
Without this mock, tests pass on machines where the GUI has never been
launched but fail unpredictably on developer machines that have actually
saved GUI state.
"""

import tkinter as tk
from collections.abc import Iterator

import pytest

from edf_bill_fetcher.ui.app import App


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
def app(monkeypatch):
    monkeypatch.setattr(App, "_load_config", lambda self: None)
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

    def test_output_folder_empty_defaults_to_source_dir(self, monkeypatch):
        monkeypatch.setattr(App, "_load_config", lambda self: None)
        root = tk.Tk()
        root.withdraw()
        try:
            app = App(root)
            assert app.output_folder.get() == ""
        finally:
            root.destroy()
