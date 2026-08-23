"""Main-window scroll + geometry smoke test (post-release wave Task 16).

Pins the layout contract for the App main window:
- content lives in a scrollable canvas with an always-visible ttk.Scrollbar
- default geometry height fits 768px-tall screens (780x700)
- minsize(720, 600) keeps the layout usable when shrunk
- mousewheel scrolling works on Windows/macOS (<MouseWheel>) and
  X11 (<Button-4>/<Button-5>)
- the orange header stays pinned outside the scroll area

Pure layout assertions -- no extraction behaviour is exercised here.
"""

from __future__ import annotations

import tkinter as tk
from collections.abc import Generator
from tkinter import ttk

import pytest

from edf_bill_fetcher.ui.app import EDF_ORANGE, App


def _make_app(root: tk.Tk) -> App:
    """Construct App with _load_config suppressed (mirrors test_ui_app_dialog_handlers)."""
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
    try:
        r.update_idletasks()
    except tk.TclError:
        pass
    r.destroy()


@pytest.fixture
def app(root: tk.Tk) -> App:
    return _make_app(root)


def _walk(widget: tk.Misc) -> Generator[tk.Misc, None, None]:
    """Yield widget and every descendant."""
    yield widget
    for child in widget.winfo_children():
        yield from _walk(child)


class TestMainWindowScroll:
    def test_scrollbar_widget_exists(self, app: App, root: tk.Tk) -> None:
        scrollbars = [w for w in _walk(root) if isinstance(w, ttk.Scrollbar)]
        assert scrollbars, "main window must contain an always-visible ttk.Scrollbar"

    def test_canvas_scroll_region_wired(self, app: App, root: tk.Tk) -> None:
        canvases = [w for w in _walk(root) if isinstance(w, tk.Canvas)]
        assert canvases, "main window content must live in a tk.Canvas"
        canvas = canvases[0]
        # Scrollbar command wired both ways: yscrollcommand drives the scrollbar.
        assert canvas.cget("yscrollcommand") != ""

    def test_default_geometry_fits_768px_screen(self, app: App, root: tk.Tk) -> None:
        # geometry() reports 1x1 while withdrawn, so map the window first;
        # under xvfb (no WM) the requested 780x700 is honored exactly.
        root.deiconify()
        root.update_idletasks()
        assert root.winfo_height() <= 768, (
            f"default height {root.winfo_height()} must fit a 768px-tall screen"
        )
        assert root.winfo_width() == 780
        root.withdraw()

    def test_minsize_keeps_layout_usable(self, app: App, root: tk.Tk) -> None:
        assert root.minsize() == (720, 600)

    def test_mousewheel_bindings_present(self, app: App, root: tk.Tk) -> None:
        # <MouseWheel> covers Windows/macOS; <Button-4>/<Button-5> cover X11.
        # Bindings live on the toplevel bindtag so wheel events over any
        # descendant of the main window scroll the canvas, and so the
        # ReportOptionsDialog's bind_all/unbind_all lifecycle cannot strip them.
        assert root.bind("<MouseWheel>"), "main window must scroll on <MouseWheel>"
        assert root.bind("<Button-4>"), "main window must scroll on <Button-4> (X11 wheel up)"
        assert root.bind("<Button-5>"), "main window must scroll on <Button-5> (X11 wheel down)"

    def test_header_pinned_outside_scroll_area(self, app: App, root: tk.Tk) -> None:
        # The orange header is a direct child of root, not inside the canvas.
        headers = [
            w
            for w in root.winfo_children()
            if isinstance(w, tk.Frame) and w.cget("bg") == EDF_ORANGE
        ]
        assert headers, "orange header must be pinned at the top of the main window"
