"""Section 2 + Section 3 -- relocate save_filtered, auto-generate, relabel dedup, amalgamate."""

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


def _find_checkbutton_by_label(root: tk.Misc, label_substr: str) -> tk.Checkbutton | None:
    """Find a checkbutton whose text contains the given substring."""
    for child in _walk_children(root):
        if child.winfo_class() == "Checkbutton":
            try:
                text = str(child.cget("text"))
                if label_substr in text:
                    return child  # type: ignore[return-value]
            except tk.TclError:
                pass
    return None


def _find_button_by_label(root: tk.Misc, label_substr: str) -> tk.Button | None:
    """Find a button whose text contains the given substring."""
    for child in _walk_children(root):
        if child.winfo_class() in ("Button", "TButton"):
            try:
                text = str(child.cget("text"))
                if label_substr in text:
                    return child  # type: ignore[return-value]
            except tk.TclError:
                pass
    return None


@pytest.fixture
def app(monkeypatch):
    # Mock ``App._load_config`` so tests observe the HARDCODED defaults in
    # ``App.__init__`` rather than whatever the developer's local
    # ``~/.edf_collector/config.json`` has been saved to by real GUI use.
    monkeypatch.setattr(App, "_load_config", lambda self: None)
    root = tk.Tk()
    root.withdraw()
    try:
        yield App(root)
    finally:
        root.destroy()


class TestFilterBelowRelocate:
    def test_save_filtered_label_reworded(self, app):
        texts = _all_widget_text(app.root)
        assert any("Keep filtered-out records" in t for t in texts)

    def test_save_filtered_default_on(self, app):
        assert app.save_filtered.get() is True

    def test_save_filtered_disabled_when_filter_below_off(self, app):
        """Filter-below checkbox toggled OFF should disable the save_filtered child."""
        app.filter_below.set(True)  # start True so invoke toggles to False
        chk_filt = _find_checkbutton_by_label(app.root, "Filter results below")
        assert chk_filt is not None
        chk_filt.invoke()  # toggles to False, then runs _update_sf_state
        assert app.filter_below.get() is False
        chk_sf = _find_checkbutton_by_label(app.root, "Keep filtered-out")
        assert chk_sf is not None
        assert str(chk_sf.cget("state")) == "disabled"

    def test_save_filtered_enabled_when_filter_below_on(self, app):
        """Filter-below checkbox toggled ON should enable the save_filtered child."""
        app.filter_below.set(False)  # start False so invoke toggles to True
        chk_filt = _find_checkbutton_by_label(app.root, "Filter results below")
        assert chk_filt is not None
        chk_filt.invoke()  # toggles to True, then runs _update_sf_state
        assert app.filter_below.get() is True
        chk_sf = _find_checkbutton_by_label(app.root, "Keep filtered-out")
        assert chk_sf is not None
        assert str(chk_sf.cget("state")) == "normal"


class TestAutoGenerateReport:
    def test_auto_generate_label_exists(self, app):
        texts = _all_widget_text(app.root)
        assert any("Auto-generate report" in t for t in texts)

    def test_auto_generate_defaults_false(self, app):
        assert app.auto_generate_report.get() is False

    def test_report_options_button_exists(self, app):
        texts = _all_widget_text(app.root)
        assert any("Report Options" in t for t in texts)


class TestDedupLabels:
    def test_use_dedup_label_reworded(self, app):
        texts = _all_widget_text(app.root)
        assert any("Drop duplicates found across sources" in t for t in texts)

    def test_save_dups_label_reworded(self, app):
        texts = _all_widget_text(app.root)
        assert any("Record dropped duplicates on side sheet" in t for t in texts)


class TestAmalgamateToggle:
    def test_amalgamate_label_exists(self, app):
        texts = _all_widget_text(app.root)
        assert any("Build hybrid row" in t for t in texts)

    def test_amalgamate_defaults_false(self, app):
        assert app.amalgamate_duplicates.get() is False

    def test_amalgamate_disabled_when_dedup_off(self, app):
        """use_dedup toggled OFF should disable amalgamate."""
        app.use_dedup.set(True)  # start True so invoke toggles to False
        app.save_dups.set(True)
        chk_dup = _find_checkbutton_by_label(app.root, "Drop duplicates")
        assert chk_dup is not None
        chk_dup.invoke()  # toggles use_dedup to False
        assert app.use_dedup.get() is False
        chk_am = _find_checkbutton_by_label(app.root, "Build hybrid row")
        assert chk_am is not None
        assert str(chk_am.cget("state")) == "disabled"

    def test_amalgamate_disabled_when_save_dups_off(self, app):
        """save_dups toggled OFF should disable amalgamate."""
        app.use_dedup.set(True)
        app.save_dups.set(True)  # start True so invoke toggles to False
        chk_sd = _find_checkbutton_by_label(app.root, "Record dropped duplicates")
        assert chk_sd is not None
        chk_sd.invoke()  # toggles save_dups to False
        assert app.save_dups.get() is False
        chk_am = _find_checkbutton_by_label(app.root, "Build hybrid row")
        assert chk_am is not None
        assert str(chk_am.cget("state")) == "disabled"

    def test_amalgamate_enabled_when_both_on(self, app):
        """use_dedup + save_dups both ON should enable amalgamate."""
        app.use_dedup.set(False)  # start False so invoke toggles to True
        app.save_dups.set(True)
        chk_dup = _find_checkbutton_by_label(app.root, "Drop duplicates")
        assert chk_dup is not None
        chk_dup.invoke()  # toggles use_dedup to True
        assert app.use_dedup.get() is True
        # Need to also trigger save_dups command to update amalgamate state
        app.save_dups.set(True)  # already True, but trigger command
        chk_sd = _find_checkbutton_by_label(app.root, "Record dropped duplicates")
        assert chk_sd is not None
        # Toggle save_dups off then on to trigger command
        chk_sd.invoke()  # toggles to False
        assert app.save_dups.get() is False
        chk_sd.invoke()  # toggles back to True
        assert app.save_dups.get() is True
        chk_am = _find_checkbutton_by_label(app.root, "Build hybrid row")
        assert chk_am is not None
        assert str(chk_am.cget("state")) == "normal"
