"""Tests for edf_bill_fetcher.ui submodule.

Verifies that the UI classes are importable from the ui package and
behave correctly.
"""

from __future__ import annotations


def test_ui_submodule_importable():
    from edf_bill_fetcher import ui

    assert ui is not None


def test_app_class_importable():
    from edf_bill_fetcher.ui import App

    assert App is not None


def test_report_options_dialog_class_importable():
    from edf_bill_fetcher.ui import ReportOptionsDialog

    assert ReportOptionsDialog is not None


def test_ui_re_exported_from_edf_collector():
    from edf_bill_fetcher.ui import App as App_ui
    from edf_bill_fetcher.ui.app import App as App_collector

    assert App_collector is App_ui

    from edf_bill_fetcher.ui import ReportOptionsDialog as ROD_ui
    from edf_bill_fetcher.ui.app import ReportOptionsDialog as ROD_collector

    assert ROD_collector is ROD_ui
