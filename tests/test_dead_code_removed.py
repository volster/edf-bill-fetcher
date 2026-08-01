"""Dead code removal -- export_report and _export_legacy removed."""

from edf_bill_fetcher.ui.app import App


def test_export_report_removed():
    assert not hasattr(App, "export_report")


def test_export_legacy_removed():
    assert not hasattr(App, "_export_legacy")
