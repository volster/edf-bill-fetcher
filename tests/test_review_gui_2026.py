"""Regression checks for persisted desktop inputs and visible selection ownership."""

import json
import os
import tkinter as tk

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QCoreApplication, QEvent  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from edf_bill_fetcher.application.evidence_review import ReviewEntry  # noqa: E402
from edf_bill_fetcher.domain.json_values import freeze_json  # noqa: E402
from edf_bill_fetcher.ui.app import App, ReportOptionsDialog  # noqa: E402
from edf_bill_fetcher.ui.qt_review import EvidenceReviewPage, EvidenceTable  # noqa: E402


@pytest.fixture
def legacy_settings(tmp_path):
    """Use real Tcl variables without needing a window or altering user settings."""
    interpreter = tk.Tcl()
    app = App.__new__(App)
    app._CONFIG_PATH = str(tmp_path / "settings.json")
    booleans = {
        "use_anchors": True,
        "use_large": True,
        "use_reading_class": True,
        "use_pdf_fields": True,
        "use_acc_filt": False,
        "filter_below": False,
        "save_filtered": True,
        "use_dedup": True,
        "save_dups": True,
        "amalgamate_duplicates": False,
        "use_domain_filter": True,
        "auto_generate_report": False,
        "save_evidence_files_var": True,
        "scan_sap_dumps_var": True,
        "generate_reconciliation_sheet_var": True,
    }
    for name, value in booleans.items():
        setattr(app, name, tk.BooleanVar(interpreter, value))
    for name in ("acc_num", "report_account_ref", "output_folder"):
        setattr(app, name, tk.StringVar(interpreter, ""))
    app.domain_filter = tk.StringVar(interpreter, "edfenergy.com")
    app.output_name = tk.StringVar(interpreter, "EDF_Dispute_Evidence.xlsx")
    app.min_amount = tk.DoubleVar(interpreter, 50.0)
    app.analysis_min = tk.DoubleVar(interpreter, 0.0)
    app._report_options = {}
    return app


@pytest.mark.parametrize("payload", [None, [], "settings", {"gui_state": None}, {"gui_state": []}])
def test_legacy_malformed_setting_containers_retain_defaults(legacy_settings, payload):
    with open(legacy_settings._CONFIG_PATH, "w", encoding="utf-8") as handle:
        json.dump(payload, handle)
    legacy_settings._load_config()
    assert legacy_settings.use_anchors.get() is True
    assert legacy_settings.min_amount.get() == 50
    assert legacy_settings.output_folder.get() == ""


def test_legacy_invalid_setting_values_do_not_enable_filters_or_rewrite_paths(legacy_settings):
    payload = {
        "gui_state": {
            "use_acc_filter": "false",
            "filter_below": 1,
            "acc_num": ["wrong", "account"],
            "output_folder": None,
            "min_amount": float("inf"),
            "analysis_min": -1,
        },
        "report_options": "pdf",
    }
    with open(legacy_settings._CONFIG_PATH, "w", encoding="utf-8") as handle:
        json.dump(payload, handle)
    legacy_settings._load_config()
    assert legacy_settings.use_acc_filt.get() is False
    assert legacy_settings.filter_below.get() is False
    assert legacy_settings.acc_num.get() == ""
    assert legacy_settings.output_folder.get() == ""
    assert legacy_settings.min_amount.get() == 50
    assert legacy_settings.analysis_min.get() == 0
    assert legacy_settings._report_options == {}


@pytest.mark.parametrize(
    "options",
    [
        [],
        "pdf",
        {"format": "wrong", "sections": ["exec_summary"]},
        {"format": "pdf", "sections": "exec_summary"},
        {"format": "pdf", "sections": []},
        {"format": "pdf", "sections": ["unknown"]},
        {"format": "pdf", "sections": [{}]},
    ],
)
def test_legacy_invalid_report_options_do_not_poison_future_runs(legacy_settings, options):
    with open(legacy_settings._CONFIG_PATH, "w", encoding="utf-8") as handle:
        json.dump({"report_options": options}, handle)
    legacy_settings._load_config()
    assert legacy_settings._report_options == {}


def test_legacy_valid_settings_preserve_aliases_and_supported_report_content(legacy_settings):
    payload = {
        "gui_state": {"use_acc_filt": True, "acc_num": "A-001", "min_amount": "12.5"},
        "report_options": {"format": "html", "sections": ["payment", "payment"]},
    }
    with open(legacy_settings._CONFIG_PATH, "w", encoding="utf-8") as handle:
        json.dump(payload, handle)
    legacy_settings._load_config()
    assert legacy_settings.use_acc_filt.get() is True
    assert legacy_settings.acc_num.get() == "A-001"
    assert legacy_settings.min_amount.get() == 12.5
    assert legacy_settings._report_options == {"format": "html", "sections": ["payment"]}
    assert "payment" in {key for key, _, _ in ReportOptionsDialog.SECTIONS}


def test_legacy_invalid_utf8_settings_do_not_prevent_opening(legacy_settings):
    with open(legacy_settings._CONFIG_PATH, "wb") as handle:
        handle.write(b"\xff\xfe")
    legacy_settings._load_config()
    assert legacy_settings.output_folder.get() == ""


def test_evidence_deselection_clears_details_and_original_text():
    application = QApplication.instance() or QApplication([])
    assert isinstance(application, QApplication)
    page = EvidenceReviewPage()
    try:
        model = EvidenceTable(
            (
                ReviewEntry(
                    "Bills & charges",
                    freeze_json({"Invoice #": "B1", "Source PDF Text": "B1 original text"}),
                    (),
                ),
            ),
            page,
        )
        page.model = model
        page.proxy.setSourceModel(model)
        page.table.selectRow(0)
        application.processEvents()
        assert "B1" in page.details.toPlainText()
        assert page.extracted_text.toPlainText() == "B1 original text"
        page.table.clearSelection()
        application.processEvents()
        assert not page.table.selectionModel().selectedRows()
        assert page.details.toPlainText() == ""
        assert page.extracted_text.toPlainText() == ""
        assert not page.open_source.isEnabled()
    finally:
        page.close()
        page.deleteLater()
        QCoreApplication.sendPostedEvents(None, QEvent.Type.DeferredDelete)
        application.processEvents()
