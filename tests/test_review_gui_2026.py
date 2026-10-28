"""The Qt review selection owns its displayed record and original extraction text."""

from PySide6.QtCore import QCoreApplication, QEvent
from PySide6.QtWidgets import QApplication

from edf_bill_fetcher.application.evidence_review import ReviewEntry
from edf_bill_fetcher.domain.json_values import freeze_json
from edf_bill_fetcher.ui.qt_review import EvidenceReviewPage, EvidenceTable


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
