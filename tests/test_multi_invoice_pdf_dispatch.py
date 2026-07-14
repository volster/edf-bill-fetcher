from __future__ import annotations

from unittest.mock import MagicMock, patch


def _fake_page(text: str) -> MagicMock:
    p = MagicMock()
    p.extract_text.return_value = text
    return p


def _fake_pdf(pages: list[str]) -> MagicMock:
    pdf = MagicMock()
    pdf.pages = [_fake_page(t) for t in pages]
    return pdf


def _build_engine():
    from edf_collector import EvidenceEngine

    eng = EvidenceEngine(config={}, update_ui_cb=lambda *a, **k: None)
    eng.config["use_acc_filter"] = False
    eng.config["use_anchors"] = False
    eng.config["use_large"] = False
    eng.config["min_amount"] = 0
    return eng


def test_process_pdf_file_dispatches_once_for_single_invoice_pdf():
    eng = _build_engine()
    pages = [
        "Invoice number: KI-31105244-0004\nDate issued: 07 October 2024\nPage 1 of 4",
        "Your charges in detail\nPage 2 of 4",
        "Get in touch\nPage 3 of 4",
        "More\nPage 4 of 4",
    ]
    fake = _fake_pdf(pages)
    with (
        patch(
            "builtins.open",
            return_value=MagicMock(
                __enter__=lambda *_a: MagicMock(read=lambda: b"dummy"),
                __exit__=lambda *_a: None,
            ),
        ),
        patch(
            "edf_collector.pdfplumber.open",
            return_value=MagicMock(__enter__=lambda _s: fake, __exit__=lambda *a: None),
        ),
        patch.object(eng, "_process_new_invoice") as mock_dispatch,
    ):
        eng.process_pdf_file(
            "/tmp/fake_invoice.pdf",
            source_label="Local PDF Folder",
            detail_label="fake_invoice.pdf",
            fallback_date=None,
        )
    assert mock_dispatch.call_count == 1
    call = mock_dispatch.call_args
    sent_text = call.args[0]
    assert "Invoice number: KI-31105244-0004" in sent_text
    assert "Page 4 of 4" in sent_text
    assert "Page 2 of 4" in sent_text
    assert call.args[2] == "fake_invoice.pdf"


def test_process_pdf_file_dispatches_once_per_slice_for_merged_pdf():
    eng = _build_engine()
    pages = [
        "Invoice number: KI-31105244-0001\nDate issued: 06 August 2024\nPage 1 of 4",
        "Page 2 of 4",
        "Page 3 of 4",
        "Page 4 of 4",
        "Invoice number: KI-31105244-0002\nDate issued: 06 September 2024\nPage 1 of 4",
        "Page 2 of 4",
        "Page 3 of 4",
        "Page 4 of 4",
        "Invoice number: KI-31105244-0003\nDate issued: 07 October 2024\nPage 1 of 4",
        "Page 2 of 4",
        "Page 3 of 4",
        "Page 4 of 4",
    ]
    fake = _fake_pdf(pages)
    with (
        patch(
            "builtins.open",
            return_value=MagicMock(
                __enter__=lambda *_a: MagicMock(read=lambda: b"dummy"),
                __exit__=lambda *_a: None,
            ),
        ),
        patch(
            "edf_collector.pdfplumber.open",
            return_value=MagicMock(__enter__=lambda _s: fake, __exit__=lambda *a: None),
        ),
        patch.object(eng, "_process_new_invoice") as mock_dispatch,
    ):
        eng.process_pdf_file(
            "/tmp/D-merged.pdf",
            source_label="Local PDF Folder",
            detail_label="D-merged.pdf",
            fallback_date=None,
        )
    assert mock_dispatch.call_count == 3
    texts = [c.args[0] for c in mock_dispatch.call_args_list]
    labels = [c.args[2] for c in mock_dispatch.call_args_list]
    attachments = [c.kwargs["attachment_name"] for c in mock_dispatch.call_args_list]
    assert "KI-31105244-0001" in texts[0]
    assert "KI-31105244-0002" not in texts[0]
    assert "KI-31105244-0002" in texts[1]
    assert "KI-31105244-0003" not in texts[1]
    assert "KI-31105244-0003" in texts[2]
    assert labels == [
        "D-merged.pdf #1",
        "D-merged.pdf #2",
        "D-merged.pdf #3",
    ]
    assert attachments == [
        "D-merged.pdf #1",
        "D-merged.pdf #2",
        "D-merged.pdf #3",
    ]


def test_process_pdf_file_single_slice_failure_does_not_lose_other_slices():
    eng = _build_engine()
    pages = [
        "Invoice number: KI-A\nDate issued: 01 January 2025\nPage 1 of 4",
        "Page 2 of 4",
        "Page 3 of 4",
        "Page 4 of 4",
        "Invoice number: KI-B\nDate issued: 02 February 2025\nPage 1 of 4",
        "Page 2 of 4",
        "Page 3 of 4",
        "Page 4 of 4",
        "Invoice number: KI-C\nDate issued: 03 March 2025\nPage 1 of 4",
        "Page 2 of 4",
        "Page 3 of 4",
        "Page 4 of 4",
    ]
    fake = _fake_pdf(pages)
    call_log: list[str] = []

    def fake_dispatch(
        text, source_label, detail_label, fallback_date, sender="", attachment_name=""
    ):
        call_log.append(detail_label)
        if detail_label.endswith("#2"):
            raise RuntimeError("slice-2 boom")

    with (
        patch(
            "builtins.open",
            return_value=MagicMock(
                __enter__=lambda *_a: MagicMock(read=lambda: b"dummy"),
                __exit__=lambda *_a: None,
            ),
        ),
        patch(
            "edf_collector.pdfplumber.open",
            return_value=MagicMock(__enter__=lambda _s: fake, __exit__=lambda *a: None),
        ),
        patch.object(eng, "_process_new_invoice", side_effect=fake_dispatch),
        patch.object(eng, "log_error") as mock_log,
    ):
        eng.process_pdf_file("/tmp/abc.pdf", "Local PDF Folder", "abc.pdf", None)
    assert call_log == ["abc.pdf #1", "abc.pdf #2", "abc.pdf #3"]
    assert mock_log.call_count == 1
    logged = mock_log.call_args.args[0]
    assert "abc.pdf #2" in logged
