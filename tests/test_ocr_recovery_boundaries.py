"""OCR failure paths must release native handles, children and temporary files."""

import subprocess
from unittest.mock import Mock

import pypdfium2 as pdfium
import pytest

from edf_bill_fetcher.io.adapters import ocr

HEADER = "level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tconf\ttext\n"


@pytest.mark.parametrize(
    "dimensions", [(0, 100), (-1, 100), (100, float("inf")), (float("nan"), 100)]
)
def test_invalid_page_dimensions_close_document_without_rendering(
    tmp_path, monkeypatch, dimensions
):
    page = Mock()
    page.get_size.return_value = dimensions
    document = Mock()
    document.__getitem__ = Mock(return_value=page)
    monkeypatch.setattr(pdfium, "PdfDocument", Mock(return_value=document))
    with pytest.raises(ValueError, match="Invalid PDF page dimensions"):
        ocr._render_page(b"original", 0, tmp_path / "page.png")
    page.render.assert_not_called()
    page.close.assert_called_once()
    document.close.assert_called_once()
    assert not (tmp_path / "page.png").exists()


@pytest.mark.parametrize("boundary", ["page", "render", "convert"])
def test_native_failure_releases_all_created_handles(tmp_path, monkeypatch, boundary):
    document, page, bitmap = Mock(), Mock(), Mock()
    document.__getitem__ = Mock(return_value=page)
    page.get_size.return_value = (600, 800)
    page.render.return_value = bitmap
    failure = OSError("damaged scan")
    if boundary == "page":
        document.__getitem__.side_effect = failure
    elif boundary == "render":
        page.render.side_effect = failure
    else:
        bitmap.to_pil.return_value.convert.side_effect = failure
    monkeypatch.setattr(pdfium, "PdfDocument", Mock(return_value=document))
    with pytest.raises(OSError, match="damaged scan"):
        ocr._render_page(b"original", 0, tmp_path / "page.png")
    document.close.assert_called_once()
    assert page.close.call_count == (0 if boundary == "page" else 1)
    assert bitmap.close.call_count == (1 if boundary == "convert" else 0)


@pytest.mark.parametrize("phase", ["before_discovery", "after_render"])
def test_early_cancellation_never_starts_child(tmp_path, monkeypatch, phase):
    discover = Mock(return_value=ocr.TesseractRuntime(tmp_path / "tesseract", None))
    images = []

    def render(data, index, path):
        path.write_bytes(b"temporary raster")
        images.append(path)

    launch = Mock()
    monkeypatch.setattr(ocr, "find_tesseract", discover)
    monkeypatch.setattr(ocr, "_render_page", render)
    monkeypatch.setattr(ocr.subprocess, "Popen", launch)
    with pytest.raises(ocr.OcrCancelled):
        ocr.PdfOcrReader().read_page(
            b"original", 0, lambda: phase == "before_discovery" or bool(images)
        )
    launch.assert_not_called()
    assert discover.call_count == (0 if phase == "before_discovery" else 1)
    assert all(not path.parent.exists() for path in images)


@pytest.mark.parametrize("failure_phase", ["render", "launch"])
def test_page_failure_cleans_raster_and_allows_next_page(tmp_path, monkeypatch, failure_phase):
    discover = Mock(return_value=ocr.TesseractRuntime(tmp_path / "tesseract", None))
    images = []

    def render(data, index, path):
        path.write_bytes(b"temporary raster")
        images.append(path)
        if index == 0 and failure_phase == "render":
            raise OSError("raster failed")

    process = Mock(returncode=0)
    process.communicate.return_value = (
        (HEADER + "5\t1\t1\t1\t1\t1\t87.5\t£125.00\n").encode(),
        b"",
    )
    process.poll.return_value = 0
    launch = Mock(
        side_effect=[OSError("launch failed"), process] if failure_phase == "launch" else [process]
    )
    monkeypatch.setattr(ocr, "find_tesseract", discover)
    monkeypatch.setattr(ocr, "_render_page", render)
    monkeypatch.setattr(ocr.subprocess, "Popen", launch)
    reader = ocr.PdfOcrReader()
    with pytest.raises(OSError, match="failed"):
        reader.read_page(b"original", 0, lambda: False)
    assert reader.read_page(b"original", 1, lambda: False) == ocr.OcrPage("£125.00", 87.5)
    assert discover.call_count == 1
    assert len(images) == 2
    assert all(not path.parent.exists() for path in images)
    process.kill.assert_not_called()


def test_slow_success_retries_communicate_without_restarting_child(tmp_path, monkeypatch):
    monkeypatch.setattr(
        ocr, "find_tesseract", lambda: ocr.TesseractRuntime(tmp_path / "tesseract", None)
    )
    monkeypatch.setattr(ocr, "_render_page", lambda *args: None)
    process = Mock(returncode=0)
    process.poll.return_value = 0
    process.communicate.side_effect = [
        subprocess.TimeoutExpired("tesseract", 0.2),
        (HEADER.encode(), b""),
    ]
    launch = Mock(return_value=process)
    monkeypatch.setattr(ocr.subprocess, "Popen", launch)
    assert ocr.PdfOcrReader().read_page(b"original", 0, lambda: False) == ocr.OcrPage("", None)
    launch.assert_called_once()
    assert process.communicate.call_count == 2
    process.kill.assert_not_called()


@pytest.mark.parametrize("confidence", ["-0.01", "100.01", "inf", "-inf", "nan", "unreadable"])
def test_invalid_confidence_never_becomes_accepted_evidence(confidence):
    with pytest.raises(ValueError):
        ocr._parse_tsv(HEADER + f"5\t1\t1\t1\t1\t1\t{confidence}\t£125.00\n")
