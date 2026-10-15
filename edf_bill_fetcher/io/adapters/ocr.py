"""Offline OCR of original PDF pages using a local or bundled Tesseract engine."""

from __future__ import annotations

import csv
import io
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from edf_bill_fetcher.helpers.pdf_utils import PDF_LOCK


class OcrCancelled(Exception):
    """The caller cancelled a running page's OCR."""


@dataclass(frozen=True)
class TesseractRuntime:
    """Resolved executable and optional explicit English language-data directory."""

    executable: Path
    tessdata: Path | None


@dataclass(frozen=True)
class OcrPage:
    """Recognized text and the lowest word confidence, on Tesseract's 0–100 scale."""

    text: str
    minimum_confidence: float | None


def requires_ocr(text: str | None, page: Any) -> bool:
    """Detect empty pages or sparse headers over scans, excluding small logos."""
    text = text or ""
    if not text.strip():
        return True
    if len(re.sub(r"\W", "", text)) >= 80 or re.search(r"£\s*\d", text):
        return False
    images = page.images
    width, height = page.width, page.height
    if not isinstance(images, list) or not all(
        isinstance(v, (int, float)) and v > 0 for v in (width, height)
    ):
        return False
    area = sum(image.get("width", 0) * image.get("height", 0) for image in images)
    return bool(area >= width * height * 0.2)


def find_tesseract() -> TesseractRuntime:
    """Prefer packaged OCR, then the user's Tesseract installation, without downloads."""
    name = "tesseract.exe" if os.name == "nt" else "tesseract"
    bundled = Path(getattr(sys, "_MEIPASS", "")) / "ocr"
    if getattr(sys, "frozen", False):
        if not (bundled / name).is_file():
            raise RuntimeError(
                "The packaged OCR executable is missing; use an updated packaged build"
            )
        if not (bundled / "tessdata" / "eng.traineddata").is_file():
            raise RuntimeError("The packaged OCR English language data is missing")
        return TesseractRuntime(bundled / name, bundled / "tessdata")
    found = shutil.which("tesseract")
    candidates = [Path(found)] if found else []
    if os.name == "nt":
        candidates.append(
            Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Tesseract-OCR" / name
        )
    for executable in candidates:
        if not executable.is_file():
            continue
        prefix = os.environ.get("TESSDATA_PREFIX")
        locations = ([Path(prefix)] if prefix else []) + [
            executable.parent / "tessdata",
            executable.resolve().parent.parent / "share" / "tessdata",
            Path("/usr/share/tesseract-ocr/5/tessdata"),
            Path("/usr/share/tesseract-ocr/4.00/tessdata"),
            Path("/opt/homebrew/share/tessdata"),
            Path("/usr/local/share/tessdata"),
        ]
        data = next((p for p in locations if (p / "eng.traineddata").is_file()), None)
        return TesseractRuntime(executable, data)
    raise RuntimeError(
        "OCR is needed but Tesseract is unavailable. Use an updated packaged build, "
        "or install Tesseract OCR with English language data and add it to PATH."
    )


def _render_page(data: bytes, index: int, path: Path) -> None:
    """Rasterize one original page, bounding both dimensions and total pixel memory."""
    import pypdfium2 as pdfium

    with PDF_LOCK:
        document = pdfium.PdfDocument(data)
        try:
            page = document[index]
            try:
                width, height = page.get_size()
                if not all(math.isfinite(v) and v > 0 for v in (width, height)):
                    raise ValueError("Invalid PDF page dimensions")
                scale = min(
                    300 / 72, 4000 / max(width, height), math.sqrt(10_000_000 / (width * height))
                )
                bitmap = page.render(scale=scale)
                try:
                    image = bitmap.to_pil().convert("RGB")
                    try:
                        image.save(path, dpi=(300, 300))
                    finally:
                        image.close()
                finally:
                    bitmap.close()
            finally:
                page.close()
        finally:
            document.close()


def _parse_tsv(output: str) -> OcrPage:
    """Keep every recognized word, including uncertain words, in original line order."""
    lines: dict[tuple[str, ...], list[str]] = {}
    scores = []
    reader = csv.DictReader(io.StringIO(output), delimiter="\t", quoting=csv.QUOTE_NONE)
    if not {"level", "text", "conf", "page_num", "block_num", "par_num", "line_num"} <= set(
        reader.fieldnames or ()
    ):
        raise ValueError("OCR returned an invalid TSV response")
    for row in reader:
        word = (row.get("text") or "").strip()
        if row.get("level") != "5" or not word:
            continue
        score = float(row["conf"])
        if not math.isfinite(score) or not 0 <= score <= 100:
            raise ValueError("OCR returned an invalid word confidence")
        key = tuple(row[k] for k in ("page_num", "block_num", "par_num", "line_num"))
        lines.setdefault(key, []).append(word)
        scores.append(score)
    return OcrPage(
        "\n".join(" ".join(words) for words in lines.values()),
        round(min(scores), 2) if scores else None,
    )


class PdfOcrReader:
    """Reuse runtime discovery while isolating each page in a cancellable subprocess."""

    def __init__(self) -> None:
        """Initialize lazy runtime discovery; text PDFs need no OCR installation."""
        self._runtime: TesseractRuntime | None = None

    def read_page(self, data: bytes, index: int, cancelled: Callable[[], bool]) -> OcrPage:
        """Recognize one page locally, allowing cancellation and a 60-second deadline."""
        if cancelled():
            raise OcrCancelled()
        if self._runtime is None:
            self._runtime = find_tesseract()
        runtime = self._runtime
        with tempfile.TemporaryDirectory(prefix="edf-ocr-") as temporary:
            image = Path(temporary) / "page.png"
            _render_page(data, index, image)
            if cancelled():
                raise OcrCancelled()
            command = [str(runtime.executable), str(image), "stdout", "-l", "eng", "--psm", "3"]
            if runtime.tessdata is not None:
                command.extend(["--tessdata-dir", str(runtime.tessdata)])
            command.extend(["-c", "tessedit_create_tsv=1"])
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            deadline = time.monotonic() + 60
            try:
                while True:
                    if cancelled():
                        raise OcrCancelled()
                    if time.monotonic() >= deadline:
                        raise TimeoutError("OCR exceeded the 60-second page limit")
                    try:
                        stdout, stderr = process.communicate(timeout=0.2)
                        break
                    except subprocess.TimeoutExpired:
                        continue
                if process.returncode:
                    raise RuntimeError(
                        "Tesseract failed: " + stderr.decode("utf-8", errors="replace")[-2000:]
                    )
                return _parse_tsv(stdout.decode("utf-8"))
            finally:
                if process.poll() is None:
                    process.kill()
                    process.communicate()
