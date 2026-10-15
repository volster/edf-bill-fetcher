"""Ship the native OCR executable, dependencies and English recognition data."""

import os

from edf_bill_fetcher.io.adapters.ocr import find_tesseract

runtime = find_tesseract()
if runtime.tessdata is None:
    raise RuntimeError("Cannot package OCR without eng.traineddata")
binaries = [(str(runtime.executable), "ocr")]
if os.name == "nt":
    binaries.extend((str(path), "ocr") for path in runtime.executable.parent.glob("*.dll"))
datas = [(str(runtime.tessdata / "eng.traineddata"), "ocr/tessdata")]
