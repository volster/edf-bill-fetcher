"""Shared PDFium synchronization for original previews and OCR rasterization."""

from threading import Lock

PDF_LOCK = Lock()  # PDFium requires serialized access, even across documents.
