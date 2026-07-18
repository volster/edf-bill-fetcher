"""Tests for the evidence-bundle save + index feature (Stream P5 / Task 8)."""

from __future__ import annotations

import os
from pathlib import Path

import pandas as pd

from evidence_bundle import (
    build_bundle_index,
    save_evidence_files,
)


def _make_files(tmp: str) -> dict[str, str]:
    paths = {}
    for name in ("A1-invoice.pdf", "B2-letter.pdf", "edf-invoice-KI-1234-0001-3.pdf"):
        p = os.path.join(tmp, name)
        with open(p, "w") as fh:
            fh.write(f"hello {name}")
        paths[name] = p
    return paths


def test_save_evidence_files_copies_all_into_dest(tmp_path: Path) -> None:
    src = tmp_path / "src"
    src.mkdir()
    src_paths = _make_files(str(src))
    df = pd.DataFrame(
        [
            {"Attachment Name": "A1-invoice.pdf"},
            {"Attachment Name": "B2-letter.pdf"},
            {"Attachment Name": "edf-invoice-KI-1234-0001-3.pdf"},
        ]
    )
    dest = tmp_path / "out"
    out = save_evidence_files(df, src_paths, str(dest))
    assert len(out) == 3
    for name, dest_path in out.items():
        assert os.path.exists(dest_path)
        # Names match what the source was.
        assert os.path.basename(dest_path) == name


def test_save_evidence_files_disambiguates_collisions(tmp_path: Path) -> None:
    src = tmp_path / "src"
    src.mkdir()
    os.makedirs(str(src), exist_ok=True)
    a = os.path.join(str(src), "twin.pdf")
    with open(a, "w") as fh:
        fh.write("first")
    b = os.path.join(str(src), "twin.pdf")  # same path -- same name collision
    src_paths = {"twin.pdf": a, "twin-2.pdf": b}
    # Two distinct Attachment Names that both resolve to "twin.pdf" — exercise
    # the dedup path by re-using the same name twice.
    df = pd.DataFrame(
        [
            {"Attachment Name": "twin.pdf"},
            {"Attachment Name": "twin.pdf"},  # same attachment name; should hit dedup
        ]
    )
    dest = tmp_path / "out"
    out = save_evidence_files(df, src_paths, str(dest))
    assert len(out) == 1
    assert os.path.exists(out["twin.pdf"])


def test_save_evidence_files_skips_missing_sources(tmp_path: Path) -> None:
    src = tmp_path / "src"
    src.mkdir()
    src_paths = _make_files(str(src))
    # 'missing.pdf' isn't in src_paths.
    df = pd.DataFrame(
        [
            {"Attachment Name": "A1-invoice.pdf"},
            {"Attachment Name": "missing.pdf"},
            {"Attachment Name": "N/A"},
        ]
    )
    dest = tmp_path / "out"
    logs: list[str] = []
    out = save_evidence_files(df, src_paths, str(dest), log=lambda m: logs.append(m))
    assert len(out) == 1
    assert "A1-invoice.pdf" in out
    assert "missing" in " ".join(logs)
    # N/A_ATTACHMENT_NAME is skipped silently (no missing-source log).
    assert all("N/A" not in (line or "") for line in logs)


def test_save_evidence_files_creates_dest_dir(tmp_path: Path) -> None:
    src = tmp_path / "src"
    src.mkdir()
    src_paths = _make_files(str(src))
    df = pd.DataFrame([{"Attachment Name": "A1-invoice.pdf"}])
    dest = tmp_path / "nested" / "dest"  # does not exist yet
    out = save_evidence_files(df, src_paths, str(dest))
    assert os.path.isdir(str(dest))
    assert os.path.exists(out["A1-invoice.pdf"])


def test_build_bundle_index_creates_docx(tmp_path: Path) -> None:
    src = tmp_path / "src"
    src.mkdir()
    src_paths = _make_files(str(src))
    df = pd.DataFrame(
        [
            {
                "Attachment Name": "A1-invoice.pdf",
                "Source": "Local PDF Folder",
                "Date": "14/05/2024",
                "Entry Type": "New Bill",
            },
            {
                "Attachment Name": "B2-letter.pdf",
                "Source": "PST PDF Attachment",
                "Date": "15/05/2024",
                "Entry Type": "Letter",
            },
            {
                "Attachment Name": "edf-invoice-KI-1234-0001-3.pdf",
                "Source": "Local PDF Folder",
                "Date": "16/05/2024",
                "Entry Type": "New Bill",
            },
        ]
    )
    out = save_evidence_files(df, src_paths, str(tmp_path / "ev"))
    docx_path = tmp_path / "evidence_index.docx"
    build_bundle_index(df, out, str(docx_path))
    assert os.path.exists(str(docx_path))
    # File must be a valid .docx (zip with the right header).
    with open(str(docx_path), "rb") as fh:
        head = fh.read(2)
    assert head == b"PK"


def test_build_bundle_index_themed_sections(tmp_path: Path) -> None:
    src = tmp_path / "src"
    src.mkdir()
    # Files matching each of: A prefix (Ombudsman), KI invoice (default D),
    # Meter-Read-History fingerprint (E).
    files = {
        "A1-ombudsman-letter.pdf": "ombud",
        "edf-invoice-KI-1234-0001-3.pdf": "inv",
        "Meter-Read-History.pdf": "mrh",
    }
    src_paths = {}
    for name, body in files.items():
        p = os.path.join(str(src), name)
        with open(p, "w") as fh:
            fh.write(body)
        src_paths[name] = p
    df = pd.DataFrame(
        [
            {
                "Attachment Name": "A1-ombudsman-letter.pdf",
                "Source": "Local PDF Folder",
                "Date": "14/05/2024",
                "Entry Type": "Letter",
                "Details": "",
            },
            {
                "Attachment Name": "edf-invoice-KI-1234-0001-3.pdf",
                "Source": "Local PDF Folder",
                "Date": "14/05/2024",
                "Entry Type": "New Bill",
                "Details": "",
            },
            {
                "Attachment Name": "Meter-Read-History.pdf",
                "Source": "Local PDF Folder",
                "Date": "14/05/2024",
                "Entry Type": "Statement Reconciliation",
                "Details": "Meter readings",
            },
        ]
    )
    out = save_evidence_files(df, src_paths, str(tmp_path / "ev"))
    docx_path = tmp_path / "evidence_index.docx"
    build_bundle_index(df, out, str(docx_path))
    import docx

    doc = docx.Document(str(docx_path))
    text = "\n".join(p.text for p in doc.paragraphs)
    # Ombudsman section header (A1 prefix mapped to A - Ombudsman).
    assert "Ombudsman" in text or "A —" in text
    # Invoices section (KI prefix or default).
    assert "Invoice" in text or "D —" in text
    # Meter-Read-History filename fingerprinted into E - Meter Readings.
    assert "Meter" in text or "E —" in text
