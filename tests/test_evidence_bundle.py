"""Tests for the evidence-bundle save + index feature (Stream P5 / Task 8)."""

from __future__ import annotations

import os
from pathlib import Path

import pandas as pd

from edf_bill_fetcher.io.writers.evidence_bundle import (
    build_bundle_index,
    sanitise_filename,
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


def test_sanitise_filename_strips_illegal_chars() -> None:
    assert sanitise_filename("KI:123/4?.pdf") == "KI_123_4_.pdf"
    assert sanitise_filename("  T78701920034.pdf ") == "T78701920034.pdf"
    assert sanitise_filename("a  b.pdf") == "a_b.pdf"


def test_save_evidence_files_names_by_invoice_number(tmp_path: Path) -> None:
    df = pd.DataFrame(
        [
            {
                "Invoice #": "T78701920034",
                "Attachment Name": "671078701920_060241004086_20190416.pdf",
            },
            {
                "Invoice #": "KI-31105244-0014",
                "Attachment Name": "edf-invoice-KI-31105244-0014-1.pdf",
            },
        ]
    )
    src = tmp_path / "src"
    src.mkdir()
    for att in df["Attachment Name"]:
        (src / att).write_bytes(b"%PDF")
    out = save_evidence_files(
        df,
        {a: str(src / a) for a in df["Attachment Name"]},
        str(tmp_path / "ev"),
    )
    names = sorted(os.path.basename(p) for p in out.values())
    assert "T78701920034.pdf" in names
    assert "KI-31105244-0014.pdf" in names


def test_save_evidence_files_fallback_when_na_invoice(tmp_path: Path) -> None:
    df = pd.DataFrame([{"Invoice #": "N/A", "Attachment Name": "A1-invoice.pdf"}])
    src = tmp_path / "src"
    src.mkdir()
    (src / "A1-invoice.pdf").write_bytes(b"%PDF")
    out = save_evidence_files(
        df,
        {"A1-invoice.pdf": str(src / "A1-invoice.pdf")},
        str(tmp_path / "ev"),
    )
    assert "A1-invoice.pdf" in os.path.basename(next(iter(out.values())))


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


# End-to-end: engine.source_paths → save_evidence_files (spec §3.9, issue 8b).
# We append an import-free function-only block here; all module-imports live
# at the top of the file.

_PDF_B64 = (
    b"JVBERi0xLjMKJZOMi54gUmVwb3J0TGFiIEdlbmVyYXRlZCBQREYgZG9jdW1lbnQgKG9wZW5zb3VyY2Up"
    b"CjEgMCBvYmoKPDwKL0YxIDIgMCBSCj4+CmVuZG9iagoyIDAgb2JqCjw8Ci9CYXNlRm9udCAvSGVsdmV0aWNh"
    b"IC9FbmNvZGluZyAvV2luQW5zaUVuY29kaW5nIC9OYW1lIC9GMSAvU3VidHlwZSAvVHlwZTEgL1R5cGUg"
    b"L0ZvbnQKPj4KZW5kb2JqCjMgMCBvYmoKPDwKL0NvbnRlbnRzIDcgMCBSIC9NZWRpYUJveCBbIDAgMCA1"
    b"OTUuMjc1NiA4NDEuODg5OCBdIC9QYXJlbnQgNiAwIFIgL1Jlc291cmNlcyA8PAovRm9udCAxIDAgUiAv"
    b"UHJvY1NldCBbIC9QREYgL1RleHQgL0ltYWdlQiAvSW1hZ2VDIC9JbWFnZUkgXQo+PiAvUm90YXRlIDAg"
    b"IC9UcmFucyA8PAo+PiAKICAvVHlwZSAvUGFnZQo+PgplbmRvYmoKNCAwIG9iago8PAovUGFnZU1vZGUg"
    b"L1VzZU5vbmUgL1BhZ2VzIDYgMCBSIC9UeXBlIC9DYXRhbG9nCj4+CmVuZG9iago1IDAgb2JqCjw8Ci9B"
    b"dXRob3IgKGFub255bW91cykgL0NyZWF0aW9uRGF0ZSAoRDoyMDI2MDcyNTEzMTIyNSswMScwMCcpIC9D"
    b"cmVhdG9yIChhbm9ueW1vdXMpIC9LZXl3b3JkcyAoKSAvTW9kRGF0ZSAoRDoyMDI2MDcyNTEzMTIyNSsw"
    b"MScwMCcpIC9Qcm9kdWNlciAoUmVwb3J0TGFiIFBERiBMaWJyYXJ5IC0gXChvcGVuc291cmNlXCkpIAog"
    b"IC9TdWJqZWN0ICh1bnNwZWNpZmllZCkgL1RpdGxlICh1bnRpdGxlZCkgL1RyYXBwZWQgL0ZhbHNlCj4+"
    b"CmVuZG9iago2IDAgb2JqCjw8Ci9Db3VudCAxIC9LaWRzIFsgMyAwIFIgXSAvVHlwZSAvUGFnZXMKPj4K"
    b"ZW5kb2JqCjcgMCBvYmoKPDwKL0ZpbHRlciBbIC9BU0NJSTg1RGVjb2RlIC9GbGF0ZURlY29kZSBdIC9M"
    b"ZW5ndGggMTQ3Cj4+CnN0cmVhbQpHYXBARVltUz8lJ0xoYkZgPlIybG8wVC5fQlE4IlRuNCcqPTlx"
    b"W1o0UCFuZzIpYXVDUiUlImpIOVRBczAvVD1gY11jXCpgODosMSM0Jy06YS9xNmxbVHJaXnRoOUNXYiku"
    b"biZfQ0hNSWwuSVo/ZTBzUlQ4XFheXzVrc21yQzdyYyojNG8uSlVoYjBDSS5oM10pfj5lbmRzdHJlYW0K"
    b"ZW5kb2JqCnhyZWYKMCA4CjAwMDAwMDAwMDAgNjU1MzUgZiAKMDAwMDAwMDA2MSAwMDAwMCBuIAowMDAw"
    b"MDAwMDkyIDAwMDAwIG4gCjAwMDAwMDAxOTkgMDAwMDAgbiAKMDAwMDAwMDQwMiAwMDAwMCBuIAowMDAw"
    b"MDAwMDQ3MCAwMDAwMCBuIAowMDAwMDAwMDczMSAwMDAwMCBuIAowMDAwMDAwMDc5MCAwMDAwMCBuIAp0"
    b"cmFpbGVyCjw8Ci9JRCAKWzw2YTM1NTI2MjYxZTZlODlmN2UyZGI0ZmVlZjY3OWYwMj48NmEzNTUyNjI2"
    b"MWU2ZTg5ZjdlMmRiNGZlZWY2NzlmMDI+XQolIFJlcG9ydExhYiBnZW5lcmF0ZWQgUERGIGRvY3VtZW50"
    b"IC0tIGRpZ2VzdCAob3BlbnNvdXJjZSkKCi9JbmZvIDUgMCBSCi9Sb290IDQgMCBSCi9TaXplIDgKPj4K"
    b"c3RhcnR4cmVmCjEwMjcKJSVFT0YK"
)


def test_save_evidence_files_copies_when_engine_populated_source_paths(
    tmp_path: Path,
) -> None:
    """End-to-end: drive EvidenceEngine through one PDF, then call
    save_evidence_files(engine-records_df, engine.source_paths, dest) and
    confirm the file lands in dest. Spec §3.9 — this is the before-the-fix
    failure mode the user reported (evidence_files/ stayed empty)."""
    import base64

    from edf_bill_fetcher.collectors.engine import EvidenceEngine

    pdf_path = tmp_path / "edf-invoice-KI-1234-0001-3.pdf"
    pdf_path.write_bytes(base64.b64decode(_PDF_B64))

    eng = EvidenceEngine(config={"acc_num": ""}, update_ui_cb=lambda *a, **k: None)
    eng.process_pdf_file(
        str(pdf_path),
        "Local PDF Folder",
        pdf_path.name,
        "01/01/2024",
    )

    assert eng.source_paths, "engine.source_paths empty post-process — fix regressed"

    df = pd.DataFrame([{"Attachment Name": pdf_path.name}])
    dest = tmp_path / "output"
    saved = save_evidence_files(df, eng.source_paths, str(dest))

    assert pdf_path.name in saved, f"save_evidence_files did not copy: {saved}"
    assert (dest / pdf_path.name).exists()


def test_save_evidence_files_logs_ambiguous_attachment(tmp_path: Path) -> None:
    (tmp_path / "invoice.pdf").write_bytes(b"x")
    df = pd.DataFrame(
        [
            {"Invoice #": "T-1", "Attachment Name": "invoice.pdf"},
            {"Invoice #": "T-2", "Attachment Name": "invoice.pdf"},
        ]
    )
    logs: list[str] = []
    stats: dict[str, int] = {}
    out = save_evidence_files(
        df,
        {"invoice.pdf": str(tmp_path / "invoice.pdf")},
        str(tmp_path / "ev"),
        log=logs.append,
        stats=stats,
    )
    assert len(out) == 1
    assert stats["ambiguous"] == 1
    assert any("ambiguous" in m for m in logs)


def test_save_evidence_files_stats_counts_missing(tmp_path: Path) -> None:
    df = pd.DataFrame(
        [
            {"Invoice #": "T-1", "Attachment Name": "a.pdf"},
            {"Invoice #": "T-2", "Attachment Name": "missing.pdf"},
        ]
    )
    logs: list[str] = []
    stats: dict[str, int] = {}
    (tmp_path / "a.pdf").write_bytes(b"x")
    out = save_evidence_files(
        df,
        {"a.pdf": str(tmp_path / "a.pdf")},
        str(tmp_path / "ev"),
        log=logs.append,
        stats=stats,
    )
    assert len(out) == 1
    assert stats["saved"] == 1
    assert stats["missing"] == 1
    assert stats["ambiguous"] == 0
