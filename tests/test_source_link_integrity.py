"""Evidence links must never substitute an unrelated same-named original."""

import hashlib
import json
from pathlib import Path

import pytest

from edf_bill_fetcher.io.reporters.source_links import (
    SourceLink,
    record_source_link,
    source_links,
)


def entry(root, source_id, path, name="bill.pdf"):
    target = root / path
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(source_id.encode())
    return {
        "source_id": source_id,
        "attachment_name": name,
        "saved_path": path,
        "sha256": hashlib.sha256(target.read_bytes()).hexdigest(),
        "status": "saved",
    }


def manifest(root, entries):
    (root / "evidence_manifest.json").write_text(json.dumps({"files": entries}), encoding="utf-8")


@pytest.mark.parametrize("failure", ["missing", "changed", "unsaved"])
def test_unavailable_original_never_links_to_another_same_named_file(tmp_path, failure):
    first = entry(tmp_path, "first", "evidence_files/first.pdf")
    second = entry(tmp_path, "second", "evidence_files/second.pdf")
    if failure == "missing":
        (tmp_path / first["saved_path"]).unlink()
    elif failure == "changed":
        (tmp_path / first["saved_path"]).write_bytes(b"Changed original")
    else:
        first["status"] = "missing"
    manifest(tmp_path, [first, second])

    links = source_links(str(tmp_path))

    assert record_source_link({"Source ID": "first", "Attachment Name": "bill.pdf"}, links) is None
    assert record_source_link({"Attachment Name": "bill.pdf"}, links) is None
    assert record_source_link({"Source ID": "second"}, links) == SourceLink(
        "evidence_files/second.pdf", "Open source"
    )


def test_explicit_source_identity_prevents_filename_fallback(tmp_path):
    manifest(tmp_path, [entry(tmp_path, "other", "evidence_files/other.pdf")])
    links = source_links(str(tmp_path))
    assert (
        record_source_link({"Source ID": "unknown", "Attachment Name": "bill.pdf"}, links) is None
    )
    assert record_source_link({"Attachment Name": "bill.pdf"}, links) == links["other"]


def test_repeated_invoice_aliases_of_one_original_remain_linkable(tmp_path):
    original = entry(tmp_path, "merged", "evidence_files/merged bills.pdf")
    manifest(tmp_path, [original, {**original, "invoice": "INV-2"}])
    links = source_links(str(tmp_path))
    expected = SourceLink("evidence_files/merged%20bills.pdf", "Open source")
    assert links == {"merged": expected, "bill.pdf": expected}
    assert record_source_link({"Source ID": "merged", "Page From": 2}, links) == SourceLink(
        "evidence_files/merged%20bills.pdf#page=2", "Open source"
    )


def test_conflicting_paths_for_one_source_are_not_resolved_by_file_availability(tmp_path):
    original = entry(tmp_path, "shared-id", "evidence_files/first.pdf")
    manifest(tmp_path, [original, {**original, "saved_path": "evidence_files/missing.pdf"}])
    assert source_links(str(tmp_path)) == {}


@pytest.mark.parametrize("payload", ["{", "null", "[]", "{}", '{"files": {}}'])
def test_invalid_manifest_has_no_source_links(tmp_path, payload):
    (tmp_path / "evidence_manifest.json").write_text(payload, encoding="utf-8")
    assert source_links(str(tmp_path)) == {}


def test_absent_bundle_and_manifest_have_no_links(tmp_path):
    assert source_links(None) == {}
    assert source_links(str(tmp_path)) == {}


@pytest.mark.parametrize(
    "path",
    [None, 7, "../outside.pdf", "/outside.pdf", "C:/outside.pdf", "bad\\file.pdf", "bad\x00.pdf"],
)
def test_unsafe_manifest_path_does_not_hide_other_valid_sources(tmp_path, path):
    valid = entry(tmp_path, "valid", "evidence_files/valid.pdf", "valid.pdf")
    manifest(tmp_path, [None, {"status": "saved", "saved_path": path}, valid])
    links = source_links(str(tmp_path))
    assert links == {
        "valid": SourceLink("evidence_files/valid.pdf", "Open source"),
        "valid.pdf": SourceLink("evidence_files/valid.pdf", "Open source"),
    }


@pytest.mark.parametrize("operation", ["resolve", "read_bytes"])
def test_one_unreadable_source_does_not_remove_other_links(tmp_path, monkeypatch, operation):
    broken = entry(tmp_path, "broken", "evidence_files/broken.pdf", "broken.pdf")
    valid = entry(tmp_path, "valid", "evidence_files/valid.pdf", "valid.pdf")
    manifest(tmp_path, [broken, valid])
    original = getattr(Path, operation)

    def unavailable(path, *args, **kwargs):
        if path.name == "broken.pdf":
            raise PermissionError("Simulated unavailable original")
        return original(path, *args, **kwargs)

    monkeypatch.setattr(Path, operation, unavailable)
    links = source_links(str(tmp_path))
    assert "broken" not in links
    assert links["valid"].href == "evidence_files/valid.pdf"


def test_symlink_outside_bundle_is_never_linked(tmp_path):
    bundle = tmp_path / "bundle"
    bundle.mkdir()
    outside = tmp_path / "outside.pdf"
    outside.write_bytes(b"Outside bundle")
    try:
        (bundle / "escape.pdf").symlink_to(outside)
    except OSError:
        pytest.skip("Creating symlinks requires privileges on this platform")
    manifest(bundle, [{"status": "saved", "source_id": "escape", "saved_path": "escape.pdf"}])
    assert source_links(str(bundle)) == {}


@pytest.mark.parametrize(
    "page", [0, -1, True, False, 1.5, float("nan"), float("inf"), "2", None, 1000000]
)
def test_invalid_page_keeps_verified_file_without_guessing_a_page(page):
    base = SourceLink("evidence_files/bill.pdf", "Open source")
    assert record_source_link({"Source ID": "known", "Page From": page}, {"known": base}) == base


@pytest.mark.parametrize("page", [2, 2.0])
def test_integer_page_is_preserved(page):
    base = SourceLink("evidence_files/bill.pdf", "Open source")
    assert record_source_link(
        {"Source ID": "known", "Page From": page}, {"known": base}
    ) == SourceLink("evidence_files/bill.pdf#page=2", "Open source")


def test_no_mapped_original_means_no_link():
    assert record_source_link({"Source ID": "unknown", "Attachment Name": "bill.pdf"}, {}) is None
