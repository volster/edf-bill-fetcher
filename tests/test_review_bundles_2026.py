"""Review regressions for authoritative source identities and portable mappings."""

import hashlib
from datetime import date, datetime, timezone

import pandas as pd
import pytest

from edf_bill_fetcher.domain.json_values import freeze_json
from edf_bill_fetcher.domain.requests import RunRequest
from edf_bill_fetcher.io.persistence.result_codec import (
    decode_result,
    encode_result,
    migrate_legacy,
)
from edf_bill_fetcher.io.writers.evidence_bundle import (
    build_bundle_manifest,
    resolve_bundle_paths,
    save_evidence_files,
)


def _source(root, name, data):
    path = root / name
    path.write_bytes(data)
    return {
        "Attachment Name": "same.pdf",
        "Saved Path": name,
        "Content Hash": hashlib.sha256(data).hexdigest(),
    }


@pytest.mark.parametrize("second_status", ["valid", "changed", "missing"])
def test_ambiguous_name_never_selects_one_occurrence(tmp_path, second_status):
    first = _source(tmp_path, "first.pdf", b"first original")
    second = _source(tmp_path, "second.pdf", b"second original")
    if second_status == "changed":
        (tmp_path / "second.pdf").write_bytes(b"changed contents")
    elif second_status == "missing":
        (tmp_path / "second.pdf").unlink()
    paths, notes = resolve_bundle_paths(
        {"source_manifest": {"first-id": first, "second-id": second}}, str(tmp_path)
    )
    assert paths["first-id"] == str(tmp_path / "first.pdf")
    if second_status == "valid":
        assert paths["second-id"] == str(tmp_path / "second.pdf")
    else:
        assert "second-id" not in paths
    assert "same.pdf" not in paths
    assert any("ambiguous" in note.casefold() for note in notes)


def test_display_alias_cannot_overwrite_another_authoritative_id(tmp_path):
    first = _source(tmp_path, "first.pdf", b"first original")
    second = _source(tmp_path, "second.pdf", b"second original")
    first["Attachment Name"], second["Attachment Name"] = "second-id", "first-id"
    paths, _notes = resolve_bundle_paths(
        {"source_manifest": {"first-id": first, "second-id": second}}, str(tmp_path)
    )
    assert paths == {
        "first-id": str(tmp_path / "first.pdf"),
        "second-id": str(tmp_path / "second.pdf"),
    }


@pytest.mark.parametrize("include_resolved_row", [False, True])
def test_missing_source_id_cannot_fall_back_to_unrelated_display_alias(
    tmp_path, include_resolved_row
):
    original = tmp_path / "original.pdf"
    original.write_bytes(b"unrelated original")
    records = [{"Attachment Name": "same.pdf", "Source ID": "missing-id"}]
    if include_resolved_row:
        records.insert(0, {"Attachment Name": "same.pdf", "Source ID": "present-id"})
    stats: dict[str, int] = {}
    saved = save_evidence_files(
        pd.DataFrame(records),
        {"same.pdf": str(original), "present-id": str(original)},
        str(tmp_path / "pack"),
        stats=stats,
    )
    assert saved == {}
    assert stats["saved"] == 0
    assert stats["missing"] == 1
    assert list((tmp_path / "pack").iterdir()) == []
    assert original.read_bytes() == b"unrelated original"


def test_invalid_saved_path_is_reported_without_aborting_review(tmp_path):
    paths, notes = resolve_bundle_paths(
        {"source_manifest": {"id": {"Saved Path": "bad\x00.pdf"}}}, str(tmp_path)
    )
    assert paths == {}
    assert len(notes) == 1
    assert "unavailable" in notes[0].casefold()


@pytest.mark.parametrize("relative", ["../outside.pdf", "..\\outside.pdf", "bad\x00.pdf"])
def test_legacy_relative_evidence_cannot_escape_or_abort_review(tmp_path, relative):
    bundle = tmp_path / "pack"
    bundle.mkdir()
    (tmp_path / "outside.pdf").write_bytes(b"outside original")
    paths, notes = resolve_bundle_paths(
        {"saved_evidence_paths": {"invoice.pdf": relative}}, str(bundle)
    )
    assert paths == {}
    assert len(notes) == 1


def test_legacy_absolute_path_remains_supported_without_a_manifest(tmp_path):
    source = tmp_path / "original.pdf"
    source.write_bytes(b"original")
    bundle = tmp_path / "pack"
    bundle.mkdir()
    paths, notes = resolve_bundle_paths(
        {"saved_evidence_paths": {"invoice.pdf": str(source)}}, str(bundle)
    )
    assert paths == {"invoice.pdf": str(source)}
    assert notes == ()


@pytest.mark.parametrize("legacy", [False, True])
def test_relative_mapping_cannot_follow_a_symlink_outside_the_pack(tmp_path, legacy):
    bundle = tmp_path / "pack"
    bundle.mkdir()
    source = tmp_path / "original.pdf"
    source.write_bytes(b"outside original")
    try:
        (bundle / "link.pdf").symlink_to(source)
    except (OSError, NotImplementedError):
        pytest.skip("Creating symlinks is unavailable on this platform")
    state = (
        {"saved_evidence_paths": {"original.pdf": "link.pdf"}}
        if legacy
        else {"source_manifest": {"id": {"Saved Path": "link.pdf"}}}
    )
    paths, notes = resolve_bundle_paths(state, str(bundle))
    assert paths == {}
    assert len(notes) == 1


@pytest.mark.parametrize("invoice", ["", "KI-10000000-0001"])
def test_unicode_original_filename_is_bounded_and_preserved_in_the_manifest(tmp_path, invoice):
    source = tmp_path / "original.pdf"
    source.write_bytes(b"original evidence")
    # This metadata can come from a message attachment even when the temporary
    # extracted source uses a short generated filename on the local filesystem.
    original_name = "💷" * 100 + "." + "証" * 20
    frame = pd.DataFrame(
        [
            {
                "Attachment Name": "source.pdf",
                "Original Filename": original_name,
                "Invoice #": invoice,
            }
        ]
    )
    saved = save_evidence_files(
        frame, {"source.pdf": str(source)}, str(tmp_path / "evidence_files")
    )
    assert saved
    copied = next((tmp_path / "evidence_files").iterdir())
    assert copied.read_bytes() == b"original evidence"
    assert len(copied.name.encode("utf-8")) <= 200
    manifest = build_bundle_manifest(frame, saved, str(tmp_path))
    assert manifest["files"][0]["original_filename"] == original_name
    assert manifest["files"][0]["saved_path"] == "evidence_files/" + copied.name


def test_bundle_manifest_omits_reversed_invoice_page_ranges(tmp_path):
    source = tmp_path / "original.pdf"
    source.write_bytes(b"original")
    manifest = build_bundle_manifest(
        pd.DataFrame(
            [
                {"Attachment Name": source.name, "Page From": 4, "Page To": 2},
                {"Attachment Name": source.name, "Page From": 2, "Page To": 4},
            ]
        ),
        {source.name: str(source)},
        str(tmp_path),
    )
    assert manifest["files"][0]["page_ranges"] == [[2, 4]]


@pytest.fixture(scope="module")
def accepted_result():
    request = RunRequest(
        "bundle-review",
        datetime(2026, 1, 31, tzinfo=timezone.utc),
        date(2026, 1, 31),
        (),
        freeze_json({"as_of": "2026-01-31"}),
    )
    return migrate_legacy(
        [
            {
                "Date": "01/01/2026",
                "Account": "A-100",
                "Entry Type": "New Bill",
                "Amount Kind": "Balance",
                "Amount (£)": 100,
                "Period Charge (£)": 60,
                "Invoice #": "B1",
            }
        ],
        request,
    )


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("frames", {"evidence": None}),
        ("frames", {}),
        ("analyses", None),
        ("analyses", {}),
        ("event_source_indexes", None),
        ("events", [{}]),
        ("matches", [{}]),
        ("position", None),
        ("domination", []),
        ("timeline", {}),
        ("union_total", "bad"),
    ],
)
def test_invalid_precomputed_payload_is_rejected_when_opening_a_saved_result(
    accepted_result, field, value
):
    payload = encode_result(accepted_result, "review-result")
    payload["analysis"]["models"]["sap"]["payload"][field] = value
    with pytest.raises(ValueError):
        decode_result(payload)


@pytest.mark.parametrize("availability", [[], {}])
def test_model_availability_has_a_controlled_validation_error(accepted_result, availability):
    payload = encode_result(accepted_result, "review-result")
    payload["analysis"]["models"]["payment"]["availability"] = availability
    with pytest.raises(ValueError, match="availability"):
        decode_result(payload)


@pytest.mark.parametrize("container", ["summaries", "details"])
def test_reconciliation_requires_row_arrays(accepted_result, container):
    payload = encode_result(accepted_result, "review-result")
    payload["analysis"]["models"]["reconciliation"]["payload"][container] = None
    with pytest.raises(ValueError, match="reconciliation"):
        decode_result(payload)
