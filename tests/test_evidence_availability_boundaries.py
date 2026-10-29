"""Evidence availability changes never silently substitute originals or reanalyse the case."""

import copy
import hashlib
from dataclasses import replace

import pytest

from edf_bill_fetcher.application.evidence_review import build_evidence_review
from edf_bill_fetcher.domain.analysis import analytical_fingerprint
from edf_bill_fetcher.domain.json_values import freeze_json
from tests.synthetic_analysis import snapshot
from tests.test_financial_identity_boundaries import bill


def with_state(accepted, state):
    extraction = replace(accepted.prepared.extraction, legacy_state=freeze_json(state))
    prepared = replace(accepted.prepared, extraction=extraction)
    return replace(
        accepted,
        prepared=prepared,
        input_fingerprint=analytical_fingerprint(prepared, accepted.methods),
    )


def source(review, identity):
    return next(
        item for entry in review.entries for item in entry.sources if item.source_id == identity
    )


@pytest.mark.parametrize("condition", ["valid", "changed", "missing", "no_path", "no_hash"])
def test_original_not_copied_into_pack_has_explicit_availability(tmp_path, condition):
    path = tmp_path / "original.pdf"
    path.write_bytes(b"Original source")
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    metadata = {
        "Original Path": str(path),
        "Content Hash": digest,
        "Attachment Name": "original.pdf",
    }
    if condition == "changed":
        path.write_bytes(b"Modified after extraction")
    elif condition == "missing":
        path.unlink()
    elif condition == "no_path":
        metadata.pop("Original Path")
    elif condition == "no_hash":
        metadata.pop("Content Hash")
    accepted = with_state(snapshot([bill()]), {"source_manifest": {"original": metadata}})
    before = copy.deepcopy(accepted)
    review = build_evidence_review(accepted, str(tmp_path))
    chosen = source(review, "original")
    if condition == "valid":
        assert chosen.path == str(path)
        assert "Verified original source" in chosen.status
    else:
        assert chosen.path is None
        assert (
            "changed" in chosen.status if condition == "changed" else "unavailable" in chosen.status
        )
    assert accepted == before


@pytest.mark.parametrize("condition", ["valid", "changed", "missing"])
def test_excluded_source_stays_bound_to_its_earlier_recorded_copy(tmp_path, condition):
    earlier = tmp_path / "excluded.pdf"
    earlier.write_bytes(b"Excluded original")
    digest = hashlib.sha256(earlier.read_bytes()).hexdigest()
    if condition == "changed":
        earlier.write_bytes(b"Changed earlier copy")
    elif condition == "missing":
        earlier.unlink()
    state = {
        "source_manifest": {
            "broken": None,
            "original": {"Attachment Name": "active.pdf", "Original Filename": "active.pdf"},
        },
        "review_excluded_source_manifest": {
            "original": {},
            "malformed": None,
            "excluded": {"Attachment Name": "excluded.pdf", "Content Hash": digest},
        },
        "review_source_paths": {"excluded": str(earlier)},
    }
    accepted = with_state(snapshot([bill()]), state)
    review = build_evidence_review(accepted, str(tmp_path))
    chosen = source(review, "excluded")
    if condition == "valid":
        assert chosen.path == str(earlier)
        assert "Verified excluded source" in chosen.status
    else:
        assert chosen.path is None
        assert "Excluded source unavailable" in chosen.status
    assert source(review, "original").title == "active.pdf"
    assert accepted.summary == review.snapshot.summary


def test_unmapped_legacy_filename_is_visible_without_inventing_a_source_path(tmp_path):
    row = bill(**{"Attachment Name": "unmapped.pdf"})
    row.pop("Source ID")
    review = build_evidence_review(snapshot([row]), str(tmp_path))
    chosen = source(review, "unmapped.pdf")
    assert chosen.path is None
    assert chosen.title == "unmapped.pdf"


@pytest.mark.parametrize("filtered", [{}, [None, {"Invoice #": "FILTERED"}]])
def test_old_filtered_metadata_cannot_create_fictitious_financial_rows(tmp_path, filtered):
    accepted = with_state(
        snapshot([bill()]), {"filtered_records": filtered, "error_log": "Preserved extraction note"}
    )
    review = build_evidence_review(accepted, str(tmp_path))
    excluded = [entry for entry in review.entries if entry.category == "Filtered out"]
    assert len(excluded) == (1 if isinstance(filtered, list) else 0)
    assert "Preserved extraction note" in review.notes
    assert review.snapshot.summary == accepted.summary
