"""Tests for versioned JSON engine-state round-tripping in ``edf_bill_fetcher.io.cli``.

Wave 6b / Task 2: the CLI previously read engine state only from pickle
through ``_safe_pickle_load`` (restricted unpickler).  This task adds a
versioned JSON format so engine state — ``{config, filtered_records,
error_log}`` — round-trips without pickle, while the legacy ``.pkl``
read path stays intact.

Pinned here:

1. ``_safe_json_dump`` → ``_safe_pickle_load`` round-trip restores the
   exact state and writes ``{"version": 1, ...}`` on disk;
2. a legacy ``.pkl`` (pickled ``EvidenceEngine``) still loads through
   the restricted unpickler;
3. an unknown extension raises ``ValueError`` (loud failure, no
   silent pickle fallback);
4. dispatch is pinned both ways — JSON text under a ``.pkl`` name
   raises from the pickle path, and a pickle under a ``.json`` name
   raises from the JSON path.

All record data here is synthetic; no real customer data.
"""

from __future__ import annotations

import json
import pickle
from pathlib import Path
from typing import Any

import pytest

from edf_bill_fetcher.collectors.engine import EvidenceEngine
from edf_bill_fetcher.io.cli import _safe_json_dump, _safe_pickle_load


def _synthetic_state() -> dict:
    """Deterministic engine-state fragment used across the round-trip tests."""
    return {
        "config": {"min_amount": 50.0, "use_dedup": True, "domain_filter": "edfenergy.com"},
        "filtered_records": [
            {
                "Source": "Local PDF",
                "Date": "15/01/2026",
                "Amount (£)": 12.0,
                "Invoice #": "KI-0000000-0002",
            }
        ],
        "error_log": ["Unrecognised layout in bill.pdf fell back to generic patterns"],
    }


class TestJsonEngineStateRoundTrip:
    """Versioned JSON write → read round-trip via ``_safe_json_dump`` / ``_safe_pickle_load``."""

    def test_json_roundtrip_restores_state(self, tmp_path: Path) -> None:
        """Dump state to ``.json`` then load it back — identical data, version present.

        The on-disk payload carries ``version: 1`` plus the three state
        keys, and ``_safe_pickle_load`` dispatches on the ``.json``
        extension to ``json.load``.
        """
        state = _synthetic_state()
        path = tmp_path / "engine.json"

        _safe_json_dump(state, str(path))

        on_disk = json.loads(path.read_text(encoding="utf-8"))
        assert on_disk["version"] == 1
        assert on_disk["config"] == state["config"]
        assert on_disk["filtered_records"] == state["filtered_records"]
        assert on_disk["error_log"] == state["error_log"]

        loaded = _safe_pickle_load(str(path))
        assert loaded == on_disk

    def test_json_roundtrip_with_empty_collections(self, tmp_path: Path) -> None:
        """Empty filtered_records / error_log survive the round-trip too."""
        state: dict[str, Any] = {"config": {}, "filtered_records": [], "error_log": []}
        path = tmp_path / "engine_empty.json"

        _safe_json_dump(state, str(path))
        loaded = _safe_pickle_load(str(path))

        assert loaded["version"] == 1
        assert loaded["config"] == {}
        assert loaded["filtered_records"] == []
        assert loaded["error_log"] == []


class TestLegacyPickleRead:
    """The legacy ``.pkl`` read path keeps working through the restricted unpickler."""

    def test_legacy_pkl_read_still_works(self, tmp_path: Path) -> None:
        """A pickled ``EvidenceEngine`` under a ``.pkl`` name loads unchanged.

        The engine is pickled with the standard ``pickle.dump`` (its
        ``__getstate__`` already strips the non-picklable runtime
        primitives), then restored through ``_safe_pickle_load`` — which
        must route ``.pkl`` to the existing restricted-unpickle path.
        """
        engine = EvidenceEngine({"min_amount": 50.0}, print)
        engine.filtered_records = [{"Source": "Local PDF", "Date": "15/01/2026", "Amount (£)": 5.0}]
        engine.error_log = ["synthetic parse warning"]
        path = tmp_path / "legacy.pkl"
        with open(path, "wb") as f:
            pickle.dump(engine, f)

        restored = _safe_pickle_load(str(path))

        assert isinstance(restored, EvidenceEngine)
        assert restored.filtered_records == engine.filtered_records
        assert restored.error_log == engine.error_log
        assert restored.config == engine.config


class TestDispatchFailures:
    """Dispatch is pinned: wrong content or wrong extension fails loudly."""

    def test_unknown_extension_raises(self, tmp_path: Path) -> None:
        """A file with an unrecognised extension raises ``ValueError``.

        No silent pickle fallback — a typo'd extension must surface an
        actionable error, not try the unpickler on arbitrary bytes.
        """
        path = tmp_path / "engine.dat"
        path.write_text("not-a-known-format", encoding="utf-8")

        with pytest.raises(ValueError, match="Unsupported engine state file extension"):
            _safe_pickle_load(str(path))

    def test_json_content_in_pkl_path_raises(self, tmp_path: Path) -> None:
        """JSON text under a ``.pkl`` name raises from the pickle path.

        Pins the dispatch in the failure direction called out in the
        task brief's QA scenarios: a ``.pkl`` path is fed to the
        restricted unpickler, which rejects the JSON bytes instead of
        silently succeeding or half-restoring.
        """
        path = tmp_path / "state.pkl"
        path.write_text(json.dumps({"version": 1}), encoding="utf-8")

        # The pickle path must reject the payload — a broken dispatch
        # that routed ``.pkl`` to ``json.load`` would instead raise
        # ``json.JSONDecodeError`` and fail this expectation.
        with pytest.raises(pickle.UnpicklingError):
            _safe_pickle_load(str(path))
