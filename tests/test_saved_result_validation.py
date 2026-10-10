"""Reject corrupt accepted results and require explicit legacy/reanalysis context."""

import copy
import json
from datetime import date, datetime, timezone
from pathlib import Path

import pytest

from edf_bill_fetcher.domain.analysis import AnalysisSnapshot
from edf_bill_fetcher.domain.json_values import freeze_json, object_value
from edf_bill_fetcher.domain.requests import RunRequest
from edf_bill_fetcher.io.persistence.result_codec import (
    decode_result,
    encode_result,
    migrate_legacy,
)
from edf_bill_fetcher.io.persistence.result_store import load_result, save_result
from edf_bill_fetcher.models.events import SapBackBillingEvent
from edf_bill_fetcher.models.snapshot_adapter import from_snapshot_json, to_snapshot_json


def request(run_id="validation", as_of="2024-04-01"):
    return RunRequest(
        run_id,
        datetime(2024, 4, 1, tzinfo=timezone.utc),
        date.fromisoformat(as_of),
        (),
        freeze_json({"as_of": as_of, "analysis_min": 0, "use_dedup": True}),
    )


@pytest.fixture(scope="module")
def snapshot():
    case = json.loads(
        (Path(__file__).parent / "fixtures/financial_contracts.json").read_text(encoding="utf-8")
    )["cases"][0]
    return migrate_legacy(case["records"], request())


@pytest.fixture
def payload(snapshot):
    return encode_result(snapshot, "validation-result")


@pytest.mark.parametrize(
    ("path", "value"),
    [
        (("schema_version",), True),
        (("schema_version",), 4),
        (("result_id",), ""),
        (("request", "sources"), {}),
        (("request", "sources"), [{"kind": "pdf_dir", "path": ""}]),
        (("request", "created_at"), "2024-04-01T00:00:00"),
        (("request", "as_of"), "20240401"),
        (("request", "config", "as_of"), "2024-04-02"),
        (("evidence", "observations"), {}),
        (("evidence", "observations", 0, "ordinal"), True),
        (("evidence", "observations", 0, "observation_id"), "a" * 64),
        (("evidence", "observations", 0, "raw", "Amount (£)"), 999),
        (("evidence", "prepared_records"), [1]),
        (("evidence", "duplicate_records"), [{"_matches_kept_idx": 1}]),
        (("evidence", "effective_config", "as_of"), "2024-04-02"),
        (("analysis", "summary", "charges"), 100),
        (("analysis", "summary", "charges"), "100.0"),
        (("analysis", "summary", "charges"), "0100.00"),
        (("analysis", "summary", "charges"), "-0.00"),
        (("analysis", "summary", "mean_daily_charge"), True),
        (("analysis", "input_fingerprint"), "a" * 64),
        (("analysis", "models", "payment", "availability"), "invented"),
        (("analysis", "models", "payment", "reason_code"), ""),
        (("analysis", "models", "payment", "payload", "name"), "ForecastResult"),
        (("analysis", "models", "payment", "payload", "fields", "count"), -1),
        (("analysis", "models", "payment", "payload", "fields", "outlier_dates"), "wrong"),
        (("analysis", "models", "compensation", "payload"), [1]),
        (("analysis", "models", "reconciliation", "payload"), {}),
        (("analysis", "models", "sap", "payload"), {}),
    ],
)
def test_corrupt_result_is_rejected(payload: dict, path: tuple, value: object) -> None:
    parent = payload
    for key in path[:-1]:
        parent = parent[key]
    parent[path[-1]] = value
    with pytest.raises(ValueError):
        decode_result(payload)


def test_partial_model_requires_reason(payload):
    entry = payload["analysis"]["models"]["payment"]
    entry.update(availability="partial", reason_code=None)
    with pytest.raises(ValueError, match="needs a reason"):
        decode_result(payload)


def test_unknown_schema_fields_are_not_silently_ignored(payload):
    payload["evidence"]["invented"] = []
    with pytest.raises(ValueError, match="Expected object fields"):
        decode_result(payload)


@pytest.mark.parametrize("mode", ["negative", "bool", "missing", "different_row", "wrong_length"])
def test_sap_events_must_reference_exact_preserved_rows(payload: dict, mode: str) -> None:
    row = {"Amount": 100, "Document": "synthetic"}
    event = SapBackBillingEvent("CLEAR-1", None, "synthetic", rows=[row])
    sap = payload["analysis"]["models"]["sap"]["payload"]
    sap["events"] = [to_snapshot_json(event)]
    sap["event_source_indexes"] = [[0]]
    payload["evidence"]["legacy_state"]["sap_financial_rows"] = [copy.deepcopy(row)]
    if mode in {"negative", "bool", "missing"}:
        sap["event_source_indexes"] = [[{"negative": -1, "bool": True, "missing": 1}[mode]]]
    elif mode == "different_row":
        sap["events"][0]["fields"]["rows"][0]["Amount"] = 101
    else:
        sap["event_source_indexes"] = []
    with pytest.raises(ValueError, match="SAP"):
        decode_result(payload)


@pytest.mark.parametrize(
    ("entity", "side", "index"),
    [
        ("invented", "sap", 0),
        ("Financial", "other", 0),
        ("Meter Read", "sap", True),
        ("Contract", "edf", -1),
        ("Contract", "sap", 0),
        ("Contract", "edf", 99),
        ("Meter Read", "sap", 0),
        ("Financial", "sap", 0),
        ("Financial", "edf", 99),
    ],
)
def test_reconciliation_cannot_claim_nonexistent_evidence(
    payload: dict, entity: str, side: str, index: object
) -> None:
    payload["analysis"]["models"]["reconciliation"]["payload"]["details"] = [
        {"entity": entity, "side": side, "index": index, "values": {}}
    ]
    with pytest.raises(ValueError, match="[Rr]econciliation"):
        decode_result(payload)


@pytest.mark.parametrize(
    "frame",
    [
        {"type": "frame"},
        {"type": "frame", "columns": "x", "rows": [], "index": []},
        {"type": "frame", "columns": ["x"], "rows": [[1, 2]], "index": [0]},
        {"type": "frame", "columns": ["x"], "rows": [[1]], "index": []},
        {"type": "frame", "columns": ["x", "x"], "rows": [], "index": []},
        {"type": "frame", "columns": [1], "rows": [], "index": []},
    ],
)
def test_malformed_stored_frame_is_rejected(frame: dict) -> None:
    with pytest.raises(ValueError, match="[Ff]rame"):
        from_snapshot_json(frame)


@pytest.mark.parametrize(
    "legacy",
    [
        None,
        {},
        {"version": True, "records": []},
        {"version": 3, "records": []},
        {"records": [1]},
        {"records": [], "config": []},
    ],
)
def test_invalid_legacy_input_cannot_be_migrated(legacy: object) -> None:
    with pytest.raises(ValueError):
        migrate_legacy(legacy, request())


def test_legacy_without_date_context_requires_explicit_request(tmp_path: Path) -> None:
    destination = tmp_path / "legacy.json"
    destination.write_text('[{"Date":"31/01/2024","Amount (£)":100}]', encoding="utf-8")
    with pytest.raises(ValueError, match="supply legacy_request"):
        load_result(destination)
    result = load_result(destination, legacy_request=request())
    notes = object_value(result.prepared.extraction.legacy_state)["migration_notes"]
    assert any("historical methodology unavailable" in note for note in notes)
    assert any("original extraction observations unavailable" in note for note in notes)
    assert any("no source links inferred" in note for note in notes)


def test_legacy_captured_context_is_preserved_without_wall_clock(tmp_path: Path) -> None:
    destination = tmp_path / "legacy.json"
    destination.write_text(
        json.dumps(
            {
                "version": 2,
                "records": [],
                "config": {"as_of": "2020-01-01"},
                "extracted_at": "2020-01-02T00:00:00+00:00",
            }
        )
    )
    result = load_result(destination)
    assert result.prepared.extraction.request.as_of == date(2020, 1, 1)
    assert result.prepared.extraction.request.created_at == datetime(
        2020, 1, 2, tzinfo=timezone.utc
    )


def test_reanalysis_requires_and_records_a_new_request(
    snapshot: AnalysisSnapshot, tmp_path: Path
) -> None:
    destination = tmp_path / "result.json"
    save_result(snapshot, destination, result_id="validation-result")
    original = destination.read_bytes()
    with pytest.raises(ValueError, match="new captured request"):
        load_result(destination, reanalyse=True)
    result = load_result(
        destination, reanalyse=True, legacy_request=request("new-run", "2024-04-02")
    )
    assert result.prepared.extraction.request.run_id == "new-run"
    assert result.prepared.extraction.request.as_of == date(2024, 4, 2)
    assert result.input_fingerprint != snapshot.input_fingerprint
    assert [object_value(row.raw) for row in result.prepared.extraction.observations] == [
        object_value(row.raw) for row in snapshot.prepared.extraction.observations
    ]
    assert result.summary.charges == snapshot.summary.charges
    assert destination.read_bytes() == original
