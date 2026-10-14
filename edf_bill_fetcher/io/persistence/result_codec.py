"""Strict, versioned result transport; v3 decoding never recalculates analysis."""

from __future__ import annotations

import hashlib
import re
from datetime import date, datetime
from decimal import Decimal
from typing import Any

import pandas as pd

from edf_bill_fetcher.domain.analysis import AnalysisSnapshot, FinancialSummary
from edf_bill_fetcher.domain.evidence import ExtractionSnapshot, Observation, PreparedSnapshot
from edf_bill_fetcher.domain.json_values import (
    canonical_text,
    freeze_json,
    object_value,
    rows_value,
)
from edf_bill_fetcher.domain.money import money_text
from edf_bill_fetcher.domain.requests import RunRequest, SourceSelection
from edf_bill_fetcher.models.events import SapBackBillingEvent, SapEdfMatch
from edf_bill_fetcher.models.snapshot_adapter import from_snapshot_json

MODEL_KEYS = frozenset(
    {
        "payment",
        "data_quality",
        "statistical",
        "forecast",
        "ofgem",
        "tariff",
        "dispute",
        "compensation",
        "sap",
        "reconciliation",
    }
)


def _object(value: Any, keys: set[str]) -> dict:
    if not isinstance(value, dict) or set(value) != keys:
        raise ValueError(f"Expected object fields: {sorted(keys)}")
    return value


def _string(value: Any) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValueError("Expected nonempty string")
    return value


def _money(value: Any, nullable: bool = False) -> Decimal | None:
    if nullable and value is None:
        return None
    if not isinstance(value, str) or not re.fullmatch(r"-?(?:0|[1-9][0-9]*)\.[0-9]{2}", value):
        raise ValueError("Money must be a canonical two-decimal string")
    if value == "-0.00":
        raise ValueError("Signed zero is not canonical money")
    return Decimal(value)


def _validate_models(value: Any) -> None:
    models = _object(value, set(MODEL_KEYS))
    expected = {
        "payment": "PaymentAnalysis",
        "data_quality": "DataQualityReport",
        "statistical": "StatisticalAnalysis",
        "forecast": "ForecastResult",
        "ofgem": "OfgemComparison",
        "tariff": "TariffAnalysis",
        "dispute": "DisputeAnalysis",
    }
    for key, entry in models.items():
        _object(entry, {"availability", "reason_code", "payload"})
        if not isinstance(entry["availability"], str) or entry["availability"] not in {
            "available",
            "partial",
            "unavailable",
            "not_applicable",
        }:
            raise ValueError("Invalid model availability")
        reason = entry["reason_code"]
        if reason is not None and (not isinstance(reason, str) or not reason):
            raise ValueError("Invalid model reason")
        if entry["availability"] in {"unavailable", "partial"} and reason is None:
            raise ValueError("Unavailable/partial analysis needs a reason")
        payload = entry["payload"]
        if key in expected and (
            not isinstance(payload, dict)
            or payload.get("type") != "model"
            or payload.get("name") != expected[key]
        ):
            raise ValueError(f"Stored {key} payload has the wrong model type")
        restored = from_snapshot_json(payload)
        if key == "compensation" and (
            not isinstance(restored, list) or any(not isinstance(row, dict) for row in restored)
        ):
            raise ValueError("Compensation rows must be objects")
        if key == "sap":
            if not isinstance(restored, dict) or not {
                "analyses",
                "events",
                "matches",
                "position",
                "frames",
                "timeline",
                "event_source_indexes",
                "union_total",
                "domination",
            } <= set(restored):
                raise ValueError("SAP payload is incomplete")
            frames = restored["frames"]
            if (
                not isinstance(frames, dict)
                or not {
                    "evidence",
                    "analysis",
                    "balance",
                    "statistical",
                    "monthly",
                    "daily_rates",
                    "yearly",
                }
                <= frames.keys()
                or any(not isinstance(frame, pd.DataFrame) for frame in frames.values())
            ):
                raise ValueError("SAP precomputed frames are incomplete or invalid")
            analyses = restored["analyses"]
            if (
                not isinstance(analyses, dict)
                or not {
                    "back_billing",
                    "rebilling",
                    "meter_rollover",
                    "contracts",
                    "evidence_index",
                }
                <= analyses.keys()
                or any(
                    not isinstance(analyses[name], pd.DataFrame)
                    for name in ("back_billing", "rebilling", "meter_rollover", "contracts")
                )
                or not isinstance(analyses["evidence_index"], dict)
            ):
                raise ValueError("SAP precomputed analyses are incomplete or invalid")
            events = restored["events"]
            source_indexes = restored["event_source_indexes"]
            if (
                not isinstance(events, list)
                or any(not isinstance(event, SapBackBillingEvent) for event in events)
                or not isinstance(source_indexes, list)
                or len(events) != len(source_indexes)
                or any(not isinstance(indexes, list) for indexes in source_indexes)
            ):
                raise ValueError("SAP event source membership disagrees")
            for event, indexes in zip(events, source_indexes, strict=True):
                if len(event.rows) != len(indexes) or any(
                    type(idx) is not int or idx < 0 for idx in indexes
                ):
                    raise ValueError("Invalid SAP source index")
            matches = restored["matches"]
            if not isinstance(matches, list) or any(
                not isinstance(match, SapEdfMatch) or match.event not in events for match in matches
            ):
                raise ValueError("Invalid SAP event match membership")
            position = restored["position"]
            if (
                not isinstance(position, dict)
                or not {"events", "reconciliation", "summary"} <= position.keys()
                or not isinstance(position["summary"], dict)
                or any(
                    not isinstance(position[name], list)
                    or any(not isinstance(row, dict) for row in position[name])
                    for name in ("events", "reconciliation")
                )
            ):
                raise ValueError("Invalid SAP position payload")
            domination = restored["domination"]
            if not isinstance(domination, dict) or any(
                not isinstance(value, list)
                or len(value) != 2
                or any(not isinstance(item, str) for item in value)
                for value in domination.values()
            ):
                raise ValueError("Invalid SAP superseded invoice mapping")
            timeline = restored["timeline"]
            if not isinstance(timeline, list) or any(
                not isinstance(row, list)
                or len(row) != 3
                or any(isinstance(item, (dict, list)) for item in row)
                for row in timeline
            ):
                raise ValueError("Invalid SAP timeline payload")
            if restored["union_total"] is not None and type(restored["union_total"]) not in (
                int,
                float,
            ):
                raise ValueError("Invalid SAP union total")
        if key == "reconciliation" and (
            not isinstance(restored, dict)
            or set(restored) != {"summaries", "details"}
            or any(not isinstance(restored[name], list) for name in ("summaries", "details"))
        ):
            raise ValueError("Invalid reconciliation payload")


def encode_result(result: AnalysisSnapshot, result_id: str) -> dict:
    """Create a detached schema-v3 object, retaining originals and extensions."""
    request = result.prepared.extraction.request
    prepared = result.prepared
    summary = result.summary
    payload = {
        "schema_version": 3,
        "result_id": _string(result_id),
        "request": {
            "run_id": request.run_id,
            "created_at": request.created_at.isoformat(),
            "as_of": request.as_of.isoformat(),
            "sources": [{"kind": source.kind, "path": source.path} for source in request.sources],
            "config": object_value(request.config),
        },
        "evidence": {
            "observations": [
                {
                    "observation_id": row.observation_id,
                    "ordinal": row.ordinal,
                    "raw": object_value(row.raw),
                }
                for row in prepared.extraction.observations
            ],
            "legacy_state": object_value(prepared.extraction.legacy_state),
            "prepared_records": rows_value(prepared.records),
            "duplicate_records": rows_value(prepared.duplicates),
            "effective_config": object_value(prepared.effective_config),
        },
        "analysis": {
            "summary": {
                "charges": money_text(summary.charges),
                "net_credits": money_text(summary.net_credits),
                "opening_balance": money_text(summary.opening_balance),
                "closing_balance": money_text(summary.closing_balance),
                "mean_daily_charge": summary.mean_daily_charge,
            },
            "models": object_value(result.models),
            "methods": object_value(result.methods),
            "input_fingerprint": result.input_fingerprint,
        },
    }
    decode_result(payload)
    return payload


def decode_result(payload: object) -> AnalysisSnapshot:
    """Validate an accepted result without collecting, preparing or computing."""
    canonical_text(payload)
    data = _object(payload, {"schema_version", "result_id", "request", "evidence", "analysis"})
    if type(data["schema_version"]) is not int or data["schema_version"] != 3:
        raise ValueError("Unsupported result schema version")
    _string(data["result_id"])
    raw_request = _object(data["request"], {"run_id", "created_at", "as_of", "sources", "config"})
    if not isinstance(raw_request["sources"], list):
        raise ValueError("Sources must be an array")
    sources = []
    for source in raw_request["sources"]:
        _object(source, {"kind", "path"})
        sources.append(SourceSelection(_string(source["kind"]), _string(source["path"])))
    as_of = date.fromisoformat(_string(raw_request["as_of"]))
    if as_of.isoformat() != raw_request["as_of"]:
        raise ValueError("Valuation date must use ISO calendar format")
    request = RunRequest(
        _string(raw_request["run_id"]),
        datetime.fromisoformat(_string(raw_request["created_at"])),
        as_of,
        tuple(sources),
        freeze_json(raw_request["config"]),
    )
    evidence = _object(
        data["evidence"],
        {
            "observations",
            "legacy_state",
            "prepared_records",
            "duplicate_records",
            "effective_config",
        },
    )
    if not isinstance(evidence["observations"], list):
        raise ValueError("Observations must be an array")
    observations = []
    for row in evidence["observations"]:
        _object(row, {"observation_id", "ordinal", "raw"})
        observation = Observation(row["observation_id"], row["ordinal"], freeze_json(row["raw"]))
        expected = hashlib.sha256(
            canonical_text([request.run_id, observation.ordinal, row["raw"]]).encode()
        ).hexdigest()
        if expected != observation.observation_id:
            raise ValueError("Observation identity does not match original evidence")
        observations.append(observation)
    extraction = ExtractionSnapshot(
        request, tuple(observations), freeze_json(evidence["legacy_state"])
    )
    prepared = PreparedSnapshot(
        extraction,
        freeze_json(evidence["prepared_records"]),
        freeze_json(evidence["duplicate_records"]),
        freeze_json(evidence["effective_config"]),
    )
    analysis = _object(data["analysis"], {"summary", "models", "methods", "input_fingerprint"})
    summary = _object(
        analysis["summary"],
        {"charges", "net_credits", "opening_balance", "closing_balance", "mean_daily_charge"},
    )
    daily = summary["mean_daily_charge"]
    if type(daily) not in (int, float):
        raise ValueError("Daily charge must be numeric")
    charges, credits = _money(summary["charges"]), _money(summary["net_credits"])
    if charges is None or credits is None:
        raise ValueError("Summary totals must be present")
    financial = FinancialSummary(
        charges,
        credits,
        _money(summary["opening_balance"], True),
        _money(summary["closing_balance"], True),
        float(daily),
    )
    _validate_models(analysis["models"])
    restored = from_snapshot_json(analysis["models"]["sap"]["payload"])
    legacy = object_value(extraction.legacy_state)
    financial_rows = legacy.get("sap_financial_rows", [])
    for event, indexes in zip(restored["events"], restored["event_source_indexes"], strict=True):
        if any(index >= len(financial_rows) for index in indexes):
            raise ValueError("SAP event references a missing source row")
        if any(
            row != financial_rows[index] for row, index in zip(event.rows, indexes, strict=True)
        ):
            raise ValueError("SAP event rows disagree with preserved source rows")
    recon = analysis["models"]["reconciliation"]["payload"]
    for detail in recon["details"]:
        if not isinstance(detail, dict) or set(detail) != {"entity", "side", "index", "values"}:
            raise ValueError("Malformed reconciliation detail")
        entity, side, index = detail["entity"], detail["side"], detail["index"]
        if (
            not isinstance(entity, str)
            or entity not in {"Contract", "Meter Read", "Financial"}
            or not isinstance(side, str)
            or side not in {"sap", "edf"}
            or type(index) is not int
            or index < 0
        ):
            raise ValueError("Invalid reconciliation source reference")
        if side == "sap":
            limit = len(
                legacy.get(
                    {
                        "Contract": "sap_contract_rows",
                        "Meter Read": "sap_meter_rows",
                        "Financial": "sap_financial_rows",
                    }[entity],
                    [],
                )
            )
        else:
            limit = (
                len(restored["analyses"]["contracts"])
                if entity == "Contract"
                else len(rows_value(prepared.records))
            )
        if index >= limit:
            raise ValueError("Reconciliation references a missing source row")
    return AnalysisSnapshot(
        prepared,
        financial,
        freeze_json(analysis["models"]),
        freeze_json(analysis["methods"]),
        analysis["input_fingerprint"],
    )


def migrate_legacy(payload: object, request: RunRequest | None = None) -> AnalysisSnapshot:
    """Analyse preserved legacy evidence, recording unavailable historical context."""
    from edf_bill_fetcher.application.analysis_service import analyse_prepared
    from edf_bill_fetcher.application.legacy_adapter import capture_legacy
    from edf_bill_fetcher.application.preparation_service import prepare_snapshot

    if isinstance(payload, list):
        state: dict = {"records": payload}
    elif isinstance(payload, dict):
        state = dict(payload)
        version = state.get("version", 1)
        if type(version) is not int or version not in (1, 2) or "schema_version" in state:
            raise ValueError("Unsupported legacy evidence version")
        if "records" not in state:
            raise ValueError("Legacy evidence must contain records")
    else:
        raise ValueError("Legacy evidence must be an object or records array")
    records = state.get("original_records", state["records"])
    if not isinstance(records, list) or any(not isinstance(row, dict) for row in records):
        raise ValueError("Legacy records must be objects")
    cfg = state.get("config", {})
    if not isinstance(cfg, dict):
        raise ValueError("Legacy configuration must be an object")
    if request is None:
        if not cfg.get("as_of") or not state.get("extracted_at"):
            raise ValueError("Legacy input lacks captured date/context; supply legacy_request")
        raw_created = _string(state["extracted_at"])
        run_id = "legacy-" + hashlib.sha256(canonical_text(payload).encode()).hexdigest()
        request = RunRequest(
            run_id,
            datetime.fromisoformat(raw_created),
            date.fromisoformat(cfg["as_of"]),
            (),
            freeze_json(cfg),
        )
    notes = ["Analysis recomputed using current legacy methods; historical methodology unavailable"]
    if "original_records" not in state:
        notes.append("Canonical-only input: original extraction observations unavailable")
    if not state.get("source_manifest"):
        notes.append("Original source manifest unavailable; no source links inferred")
    state["migration_notes"] = notes
    return analyse_prepared(prepare_snapshot(capture_legacy(records, state, request)))
