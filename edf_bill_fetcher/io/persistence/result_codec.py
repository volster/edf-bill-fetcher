"""Strict, versioned result transport; v3 decoding never recalculates analysis."""

from __future__ import annotations

import hashlib
import re
from datetime import date, datetime
from decimal import Decimal
from typing import Any

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
    for entry in models.values():
        _object(entry, {"availability", "reason_code", "payload"})
        if entry["availability"] not in {"available", "partial", "unavailable", "not_applicable"}:
            raise ValueError("Invalid model availability")
        reason = entry["reason_code"]
        if reason is not None and (not isinstance(reason, str) or not reason):
            raise ValueError("Invalid model reason")
        if entry["availability"] in {"unavailable", "partial"} and reason is None:
            raise ValueError("Unavailable/partial analysis needs a reason")
        from_snapshot_json(entry["payload"])


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
    return AnalysisSnapshot(
        prepared,
        financial,
        freeze_json(analysis["models"]),
        freeze_json(analysis["methods"]),
        analysis["input_fingerprint"],
    )
