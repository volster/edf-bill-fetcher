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


def _money(value: Any) -> Decimal:
    """Require canonical money; required totals can never decode to None."""
    if not isinstance(value, str) or not re.fullmatch(r"-?(?:0|[1-9][0-9]*)\.[0-9]{2}", value):
        raise ValueError("Money must be a canonical two-decimal string")
    if value == "-0.00":
        raise ValueError("Signed zero is not canonical money")
    return Decimal(value)


def _optional_money(value: Any) -> Decimal | None:
    """Preserve an unavailable observation while validating any amount supplied."""
    return None if value is None else _money(value)


def _validate_revisions(model: Any) -> None:
    model = _object(model, {"version", "notice", "relationships", "truncated"})
    if (
        type(model["version"]) is not int
        or model["version"] != 1
        or type(model["truncated"]) is not bool
    ):
        raise ValueError("Unsupported bill revision assessment")
    _string(model["notice"])
    links = model["relationships"]
    if not isinstance(links, list) or len(links) > 10500:
        raise ValueError("Invalid bill revision assessment")
    seen = set()
    for link in links:
        link = _object(
            link,
            {
                "older",
                "newer",
                "kind",
                "status",
                "reason",
                "basis",
                "earlier",
                "later",
                "earlier_included",
                "later_included",
                "changed_fields",
                "charge_difference",
            },
        )
        pair = (_string(link["older"]), _string(link["newer"]))
        if pair[0] == pair[1] or pair in seen:
            raise ValueError("Duplicate or identical revision endpoints")
        seen.add(pair)
        if (
            _string(link["kind"]) not in {"replacement", "cancellation_credit"}
            or _string(link["status"]) not in {"suggested", "confirmed", "rejected"}
            or not isinstance(link["reason"], str)
            or (link["status"] != "suggested" and not link["reason"].strip())
            or not isinstance(link["earlier"], dict)
            or not isinstance(link["later"], dict)
            or type(link["earlier_included"]) is not bool
            or type(link["later_included"]) is not bool
            or not isinstance(link["changed_fields"], list)
            or any(not isinstance(field, str) for field in link["changed_fields"])
            or (
                link["charge_difference"] is not None
                and type(link["charge_difference"]) not in (int, float)
            )
        ):
            raise ValueError("Invalid stored bill revision relationship")
        _string(link["basis"])


def _validate_ledger(value: Any) -> None:
    """Validate accepted history and positions without recomputing their conclusions."""
    from edf_bill_fetcher.processors.rebilling_ledger import validate_schedule

    model = _object(
        value, {"version", "notice", "documents", "comparisons", "suggestions_truncated"}
    )
    if (
        type(model["version"]) is not int
        or model["version"] != 1
        or type(model["suggestions_truncated"]) is not bool
    ):
        raise ValueError("Unsupported rebilling ledger")
    _string(model["notice"])
    docs, comparisons = model["documents"], model["comparisons"]
    if (
        not isinstance(docs, list)
        or len(docs) > 100000
        or not isinstance(comparisons, list)
        or len(comparisons) > 10500
    ):
        raise ValueError("Invalid rebilling ledger entries")
    indexed = {}

    def optional_date(day: Any) -> None:
        if day is not None:
            if not isinstance(day, str) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", day):
                raise ValueError("Invalid rebilling date")
            date.fromisoformat(day)

    for doc in docs:
        doc = _object(
            doc,
            {
                "row_id",
                "record",
                "included",
                "invoice_date",
                "reading_date",
                "period_from",
                "period_to",
                "first_demand_date",
                "cutoff_date",
                "components",
                "position",
                "needs_review",
                "incoming",
                "outgoing",
            },
        )
        key = _string(doc["row_id"])
        if (
            key in indexed
            or not isinstance(doc["record"], dict)
            or type(doc["included"]) is not bool
            or type(doc["needs_review"]) is not bool
        ):
            raise ValueError("Invalid or duplicate rebilling document")
        indexed[key] = doc
        for field in (
            "invoice_date",
            "reading_date",
            "period_from",
            "period_to",
            "first_demand_date",
            "cutoff_date",
        ):
            optional_date(doc[field])
        if doc["position"] is not None:
            normalized = validate_schedule(
                [{"id": key, "record": doc["record"]}], [doc["position"]]
            )
            if (
                normalized[0] != doc["position"]
                or doc["position"]["row_id"] != key
                or doc["first_demand_date"] != doc["position"]["first_demand_date"]
            ):
                raise ValueError("Noncanonical stored dispute position")
        elif (
            doc["first_demand_date"] is not None
            or doc["cutoff_date"] is not None
            or doc["needs_review"]
        ):
            raise ValueError("Rebilling dates or review status lack a dispute position")
        if not isinstance(doc["components"], list) or len(doc["components"]) > 10000:
            raise ValueError("Invalid rebilling components")
        for component in doc["components"]:
            component = _object(
                component, {"kind", "period_from", "period_to", "units", "rate", "amount", "detail"}
            )
            if (
                not isinstance(component["kind"], str)
                or component["kind"]
                not in {
                    "energy",
                    "standing",
                    "vat",
                    "cancellation",
                    "discount",
                }
                or not isinstance(component["detail"], str)
            ):
                raise ValueError("Invalid rebilling component kind")
            optional_date(component["period_from"])
            optional_date(component["period_to"])
            _optional_money(component["amount"])
            if any(
                component[field] is not None and type(component[field]) not in (int, float)
                for field in ("units", "rate")
            ):
                raise ValueError("Invalid rebilling component quantity")
    for doc in docs:
        for field, endpoint in (("incoming", "newer"), ("outgoing", "older")):
            if not isinstance(doc[field], list):
                raise ValueError("Invalid rebilling relationships")
            for link in doc[field]:
                if (
                    not isinstance(link, dict)
                    or link.get(endpoint) != doc["row_id"]
                    or not isinstance(link.get("older"), str)
                    or not isinstance(link.get("newer"), str)
                    or link.get("older") not in indexed
                    or link.get("newer") not in indexed
                    or not isinstance(link.get("status"), str)
                    or link.get("status") not in {"suggested", "confirmed", "rejected"}
                ):
                    raise ValueError("Rebilling relationship references an unknown document")
    seen = set()
    for comparison in comparisons:
        comparison = _object(
            comparison,
            {"older", "newer", "status", "same_period", "components", "other_components"},
        )
        pair = (_string(comparison["older"]), _string(comparison["newer"]))
        if (
            pair in seen
            or pair[0] == pair[1]
            or any(key not in indexed for key in pair)
            or type(comparison["same_period"]) is not bool
            or not isinstance(comparison["status"], str)
            or comparison["status"] not in {"suggested", "confirmed", "rejected"}
            or not isinstance(comparison["components"], list)
        ):
            raise ValueError("Invalid rebilling comparison")
        seen.add(pair)
        if not isinstance(comparison["other_components"], list):
            raise ValueError("Invalid other component comparisons")
        for component in comparison["other_components"]:
            component = _object(component, {"kind", "earlier_amount", "later_amount", "difference"})
            if not isinstance(component["kind"], str) or component["kind"] not in {
                "standing",
                "vat",
                "discount",
            }:
                raise ValueError("Invalid other component comparison kind")
            for field in ("earlier_amount", "later_amount", "difference"):
                _optional_money(component[field])
        for component in comparison["components"]:
            component = _object(
                component,
                {
                    "period_from",
                    "period_to",
                    "status",
                    "charge_difference",
                    "units_effect",
                    "rate_effect",
                    "residual",
                },
            )
            optional_date(component["period_from"])
            optional_date(component["period_to"])
            if not isinstance(component["status"], str) or component["status"] not in {
                "matched",
                "unallocated",
            }:
                raise ValueError("Invalid component comparison status")
            for field in ("charge_difference", "units_effect", "rate_effect", "residual"):
                validator = _optional_money if component["status"] == "unallocated" else _money
                validator(component[field])


def _validate_claims(value: Any, ledger: dict) -> None:
    """Check preserved proposal shapes and identities without recalculating totals."""
    from edf_bill_fetcher.processors.claim_reconciliation import validate_state

    model = _object(
        value,
        {
            "version",
            "notice",
            "input",
            "groups",
            "overlaps",
            "unresolved_positions",
            "reviewed_subtotal",
            "proposed_reduction",
            "ready",
            "ledgers",
        },
    )
    if (
        type(model["version"]) is not int
        or model["version"] != 1
        or type(model["ready"]) is not bool
    ):
        raise ValueError("Unsupported claim reconciliation model")
    _string(model["notice"])
    items = [
        {"id": doc["row_id"], "included": doc["included"], "record": doc["record"]}
        for doc in ledger["documents"]
    ]
    state = validate_state(items, model["input"])
    if state != model["input"]:
        raise ValueError("Stored claim inputs are not canonical")
    inputs = {group["id"]: group for group in state["groups"]}
    known = {item["id"] for item in items}
    if not isinstance(model["unresolved_positions"], list) or any(
        not isinstance(key, str) or key not in known for key in model["unresolved_positions"]
    ):
        raise ValueError("Unknown unresolved claim position")
    if len(set(model["unresolved_positions"])) != len(model["unresolved_positions"]):
        raise ValueError("Duplicate unresolved claim position")
    _money(model["reviewed_subtotal"])
    (_money if model["ready"] else _optional_money)(model["proposed_reduction"])
    if not model["ready"] and model["proposed_reduction"] is not None:
        raise ValueError("An unresolved reconciliation cannot publish a full proposed total")
    if not isinstance(model["groups"], list) or len(model["groups"]) != len(inputs):
        raise ValueError("Claim group membership disagrees")

    def validate_scope(value: Any) -> None:
        scope = _object(value, {"accounts", "Fuel", "Meter", "Register"})
        if any(
            not isinstance(values, list)
            or any(not isinstance(item, str) for item in values)
            or len(set(values)) != len(values)
            for values in scope.values()
        ):
            raise ValueError("Invalid claim ledger scope")

    seen = set()
    for group in model["groups"]:
        group = _object(
            group,
            {
                "input",
                "scope",
                "needs_review",
                "issues",
                "relief_received",
                "cancellation_offsets",
                "further_reduction",
                "ready",
            },
        )
        source = group["input"]
        if (
            not isinstance(source, dict)
            or not isinstance(source.get("id"), str)
            or source["id"] in seen
            or inputs.get(source["id"]) != source
        ):
            raise ValueError("Stored claim group disagrees with preserved inputs")
        seen.add(source["id"])
        if (
            any(type(group[field]) is not bool for field in ("needs_review", "ready"))
            or not isinstance(group["issues"], list)
            or any(not isinstance(issue, str) for issue in group["issues"])
        ):
            raise ValueError("Invalid claim group review status")
        validate_scope(group["scope"])
        if group["ready"] and (
            group["needs_review"]
            or group["issues"]
            or source["status"] != "reviewed"
            or group["further_reduction"] is None
        ):
            raise ValueError("Unresolved claim group cannot be ready for totaling")
        for field in ("relief_received", "cancellation_offsets", "further_reduction"):
            (_optional_money if field == "further_reduction" else _money)(group[field])
    if (
        not isinstance(model["overlaps"], list)
        or len(model["overlaps"]) > 500000
        or not isinstance(model["ledgers"], list)
    ):
        raise ValueError("Invalid claim overlap or ledger results")
    for overlap in model["overlaps"]:
        overlap = _object(overlap, {"groups", "kind", "resolved", "needs_review", "reason"})
        pair = overlap["groups"]
        if (
            not isinstance(pair, list)
            or len(pair) != 2
            or any(not isinstance(key, str) or key not in inputs for key in pair)
            or pair[0] >= pair[1]
            or not isinstance(overlap["kind"], str)
            or overlap["kind"]
            not in {
                "shared_documents",
                "unverified_account_relationship",
                "unknown_periods",
                "overlapping_periods",
            }
            or any(type(overlap[field]) is not bool for field in ("resolved", "needs_review"))
            or not isinstance(overlap["reason"], str)
        ):
            raise ValueError("Invalid claim overlap review")
        if overlap["resolved"] and (
            overlap["kind"] == "shared_documents" or overlap["needs_review"]
        ):
            raise ValueError("A shared or stale claim overlap cannot be resolved")
    for entry in model["ledgers"]:
        entry = _object(
            entry,
            {
                "scope",
                "groups",
                "reduction",
                "observed_balance",
                "balance_date",
                "illustrative_balance",
            },
        )
        if not isinstance(entry["groups"], list) or any(
            not isinstance(key, str) or key not in inputs for key in entry["groups"]
        ):
            raise ValueError("Unknown ledger claim group")
        validate_scope(entry["scope"])
        if len(set(entry["groups"])) != len(entry["groups"]):
            raise ValueError("Duplicate ledger claim group")
        _money(entry["reduction"])
        _optional_money(entry["observed_balance"])
        _optional_money(entry["illustrative_balance"])
        if entry["balance_date"] is not None:
            date.fromisoformat(_string(entry["balance_date"]))
    if model["ready"] and (
        model["unresolved_positions"]
        or any(not g["ready"] for g in model["groups"] if g["input"]["status"] != "withdrawn")
        or any(not o["resolved"] for o in model["overlaps"])
    ):
        raise ValueError("Unresolved claims cannot publish a full proposed total")


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
            if "ledger_overview" in restored:
                rows = restored["ledger_overview"]
                if not isinstance(rows, list) or len(rows) > 100000:
                    raise ValueError("Invalid ledger overview")
                for row in rows:
                    row = _object(
                        row,
                        {
                            "scope",
                            "references",
                            "records",
                            "charges",
                            "net_credits",
                            "opening_date",
                            "opening_balance",
                            "closing_date",
                            "closing_balance",
                        },
                    )
                    if (
                        not isinstance(row["scope"], list)
                        or len(row["scope"]) != 4
                        or any(not isinstance(part, str) for part in row["scope"])
                        or not isinstance(row["references"], list)
                        or any(not isinstance(ref, str) or not ref for ref in row["references"])
                        or type(row["records"]) is not int
                        or row["records"] < 1
                    ):
                        raise ValueError("Invalid ledger overview identity")
                    for key in ("charges", "net_credits"):
                        _money(row[key])
                    for prefix in ("opening", "closing"):
                        _optional_money(row[prefix + "_balance"])
                        day = row[prefix + "_date"]
                        if (day is None) != (row[prefix + "_balance"] is None):
                            raise ValueError("Ledger balance requires a dated observation")
                        if day is not None:
                            if (
                                not isinstance(day, str)
                                or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", day)
                                or date.fromisoformat(day).isoformat() != day
                            ):
                                raise ValueError("Invalid ledger observation date")
            if "bill_revisions" in restored:
                _validate_revisions(restored["bill_revisions"])
            if "rebilling_ledger" in restored:
                _validate_ledger(restored["rebilling_ledger"])
            if "claim_reconciliation" in restored:
                if "rebilling_ledger" not in restored:
                    raise ValueError("Claim reconciliation needs a preserved rebilling ledger")
                _validate_claims(restored["claim_reconciliation"], restored["rebilling_ledger"])
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
                or not isinstance(value[0], str)
                or not value[0].strip()
                or type(value[1]) is not bool
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
    as_of_text = _string(raw_request["as_of"])
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", as_of_text):
        raise ValueError("Valuation date must use ISO calendar format")
    as_of = date.fromisoformat(as_of_text)
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
    financial = FinancialSummary(
        charges,
        credits,
        _optional_money(summary["opening_balance"]),
        _optional_money(summary["closing_balance"]),
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
