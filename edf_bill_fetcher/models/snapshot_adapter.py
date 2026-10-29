"""Explicit legacy scalar conversion at the immutable snapshot boundary."""

from __future__ import annotations

import math
import sys
from dataclasses import fields, is_dataclass
from datetime import date, datetime
from decimal import Decimal
from functools import cache
from types import UnionType
from typing import Any, Union, get_args, get_origin, get_type_hints

import numpy as np
import pandas as pd

from edf_bill_fetcher.models import report_models
from edf_bill_fetcher.models.events import SapBackBillingEvent, SapEdfMatch

MODEL_TYPES = {
    cls.__name__: cls
    for cls in (
        report_models.PaymentAnalysis,
        report_models.DataQualityReport,
        report_models.StatisticalAnalysis,
        report_models.ForecastResult,
        report_models.OfgemRow,
        report_models.OfgemComparison,
        report_models.TariffAnalysis,
        report_models.DisputeAnalysis,
        SapBackBillingEvent,
        SapEdfMatch,
    )
}


def to_snapshot_json(
    value: object, *, path: str = "$", diagnostics: list[dict] | None = None
) -> Any:
    """Convert known legacy values to JSON, rejecting unsupported objects by path."""
    if value is None or value is pd.NA or value is pd.NaT:
        return None
    if isinstance(value, np.generic):
        return to_snapshot_json(value.item(), path=path, diagnostics=diagnostics)
    if isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, (float, Decimal)):
        finite = value.is_finite() if isinstance(value, Decimal) else math.isfinite(value)
        if not finite:
            if diagnostics is not None:
                diagnostics.append({"code": "NONFINITE_VALUE", "path": path})
            return None
        return str(value) if isinstance(value, Decimal) else value
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, pd.DataFrame):
        return {
            "type": "frame",
            "columns": to_snapshot_json(
                list(value.columns), path=f"{path}.columns", diagnostics=diagnostics
            ),
            "rows": to_snapshot_json(
                value.values.tolist(), path=f"{path}.rows", diagnostics=diagnostics
            ),
            "index": to_snapshot_json(
                list(value.index), path=f"{path}.index", diagnostics=diagnostics
            ),
        }
    if is_dataclass(value) and not isinstance(value, type):
        if (
            type(value).__name__ not in MODEL_TYPES
            or type(value) is not MODEL_TYPES[type(value).__name__]
        ):
            raise ValueError(f"Unsupported model at {path}")
        return {
            "type": "model",
            "name": type(value).__name__,
            "fields": {
                field.name: to_snapshot_json(
                    getattr(value, field.name), path=f"{path}.{field.name}", diagnostics=diagnostics
                )
                for field in fields(value)
                if not (
                    type(value) is report_models.DataQualityReport
                    and field.name == "assessment"
                    and getattr(value, field.name) is None
                )
            },
        }
    if isinstance(value, dict):
        if any(not isinstance(key, str) for key in value):
            raise ValueError(f"Non-string key at {path}")
        return {
            key: to_snapshot_json(item, path=f"{path}.{key}", diagnostics=diagnostics)
            for key, item in value.items()
        }
    if isinstance(value, (list, tuple, np.ndarray)):
        return [
            to_snapshot_json(item, path=f"{path}[{index}]", diagnostics=diagnostics)
            for index, item in enumerate(value)
        ]
    if isinstance(value, (set, frozenset)) and path.endswith(".seen_pdf_hashes"):
        if all(isinstance(item, str) for item in value):
            return sorted(value)
    raise ValueError(f"Unsupported snapshot value at {path}: {type(value).__name__}")


@cache
def _model_hints(cls: type) -> dict:
    """Resolve annotations only for the fixed known model classes."""
    return get_type_hints(cls, globalns={**vars(sys.modules[cls.__module__]), "pd": pd})


def _valid_field(value: object, annotation: Any) -> bool:
    """Validate model scalar/container shapes; unavailable numeric metrics stay null."""
    if annotation is Any:
        return True
    origin, args = get_origin(annotation), get_args(annotation)
    if origin in (Union, UnionType):
        return any(_valid_field(value, item) for item in args)
    if annotation is float:
        return value is None or type(value) in (int, float)
    if annotation in (int, bool, str, type(None)):
        return type(value) is annotation
    if origin is list:
        return isinstance(value, list) and all(_valid_field(item, args[0]) for item in value)
    if origin is dict:
        return isinstance(value, dict) and all(
            _valid_field(key, args[0]) and _valid_field(item, args[1])
            for key, item in value.items()
        )
    if origin is tuple:
        return (
            isinstance(value, tuple)
            and len(value) == len(args)
            and all(_valid_field(item, kind) for item, kind in zip(value, args, strict=True))
        )
    return isinstance(value, annotation)


def from_snapshot_json(value: object) -> Any:
    """Restore only fixed computed-model types and fresh frames, without analysis."""
    if isinstance(value, list):
        return [from_snapshot_json(item) for item in value]
    if not isinstance(value, dict):
        return value
    if value.get("type") == "frame":
        if set(value) != {"type", "columns", "rows", "index"}:
            raise ValueError("Malformed stored frame")
        if (
            not isinstance(value["columns"], list)
            or not isinstance(value["rows"], list)
            or not isinstance(value["index"], list)
        ):
            raise ValueError("Stored frame axes must be arrays")
        if len(value["index"]) != len(value["rows"]) or any(
            not isinstance(row, list) or len(row) != len(value["columns"]) for row in value["rows"]
        ):
            raise ValueError("Stored frame dimensions disagree")
        if any(not isinstance(column, str) for column in value["columns"]) or len(
            set(value["columns"])
        ) != len(value["columns"]):
            raise ValueError("Stored frame columns must be unique strings")
        frame = pd.DataFrame(value["rows"], columns=value["columns"], index=value["index"])
        for column in ("_dt", "_dt_parsed"):
            if column in frame:
                frame[column] = pd.to_datetime(frame[column], errors="raise")
        return frame
    if value.get("type") == "model":
        if set(value) != {"type", "name", "fields"} or value.get("name") not in MODEL_TYPES:
            raise ValueError("Unsupported stored model")
        cls = MODEL_TYPES[value["name"]]
        expected = {field.name for field in fields(cls)}
        legacy_quality = (
            cls is report_models.DataQualityReport
            and isinstance(value["fields"], dict)
            and set(value["fields"]) == expected - {"assessment"}
        )
        if not isinstance(value["fields"], dict) or (
            set(value["fields"]) != expected and not legacy_quality
        ):
            raise ValueError("Stored model fields disagree with its contract")
        data = {key: from_snapshot_json(item) for key, item in value["fields"].items()}
        if legacy_quality:
            data["assessment"] = None
        elif cls is report_models.DataQualityReport and data["assessment"] is None:
            raise ValueError(
                "A stored quality assessment must be present or omitted in the legacy shape"
            )
        if cls is report_models.DisputeAnalysis:
            data["flags"] = [tuple(flag) for flag in data["flags"]]
        if cls is SapBackBillingEvent:
            for key in ("clearing_date",):
                data[key] = pd.Timestamp(data[key]) if data[key] is not None else pd.NaT
            data["posting_date_range"] = tuple(data["posting_date_range"])
        for field_name, annotation in _model_hints(cls).items():
            if not _valid_field(data[field_name], annotation):
                raise ValueError(f"Invalid stored {cls.__name__}.{field_name}")
            if annotation is int and data[field_name] < 0:
                raise ValueError(f"Negative stored count: {field_name}")
        if cls is report_models.DataQualityReport and data["assessment"] is not None:
            from edf_bill_fetcher.processors.quality_assessment import validate_quality_assessment

            validate_quality_assessment(
                data["assessment"], data["total_records"], data["duplicate_count"]
            )
            denominator = data["total_records"] + data["assessment"]["duplicates_removed"]
            rate = data["duplicate_count"] / denominator if denominator else 0
            if data["duplicate_rate"] is None or abs(data["duplicate_rate"] - rate) > 1e-12:
                raise ValueError("Quality duplicate rate disagrees with its scope")
        return cls(**data)
    return {key: from_snapshot_json(item) for key, item in value.items()}
