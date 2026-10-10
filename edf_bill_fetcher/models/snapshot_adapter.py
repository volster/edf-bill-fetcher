"""Explicit legacy scalar conversion at the immutable snapshot boundary."""

from __future__ import annotations

import math
from dataclasses import fields, is_dataclass
from datetime import date, datetime
from decimal import Decimal
from typing import Any

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
        frame = pd.DataFrame(value["rows"], columns=value["columns"], index=value["index"])
        for column in ("_dt", "_dt_parsed"):
            if column in frame:
                frame[column] = pd.to_datetime(frame[column], errors="raise")
        return frame
    if value.get("type") == "model":
        if set(value) != {"type", "name", "fields"} or value.get("name") not in MODEL_TYPES:
            raise ValueError("Unsupported stored model")
        cls = MODEL_TYPES[value["name"]]
        if not isinstance(value["fields"], dict) or set(value["fields"]) != {
            field.name for field in fields(cls)
        }:
            raise ValueError("Stored model fields disagree with its contract")
        data = {key: from_snapshot_json(item) for key, item in value["fields"].items()}
        if cls is report_models.DisputeAnalysis:
            data["flags"] = [tuple(flag) for flag in data["flags"]]
        if cls is SapBackBillingEvent:
            for key in ("clearing_date",):
                data[key] = pd.Timestamp(data[key]) if data[key] is not None else pd.NaT
            data["posting_date_range"] = tuple(data["posting_date_range"])
        return cls(**data)
    return {key: from_snapshot_json(item) for key, item in value.items()}
