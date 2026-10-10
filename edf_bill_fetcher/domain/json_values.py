"""Strict canonical JSON values with immutable storage and fresh decoding."""

from __future__ import annotations

import json
import math
from dataclasses import dataclass
from typing import Any, cast


def _unique_object(pairs: list[tuple[str, object]]) -> dict[str, object]:
    result: dict[str, object] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"Duplicate JSON key: {key}")
        result[key] = value
    return result


def _invalid_constant(value: str) -> object:
    raise ValueError(f"Nonstandard JSON constant: {value}")


def strict_loads(text: str) -> object:
    """Decode JSON without duplicate keys or nonfinite numeric constants."""
    if not isinstance(text, str):
        raise ValueError("JSON text must be a string")
    value = json.loads(text, object_pairs_hook=_unique_object, parse_constant=_invalid_constant)
    _validate(value, "$", set())
    return cast(object, value)


def _validate(value: object, path: str, active: set[int]) -> None:
    if value is None or type(value) in (bool, int):
        return
    if isinstance(value, str):
        try:
            value.encode("utf-8")
        except UnicodeEncodeError as error:
            raise ValueError(f"Invalid Unicode at {path}") from error
        return
    if type(value) is float:
        if not math.isfinite(value):
            raise ValueError(f"Nonfinite number at {path}")
        return
    if type(value) not in (dict, list):
        raise ValueError(f"Unsupported JSON value at {path}: {type(value).__name__}")
    marker = id(value)
    if marker in active:
        raise ValueError(f"Cyclic JSON value at {path}")
    active.add(marker)
    if isinstance(value, dict):
        for key, item in value.items():
            if not isinstance(key, str):
                raise ValueError(f"Non-string JSON key at {path}")
            _validate(key, path, active)
            _validate(item, f"{path}.{key}", active)
    else:
        for index, item in enumerate(cast(list, value)):
            _validate(item, f"{path}[{index}]", active)
    active.remove(marker)


def canonical_text(value: object) -> str:
    """Encode validated standard JSON in deterministic UTF-8-compatible form."""
    _validate(value, "$", set())
    return json.dumps(
        value, sort_keys=True, separators=(",", ":"), ensure_ascii=False, allow_nan=False
    )


@dataclass(frozen=True)
class FrozenJSON:
    """Canonical JSON text whose nested values cannot mutate in place."""

    text: str

    def __post_init__(self) -> None:
        """Reject malformed or noncanonical direct construction."""
        if canonical_text(strict_loads(self.text)) != self.text:
            raise ValueError("FrozenJSON text must be canonical")


def freeze_json(value: object) -> FrozenJSON:
    """Validate and detach a JSON value into immutable canonical storage."""
    return FrozenJSON(canonical_text(value))


def thaw_json(value: FrozenJSON) -> object:
    """Return a fresh JSON structure on each decode."""
    return strict_loads(value.text)


def object_value(value: FrozenJSON) -> dict[str, Any]:
    """Decode an object with a checked boundary type."""
    decoded = thaw_json(value)
    if not isinstance(decoded, dict):
        raise ValueError("Expected JSON object")
    return decoded


def rows_value(value: FrozenJSON) -> list[dict[str, Any]]:
    """Decode an array of evidence objects with a checked boundary type."""
    decoded = thaw_json(value)
    if not isinstance(decoded, list) or any(not isinstance(row, dict) for row in decoded):
        raise ValueError("Expected JSON array of objects")
    return decoded
