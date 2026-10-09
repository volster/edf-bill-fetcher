"""Atomic publication of complete generated files."""

from __future__ import annotations

import inspect
import os
import tempfile
from functools import wraps
from pathlib import Path


def atomic_output(parameter="output_path"):
    """Write beside the destination and publish only after successful generation."""

    def decorate(function):
        signature = inspect.signature(function)

        @wraps(function)
        def wrapped(*args, **kwargs):
            bound = signature.bind(*args, **kwargs)
            destination = Path(bound.arguments[parameter])
            fd, temporary = tempfile.mkstemp(
                prefix=".edf-", suffix=destination.suffix, dir=destination.parent
            )
            os.close(fd)
            try:
                bound.arguments[parameter] = temporary
                result = function(*bound.args, **bound.kwargs)
                os.replace(temporary, destination)
                return str(destination) if result == temporary else result
            finally:
                if os.path.exists(temporary):
                    os.unlink(temporary)

        return wrapped

    return decorate


def json_safe(value):
    """Serialize missing/nonfinite values as JSON null and retain structured metadata."""
    import math

    if isinstance(value, dict):
        return {str(key): json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [json_safe(item) for item in value]
    if isinstance(value, float) and not math.isfinite(value):
        return None
    return value


@atomic_output("path")
def write_json(payload, path):
    """Publish a JSON artifact atomically without nonstandard NaN values."""
    import json

    with open(path, "w", encoding="utf-8") as handle:
        json.dump(json_safe(payload), handle, indent=2, default=str, allow_nan=False)
