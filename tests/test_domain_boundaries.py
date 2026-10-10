"""Input contracts that prevent mutable, ambiguous or nonfinite financial evidence."""

from dataclasses import replace
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal
from typing import Any, cast

import pytest

from edf_bill_fetcher.application.legacy_adapter import capture_legacy
from edf_bill_fetcher.domain.json_values import (
    FrozenJSON,
    freeze_json,
    object_value,
    rows_value,
    strict_loads,
)
from edf_bill_fetcher.domain.money import money_text, parse_money
from edf_bill_fetcher.domain.requests import RunRequest, SourceSelection


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("£1,234.565", Decimal("1234.57")),
        ("-1.005", Decimal("-1.01")),
        ("-0.004", Decimal("0.00")),
        (0, Decimal("0.00")),
        (None, None),
        (True, None),
        ("NaN", None),
        (float("inf"), None),
        ("1e400", None),
        ("N/A", None),
    ],
)
def test_money_rounding_sign_and_missingness(value, expected):
    assert parse_money(value) == expected
    if expected is not None:
        assert money_text(parse_money(value)) == format(expected, ".2f")
    else:
        assert money_text(parse_money(value)) is None


def test_nonfinite_money_cannot_be_formatted_as_known():
    with pytest.raises(ValueError, match="nonfinite"):
        money_text(Decimal("NaN"))


@pytest.mark.parametrize(
    "text", ['{"nested":{"x":1,"x":2}}', "Infinity", "-Infinity", '"\\ud800"', '{"\\ud800":1}']
)
def test_json_rejects_ambiguity_and_invalid_unicode(text: str) -> None:
    with pytest.raises(ValueError):
        strict_loads(text)


@pytest.mark.parametrize("value", [b"{}", {1: "value"}, (1, 2), {"a": object()}, float("nan")])
def test_non_json_values_do_not_enter_immutable_evidence(value: object) -> None:
    with pytest.raises(ValueError):
        freeze_json(value)


def test_cyclic_evidence_rejected_but_shared_nested_values_detached():
    cyclic: list = []
    cyclic.append(cyclic)
    with pytest.raises(ValueError, match="Cyclic"):
        freeze_json(cyclic)
    shared = {"values": [1]}
    original = {"a": shared, "b": shared}
    frozen = freeze_json(original)
    shared["values"].append(2)
    decoded = object_value(frozen)
    decoded["a"]["values"].append(3)
    assert decoded["b"]["values"] == [1]
    assert object_value(frozen) == {"a": {"values": [1]}, "b": {"values": [1]}}


@pytest.mark.parametrize("text", ['{ "a": 1 }', '{"b":2,"a":1}'])
def test_direct_frozen_text_must_be_canonical(text: str) -> None:
    with pytest.raises(ValueError, match="canonical"):
        FrozenJSON(text)


def test_json_shape_is_checked_at_object_and_rows_boundaries():
    with pytest.raises(ValueError, match="object"):
        object_value(freeze_json([]))
    with pytest.raises(ValueError, match="array of objects"):
        rows_value(freeze_json([{}, 1]))
    with pytest.raises(ValueError, match="string"):
        strict_loads(cast(Any, b"{}"))


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("run_id", " "),
        ("created_at", datetime(2024, 1, 1)),
        ("created_at", datetime(2024, 1, 1, tzinfo=timezone(timedelta(hours=1)))),
        ("as_of", datetime(2024, 1, 1)),
        ("sources", []),
        ("sources", ("fake",)),
        ("sources", (SourceSelection("htm", "a"), SourceSelection("htm", "a"))),
        ("config", freeze_json({"as_of": "2024-01-02"})),
    ],
)
def test_request_cannot_capture_ambiguous_context(field: str, value: object) -> None:
    request = RunRequest(
        "run",
        datetime(2024, 1, 1, tzinfo=timezone.utc),
        date(2024, 1, 1),
        (),
        freeze_json({"as_of": "2024-01-01"}),
    )
    with pytest.raises(ValueError):
        replace(request, **cast(Any, {field: value}))


@pytest.mark.parametrize(("kind", "path"), [("unknown", "file"), ("htm", " ")])
def test_unsupported_source_selections_rejected(kind: str, path: str) -> None:
    with pytest.raises(ValueError):
        SourceSelection(kind, path)


def test_capture_preserves_originals_and_distinct_repeated_rows():
    request = RunRequest(
        "run",
        datetime(2024, 1, 1, tzinfo=timezone.utc),
        date(2024, 1, 1),
        (),
        freeze_json({"as_of": "2024-01-01"}),
    )
    rows: list[dict[str, Any]] = [{"Amount (£)": "bad", "extension": [1]}] * 2
    evidence = capture_legacy(rows, {}, request)
    rows[0]["extension"].append(2)
    assert len({row.observation_id for row in evidence.observations}) == 2
    assert object_value(evidence.observations[0].raw) == {"Amount (£)": "bad", "extension": [1]}
    with pytest.raises(ValueError, match="conflicts"):
        capture_legacy(rows, {"records": []}, request)
    with pytest.raises(ValueError, match="array of objects"):
        capture_legacy(cast(Any, [1]), {}, request)
