"""Fixed business examples, authored independently of the candidate implementation.

Expected values must not be regenerated from application output. These cases protect
financial roles, transaction identity, evidence retention and forecast eligibility.
"""

import copy
import json
from pathlib import Path

import pytest

from edf_bill_fetcher.application.acceptance_api import probe

CASES = json.loads(
    (Path(__file__).parent / "fixtures/financial_contracts.json").read_text(encoding="utf-8")
)["cases"]


def assert_value(actual, expected):
    if type(expected) is float:
        assert type(actual) in (int, float)
        assert actual == pytest.approx(expected, rel=1e-9, abs=1e-9)
    elif type(expected) is bool:
        assert actual is expected
    else:
        assert actual == expected


@pytest.mark.parametrize("case", CASES, ids=lambda case: case["id"] + "-" + case["name"])
def test_independent_financial_contract(case: dict) -> None:
    original = copy.deepcopy(case)
    actual = probe(case["records"], case["config"])
    for key, expected in case["expected"].items():
        assert key in actual
        assert_value(actual[key], expected)
    for check in case.get("row_checks", []):
        matches = [
            row
            for row in actual["records"]
            if all(row.get(key) == value for key, value in check["select"].items())
        ]
        assert len(matches) == 1
        row = matches[0]
        for key, expected in check.get("values", {}).items():
            assert key in row
            assert_value(row[key], expected)
        if "source_ids" in check:
            assert sorted(row["Source IDs"]) == sorted(check["source_ids"])
    assert case == original
