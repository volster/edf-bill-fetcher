"""Unit tests for the run-diff processor (``processors.run_diff``).

``diff_records`` binds the contract for comparing two record sets
(``--records-json`` shape: a list of dicts with ``Date``,
``Amount (£)``, ``Source``, …).  These tests pin:

- identical sets → empty diff (no phantom added/removed/changed),
- added / removed / changed classification with ``changed_fields``
  populated,
- NaN / None normalization in the canonical key (a NaN key field
  matches another NaN, not a real number),
- deterministic ordering of every output list (sorted by ``Date``).
"""

from __future__ import annotations

import math
from typing import Any

from edf_bill_fetcher.processors.run_diff import diff_records


def _record(
    date: str, amount: float | None, source: str = "Local PDF", **extra: Any
) -> dict[str, Any]:
    """Build a minimal record dict in the ``--records-json`` shape."""
    record: dict[str, Any] = {"Date": date, "Amount (£)": amount, "Source": source}
    record.update(extra)
    return record


class TestIdenticalSets:
    def test_identical_sets_produce_empty_diff(self):
        old_records = [
            _record("2026-03-01", 120.50),
            _record("2026-04-01", 130.00),
        ]
        new_records = [dict(row) for row in old_records]

        result = diff_records(old_records, new_records)

        assert result == {"added": [], "removed": [], "changed": []}


class TestClassification:
    def test_one_added_one_removed_one_amount_changed(self):
        old_records = [
            _record("2026-03-01", 120.50),
            _record("2026-04-01", 130.00),
            _record("2026-05-01", 140.00),  # dropped in the new run
        ]
        new_records = [
            _record("2026-03-01", 120.50),
            _record("2026-04-01", 155.00),  # same key, new amount
            _record("2026-06-01", 160.00),  # brand new row
        ]

        result = diff_records(old_records, new_records)

        assert [row["Date"] for row in result["added"]] == ["2026-06-01"]
        assert [row["Date"] for row in result["removed"]] == ["2026-05-01"]

        assert len(result["changed"]) == 1
        change = result["changed"][0]
        assert change["old_row"]["Amount (£)"] == 130.00
        assert change["new_row"]["Amount (£)"] == 155.00
        assert change["changed_fields"] == ["Amount (£)"]
        assert change["old_values"] == [130.00]
        assert change["new_values"] == [155.00]

    def test_non_key_field_change_detected(self):
        old_records = [
            _record("2026-04-01", 130.00, Details="Automatic estimate"),
        ]
        new_records = [
            _record("2026-04-01", 130.00, Details="Actual reading"),
        ]

        result = diff_records(old_records, new_records)

        assert result["added"] == []
        assert result["removed"] == []
        assert len(result["changed"]) == 1
        change = result["changed"][0]
        assert change["changed_fields"] == ["Details"]
        assert change["old_values"] == ["Automatic estimate"]
        assert change["new_values"] == ["Actual reading"]
        # Key fields are identical by construction and never reported.
        assert "Date" not in change["changed_fields"]
        assert "Amount (£)" not in change["changed_fields"]
        assert "Source" not in change["changed_fields"]

    def test_custom_key_fields(self):
        old_records = [
            {"Invoice #": "K123", "Date": "2026-04-01", "Amount (£)": 130.0},
        ]
        new_records = [
            {"Invoice #": "K123", "Date": "2026-04-01", "Amount (£)": 155.0},
        ]

        result = diff_records(old_records, new_records, key_fields=("Invoice #",))

        assert result["added"] == []
        assert result["removed"] == []
        assert len(result["changed"]) == 1
        assert result["changed"][0]["changed_fields"] == ["Amount (£)"]


class TestNanHandling:
    def test_nan_amount_matches_nan_not_number(self):
        old_records = [
            _record("2026-03-01", 120.50),
            _record("2026-04-01", float("nan")),
        ]
        new_records = [
            _record("2026-03-01", 120.50),
            _record("2026-04-01", float("nan")),
            _record("2026-05-01", float("nan")),  # genuinely new NaN row
        ]

        result = diff_records(old_records, new_records)

        # The shared NaN row matches (NaN normalizes like an empty key).
        assert result["changed"] == []
        assert [row["Date"] for row in result["added"]] == ["2026-05-01"]
        assert math.isnan(result["added"][0]["Amount (£)"])

    def test_nan_versus_number_reported_as_amount_change(self):
        # Same Date + Source, amount went from unknown (NaN) to £130.00:
        # the relaxed key pairs them as the same bill whose amount changed.
        old_records = [_record("2026-04-01", float("nan"))]
        new_records = [_record("2026-04-01", 130.00)]

        result = diff_records(old_records, new_records)

        assert result["added"] == []
        assert result["removed"] == []
        assert len(result["changed"]) == 1
        change = result["changed"][0]
        assert change["changed_fields"] == ["Amount (£)"]
        assert len(change["old_values"]) == 1
        assert math.isnan(change["old_values"][0])
        assert change["new_values"] == [130.00]


class TestOrdering:
    def test_outputs_sorted_by_date_regardless_of_input_order(self):
        old_records = [
            _record("2026-03-01", 110.00),
            _record("2026-02-01", 100.00),
            _record("2026-01-01", 90.00),
        ]
        new_records = [
            _record("2026-04-01", 200.00),
            _record("2026-02-01", 100.00),
            _record("2026-03-01", 110.00),
        ]

        expected = diff_records(old_records, new_records)
        shuffled = diff_records(list(reversed(old_records)), list(reversed(new_records)))

        assert [row["Date"] for row in expected["added"]] == ["2026-04-01"]
        assert [row["Date"] for row in expected["removed"]] == ["2026-01-01"]
        assert shuffled == expected
