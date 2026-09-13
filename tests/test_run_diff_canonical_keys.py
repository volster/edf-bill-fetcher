import pytest

from edf_bill_fetcher.processors.run_diff import diff_records


@pytest.mark.parametrize("count", [1, 2])
def test_canonical_key_representations_are_unchanged(count: int) -> None:
    old = [{"id": 1, "Details": "same"} for _ in range(count)]
    new = [{"id": "1", "Details": "same"} for _ in range(count)]

    result = diff_records(old, new, key_fields=("id",))

    assert result == {"added": [], "removed": [], "changed": []}


def test_relaxed_match_reports_amount_not_canonical_identity() -> None:
    old = [{"id": 1, "Amount (£)": 100}]
    new = [{"id": "1", "Amount (£)": 150}]

    result = diff_records(old, new, key_fields=("id", "Amount (£)"))

    assert result["added"] == result["removed"] == []
    assert result["changed"] == [
        {
            "old_row": old[0],
            "new_row": new[0],
            "changed_fields": ["Amount (£)"],
            "old_values": [100],
            "new_values": [150],
        }
    ]
