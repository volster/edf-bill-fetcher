import pandas as pd

from edf_bill_fetcher.processors.detection import compute_transitive_domination


def test_transitive_domination_simple_chain() -> None:
    # A -> B -> C (A kills B, B kills C, so A transitively kills C)
    rebilling_df = pd.DataFrame(
        [
            {"Killer Invoice": "A", "Killed Invoice": "B"},
            {"Killer Invoice": "B", "Killed Invoice": "C"},
        ]
    )
    back_billing_rows = pd.DataFrame(
        [
            {"Invoice #": "A", "Period From": "2020-01-01", "Period To": "2021-06-01"},
            {"Invoice #": "B", "Period From": "2020-06-01", "Period To": "2021-12-01"},
            {"Invoice #": "C", "Period From": "2021-01-01", "Period To": "2022-06-01"},
        ]
    )
    result = compute_transitive_domination(rebilling_df, back_billing_rows)
    # C is superseded by A (transitive root)
    assert "C" in result
    assert result["C"][0] == "A"
    # B is superseded by A (direct)
    assert "B" in result
    assert result["B"][0] == "A"
    # A is live (not in result)
    assert "A" not in result


def test_transitive_domination_partial_overlap() -> None:
    # A -> B, but A's period does not fully contain B's period
    rebilling_df = pd.DataFrame(
        [
            {"Killer Invoice": "A", "Killed Invoice": "B"},
        ]
    )
    back_billing_rows = pd.DataFrame(
        [
            {"Invoice #": "A", "Period From": "2020-06-01", "Period To": "2021-06-01"},
            {"Invoice #": "B", "Period From": "2020-01-01", "Period To": "2021-12-01"},
        ]
    )
    result = compute_transitive_domination(rebilling_df, back_billing_rows)
    assert "B" in result
    assert result["B"][0] == "A"
    assert result["B"][1] is True  # partial_overlap flag


def test_transitive_domination_no_edges() -> None:
    rebilling_df = pd.DataFrame(columns=["Killer Invoice", "Killed Invoice"])
    back_billing_rows = pd.DataFrame(
        [
            {"Invoice #": "A", "Period From": "2020-01-01", "Period To": "2021-06-01"},
        ]
    )
    result = compute_transitive_domination(rebilling_df, back_billing_rows)
    assert result == {}  # all rows are live
