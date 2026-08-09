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


def test_transitive_domination_non_bb_killer_supersedes_bb_invoice() -> None:
    # Gap 2: a back-billing invoice (B) is killed by a NON-back-billing
    # invoice (R, a regular monthly rebill). R is NOT in back_billing_rows.
    # Under the widened edge filter (v in bb_ids), the edge (R kills B)
    # is retained, B is marked Superseded, and R is the survivor even
    # though R has no row in the back-billing sheet.
    rebilling_df = pd.DataFrame(
        [
            {"Killer Invoice": "R", "Killed Invoice": "B"},
        ]
    )
    back_billing_rows = pd.DataFrame(
        [
            {"Invoice #": "B", "Period From": "2020-01-01", "Period To": "2021-06-01"},
        ]
    )
    result = compute_transitive_domination(rebilling_df, back_billing_rows)
    assert "B" in result
    survivor, partial_overlap = result["B"]
    assert survivor == "R"
    # R is not in period_map (not a bb row), so partial_overlap defaults to False.
    assert partial_overlap is False
    # R itself is not a key (it's not a bb row, so it's never a target).
    assert "R" not in result


def test_transitive_domination_non_bb_killer_via_transitive_chain() -> None:
    # Gap 2 transitive case: R (non-bb) kills A (bb), A kills B (bb).
    # B should be superseded by R (the transitive root), even though R
    # is not a back-billing row. A is superseded by R directly.
    rebilling_df = pd.DataFrame(
        [
            {"Killer Invoice": "R", "Killed Invoice": "A"},
            {"Killer Invoice": "A", "Killed Invoice": "B"},
        ]
    )
    back_billing_rows = pd.DataFrame(
        [
            {"Invoice #": "A", "Period From": "2020-01-01", "Period To": "2021-06-01"},
            {"Invoice #": "B", "Period From": "2020-06-01", "Period To": "2021-12-01"},
        ]
    )
    result = compute_transitive_domination(rebilling_df, back_billing_rows)
    # A is superseded by R (direct, non-bb killer).
    assert "A" in result
    assert result["A"][0] == "R"
    assert result["A"][1] is False  # R not in period_map -> partial_overlap False
    # B is superseded by R (transitive root, non-bb).
    assert "B" in result
    assert result["B"][0] == "R"
    assert result["B"][1] is False  # R not in period_map -> partial_overlap False
    # R is not a target (not a bb row).
    assert "R" not in result
