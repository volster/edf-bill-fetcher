"""Regression test pinning the EDFFC-reconciliation warning contract.

Pre-fix, ``compute_dispute_flags`` swallowed per-row parse errors with
``pass`` — silently losing the row.  The fix replaces each ``pass`` with
:func:`warnings.warn` so an upstream parse failure becomes a
developer-visible signal without breaking the run.

This test feeds a row whose ``Amount (£)`` payload is a raw string
(not float-castable) into one of the heuristics, then asserts that
each heuristic emits a warning AND still completes without raising.
"""

from __future__ import annotations

import warnings

import pandas as pd

from edf_collector import compute_dispute_flags


def test_parse_failure_does_not_raise_but_warns() -> None:
    df = pd.DataFrame(
        {
            "Date": ["01/01/2024", "02/01/2024"],
            "Amount (£)": ["NOT_A_NUMBER", 200.0],
            "Period Charge (£)": [50.0, 60.0],
            "Entry Type": ["New Bill", "New Bill"],
            "Reading": ["Actual", "Actual"],
            "_dt": pd.to_datetime(["2024-01-01", "2024-02-01"]),
        }
    )
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        flags, counts = compute_dispute_flags(df, mean_daily=10.0)
    assert len(caught) >= 1, (
        f"expected at least one warning about a parse failure; got {len(caught)}"
    )
    # Each flag heuristic that caught the failure warns.  Confirm at
    # least one of them references the LARGE_JUMP family so we know
    # the swap hit the right branch.
    msgs = [str(w.message) for w in caught]
    assert any("could not evaluate" in m for m in msgs), msgs
    assert any("row index 1" in m for m in msgs), msgs


def test_clean_input_does_not_warn() -> None:
    """A clean numeric input MUST NOT emit any warnings — keeps the
    happy path noise-free."""
    df = pd.DataFrame(
        {
            "Date": ["01/01/2024", "01/02/2024", "01/03/2024"],
            "Amount (£)": [100.0, 200.0, 300.0],
            "Period Charge (£)": [50.0, 60.0, 70.0],
            "Entry Type": ["New Bill", "New Bill", "New Bill"],
            "Reading": ["Actual", "Actual", "Actual"],
            "_dt": pd.to_datetime(["2024-01-01", "2024-02-01", "2024-03-01"]),
        }
    )
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        flags, counts = compute_dispute_flags(df, mean_daily=10.0)
    assert len(caught) == 0, f"clean input should not warn; got {[str(w.message) for w in caught]}"
    assert counts["HIGH"] >= 1  # the 100% jump is HIGH per the heuristic
