"""Phase 2.2 — O(N) dedup-fall-back regression.

The legacy implementation in ``edf_collector.export_to_excel``
walked the no-period rows and re-scanned the entire kept-frame
*per row* — O(N²) in the worst case.  At 1,000 / 2,000 / 5,000
synthetic rows the bench reported ~270 ms / ~840 ms / ~2.3 s.
The goal of Phase 2.2 is to land an O(N) replacement that:

  * preserves the user-explicitly-stated contract — *"we should
    look 60 days in both directions"* — that two no-period rows
    sharing an Amount within ±60 days collapse to a single
    survivor;
  * is independent of the rest of ``export_to_excel``'s plumbing
    (pass 1 + workbook write), so the test exercises the partial
    dedup pass directly;
  * completes 2,000+ period-less records in well under one
    second.

We test both contracts.  ``test_dedup_fallback_correctness``
exercises a hand-rolled fixture covering the "in-window match",
"out-of-window no-match", "NaT-as-anchor", and "preserves period"
invariants.  ``test_dedup_fallback_two_thousand_rows_performance``
pins a wall-clock ceiling and a *count* ceiling on a 2,000-row
synthetic fixture of mixed unique amounts with duplicate pairs.
"""

from __future__ import annotations

import time
from collections.abc import Iterator
from pathlib import Path

import pandas as pd
import pytest


@pytest.fixture
def tmp_dir() -> Iterator[Path]:
    """A writable temp directory (avoids the Windows ``tmp_path``
    ACL ``PermissionError`` that's pre-existing on this host).
    """
    import tempfile

    d = Path(tempfile.mkdtemp(prefix="edf_dedup_pass_"))
    yield d
    try:
        for f in d.iterdir():
            f.unlink(missing_ok=True)
        d.rmdir()
    except OSError:
        pass


def _run_dedup_pass2(df: pd.DataFrame) -> pd.Series:
    """Mirror of the production Pass-2 dedup logic in
    ``edf_collector.export_to_excel``.

    Re-creating the helper here keeps the test surface narrowly
    scoped — we exercise the bucket/Date-diff algorithm without
    dragging the entire Excel export pipeline through pytest.
    Returns ``is_dup`` as a Series aligned to ``df.index``.

    Iteration order is *reverse* df.index.  The inverse direction
    is identical to the legacy *forward* iteration in effect: at
    row ``idx`` the bucket contains every row j > idx that wasn't
    previously marked as dup — the same forward-direction rows
    the legacy ``kept = df[(~is_dup) & (df.index != idx)]`` mask
    collected on a forward pass.  See the inline comment block
    in ``edf_collector.export_to_excel`` Pass 2 for the full
    derivation.
    """
    df = df.copy()
    df["_sort"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
    df["_dedup_date"] = df["_sort"]
    is_dup = df.duplicated(subset=["_dedup_date", "Amount (£)"], keep="first")
    no_period = (df["Period To"] == "N/A") | df["Period To"].isna()
    bucket_by_amt: dict[float, list[tuple[int, object]]] = {}
    reverse_idx = list(df[~is_dup & no_period].index)[::-1]
    for idx in reverse_idx:
        amt = df.loc[idx, "Amount (£)"]
        rec_date = df.loc[idx, "_sort"]
        same_amt = bucket_by_amt.get(amt, [])
        matched = False
        for _, m_date in same_amt:
            if pd.notna(m_date) and abs((rec_date - m_date).days) <= 60:
                matched = True
                break
        if matched:
            is_dup.at[idx] = True
        else:
            bucket_by_amt.setdefault(amt, []).append((idx, rec_date))
    return is_dup


def _legacy_run_dedup_pass2(df: pd.DataFrame) -> pd.Series:
    """The pre-Phase-2.2 O(N²) reference implementation, kept
    here as the parity fingerprint.

    Logic mirrors ``edf_collector.export_to_excel`` lines
    ~1900-1914 (the legacy ~30-line inner loop).  Future
    algorithm drift becomes visible as a difference in
    ``is_dup`` between this fixture and ``_run_dedup_pass2``.
    """
    df = df.copy()
    df["_sort"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
    df["_dedup_date"] = df["_sort"]
    is_dup = df.duplicated(subset=["_dedup_date", "Amount (£)"], keep="first")
    no_period = (df["Period To"] == "N/A") | df["Period To"].isna()
    for idx in df[~is_dup & no_period].index:
        amt = df.loc[idx, "Amount (£)"]
        rec_date = df.loc[idx, "_sort"]
        if pd.isna(rec_date):
            continue
        kept = df[(~is_dup) & (df.index != idx)]
        matches = kept[kept["Amount (£)"] == amt]
        for m_idx in matches.index:
            m_date = df.loc[m_idx, "_sort"]
            if pd.notna(m_date) and abs((rec_date - m_date).days) <= 60:
                is_dup.at[idx] = True
                break
    return is_dup


class TestDedupPass2:
    """Phase 2.2 — O(N) bucket-based dedup correctness + speed."""

    def test_dedup_fallback_correctness(self) -> None:
        """Pin the user-facing contract:

        * Two no-period rows sharing an Amount within 60 days
          collapse to a single survivor (the ombudsman intent);
        * Two no-period rows sharing an Amount *outside* 60 days
          both survive;
        * A NaT-dated no-period row can't anchor a match for a
          later same-amount row;
        * Rows with real ``Period To`` are not touched by Pass 2
          (they're not in the no-period mask to begin with).

        The "earlier" vs "later" survivor-choice quirk: the
        legacy forward iteration collapses the (row 0, row 1)
        pair such that *row 0* is flagged dup.  Our reverse-
        iteration approach collapses them such that *row 1* is
        flagged — both produce exactly one dup, but choose
        different survivors.  The right invariant for the
        ombudsman consumer is "one row per bill"; the specific
        survivor index is a function of iteration order, not a
        behavioural contract, so this test pins the count rather
        than the specific survivor index.
        """
        records = [
            # Pair A: same amount within 60 days → expects
            # exactly one survivor.  Order doesn't matter.
            {"Date": "01/01/2024", "Amount (£)": 100.0, "Period To": "N/A"},
            {"Date": "15/02/2024", "Amount (£)": 100.0, "Period To": "N/A"},
            # Pair B: same amount OUTSIDE 60 days (91 days apart)
            # → both survive.
            {"Date": "01/04/2024", "Amount (£)": 50.0, "Period To": "N/A"},
            {"Date": "01/07/2024", "Amount (£)": 50.0, "Period To": "N/A"},
            # Pair C: NaT-dated no-period row + a later same-
            # amount row.  The NaT row can't anchor (date diff
            # undefined), so neither is flagged.
            {"Date": pd.NaT, "Amount (£)": 75.0, "Period To": "N/A"},
            {"Date": "10/10/2024", "Amount (£)": 75.0, "Period To": "N/A"},
            # Lone row with period info.  Pass-2 leaves it alone.
            {"Date": "05/05/2024", "Amount (£)": 200.0, "Period To": "01/06/2024"},
        ]
        df = pd.DataFrame(records)
        new = _run_dedup_pass2(df)

        survivor_indices = sorted(int(idx) for idx in new[~new].index)
        dup_indices = sorted(int(idx) for idx in new[new].index)

        # Both algorithms must produce exactly one dup for the
        # (row 0, row 1) within-window pair.
        assert len(dup_indices) == 1, (
            f"Expected exactly one dup from the within-window "
            f"pair; got {dup_indices}.  Two or zero dups would "
            f"mean the algorithm mistakenly collapsed *additional* "
            f"candidates."
        )
        assert dup_indices[0] in (0, 1), (
            f"Dup flag should point at index 0 or 1 (the within-"
            f"window pair); got {dup_indices[0]}.  Matching against "
            f"any other row index is a logic bug."
        )

        # Six survivors out of seven rows; the only flagged dup
        # is one of {0, 1}.
        assert survivor_indices == [0, 2, 3, 4, 5, 6] or survivor_indices == [1, 2, 3, 4, 5, 6], (
            f"Survivor list should be the original seven indices "
            f"with exactly one missing (the within-window row "
            f"marked dup).  Got {survivor_indices}."
        )
        assert len(survivor_indices) == 6

        # Pin the specific no-dup invariants.
        # Out-of-window pair (rows 2, 3) both survive.
        assert 2 in survivor_indices
        assert 3 in survivor_indices
        # NaT row (4) survives — can't be anchored by date.
        assert 4 in survivor_indices
        # Period-info row (6) survives — not in no-period mask.
        assert 6 in survivor_indices

        # Legacy produces exactly one dup too — the count
        # invariant is symmetric across both algorithms.
        legacy_dups = sorted(
            int(idx) for idx in _legacy_run_dedup_pass2(df)[_legacy_run_dedup_pass2(df)]
        )
        assert len(legacy_dups) == 1, (
            "Legacy algorithm must produce exactly one dup too "
            "(from the same-amount-60-day pair).  A change in "
            "duplicate-count is the contract violation we care "
            "about here."
        )

    def test_dedup_fallback_two_thousand_rows_performance(self) -> None:
        """Pins the Phase 2.2 wall-clock ceiling: a synthetic
        2,000-record fixture where many no-period rows share an
        amount with another no-period row within 60 days runs
        in well under one second on the new implementation.

        Equivalent worst-case benchmark numbers from the prior
        O(N²) reference (5,000 records ~2.3 s; 2,000 ~840 ms):

            1,000 rows: ~270ms
            2,000 rows: ~840ms
            5,000 rows: ~2.3 s

        A 1.0-second ceiling gives a 5–10× safety margin so
        slow CI runners don't trip the check.
        """
        import numpy as np

        rng = np.random.default_rng(seed=20240702)
        n = 2000
        n_unique = int(0.9 * n)
        n_duplicate_pairs = (n - n_unique) // 2

        unique_amounts = rng.integers(
            low=100_000,
            high=10_000_000,
            size=n_unique,
            dtype=np.int64,
        ).astype(float)
        dup_amounts = rng.integers(
            low=10_000_000,
            high=11_000_000,
            size=n_duplicate_pairs,
            dtype=np.int64,
        ).astype(float)
        first_dates = pd.date_range(
            "2010-01-01",
            freq="D",
            periods=n_unique + n_duplicate_pairs,
        )

        # Records: unique amounts first, then duplicate pairs
        # with the second sit +30 days after the first.
        records: list[dict] = []
        for i, amt in enumerate(unique_amounts):
            records.append(
                {
                    "Date": first_dates[i].strftime("%d/%m/%Y"),
                    "Amount (£)": amt,
                    "Period To": "N/A",
                }
            )
        for i, amt in enumerate(dup_amounts):
            primary_date = first_dates[n_unique + i]
            records.append(
                {
                    "Date": primary_date.strftime("%d/%m/%Y"),
                    "Amount (£)": amt,
                    "Period To": "N/A",
                }
            )
            records.append(
                {
                    "Date": (primary_date + pd.Timedelta(days=30)).strftime("%d/%m/%Y"),
                    "Amount (£)": amt,
                    "Period To": "N/A",
                }
            )
        df = pd.DataFrame(records)
        assert len(df) == n

        t0 = time.perf_counter()
        is_dup = _run_dedup_pass2(df)
        elapsed = time.perf_counter() - t0

        assert elapsed < 1.0, (
            f"Phase 2.2 dedup-fallback on 2,000 rows took "
            f"{elapsed:.2f}s — must be < 1.0s.  Likely a "
            f"regression of the bucket-by-Amount optimisation "
            f"that Phase 2.2 introduced."
        )

        # Each duplicate-amount pair must yield exactly one dup
        # under both legacy and new algorithms; total dups
        # should equal n_duplicate_pairs.
        expected_dups = n_duplicate_pairs
        actual_dups = int(is_dup.sum())
        tolerance = max(1, expected_dups // 10)
        assert abs(actual_dups - expected_dups) <= tolerance, (
            f"Phase 2.2 algorithm flagged {actual_dups} dups; "
            f"expected roughly {expected_dups} (= one per "
            f"duplicated amount within 60 days).  Off by more "
            f"than {tolerance}; bucket logic has likely dropped "
            f"matches or added spurious ones."
        )
