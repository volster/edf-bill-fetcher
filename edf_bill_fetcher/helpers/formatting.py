"""Cell-formatting helpers shared across the evidence workbook.

These helpers were extracted from ``edf_collector.py`` as part of the
modularization refactor (Task 3).  They cover:

- ``apply_currency_format`` / ``apply_int_format`` — coerce a cell's
  value to a numeric type and pin its ``number_format`` so Excel
  renders it as currency / integer.
- ``account_number_matches`` — token-aware account-number filter
  predicate (Phase 1.3) that replaces the naive "digits substring"
  check.
"""

from __future__ import annotations

import re

import openpyxl.cell
import pandas as pd


def apply_currency_format(cell: openpyxl.cell.Cell) -> None:
    """Coerce cell value to float and apply a currency number format."""
    if isinstance(cell.value, str):
        cell.value = float(cell.value)
    cell.number_format = "\u00a3#,##0.00"


def apply_int_format(cell: openpyxl.cell.Cell) -> None:
    """Coerce cell value to int and apply an integer number format."""
    if isinstance(cell.value, str):
        cell.value = int(float(cell.value))
    cell.number_format = "#,##0"


def account_number_matches(acc_filter: str, text: str) -> bool:
    r"""Return True when ``acc_filter`` appears as a standalone digit run in ``text``.

    Phase 1.3: replaces the old "is the digits substring contained anywhere"
    check, which false-matched an unrelated longer string (invoice number,
    meter serial, phone number) when the configured account number happened
    to be a subset of those digits.

    Strategy
    --------
    Split ``text`` into "tokens" — sequences of alphanumeric characters and
    hyphens (e.g., "A-12345678", "31", "555", "4444").  For each token,
    strip non-digit characters to get its digit-only form.  Look up the
    normalized ``acc_filter`` in the resulting list.

    This preserves the natural word boundaries of ``"Account number:
    31 555 4444"`` — the tokens are ``["Account", "number", "31", "555",
    "4444", "Current", "balance", "240", "50", "debit"]`` and their
    digit-only forms are ``["", "", "31", "555", "4444", "", "", "240",
    "50", ""]``.  The filter ``"31"`` only matches the "31" token —
    whereas the naive ``digits_only in text_no_sep`` would have produced
    ``"315554444"`` from the same input and dropped the right answer.

    Crucially, hyphens *inside* a token are preserved during tokenization
    so ``"A-12345678"`` is one token whose digit-only form is
    ``"12345678"`` — fixing the false-negative where the original
    ``re.findall(r"\\d+", ...)`` split it into ``["123", "45678"]``.

    Invariant
    ---------
    Pure-substring match (digits in collapse-stripped text) being True
    does NOT imply this helper returns True (we tighten the
        predicate).  The reverse direction holds: a real standalone digit
        run from the original text, after the token-based split, still
        matches.  No legitimate standalone occurrence is dropped.
    """
    if not acc_filter:
        return True  # empty filter matches everything
    digits_only = re.sub(r"\D", "", str(acc_filter))
    if not digits_only:
        return True  # unusable filter — pass rather than silently reject

    # Tokenize: split on whitespace and punctuation EXCEPT hyphens that
    # are inside alphanumeric sequences.  The pattern [A-Za-z0-9-]+
    # captures words, numbers, and hyphenated codes like "A-12345678"
    # or "A123-456" as single tokens.
    tokens = re.findall(r"[A-Za-z0-9-]+", text or "")

    # For each token, strip non-digits to get its digit-only form.
    # Example: "A-12345678" → "12345678", "31" → "31", "555" → "555"
    token_digits = [re.sub(r"\D", "", t) for t in tokens]

    return digits_only in token_digits


__all__ = [
    "apply_currency_format",
    "apply_int_format",
    "account_number_matches",
    "_is_populated",
    "_amalgamate_cluster",
    "_apply_amalgamate_to_kept_frame",
]


def _is_populated(value: object) -> bool:
    """Return True iff ``value`` counts as a populated field for completeness scoring."""
    if value is None:
        return False
    try:
        if isinstance(value, float) and pd.isna(value):
            return False
    except (TypeError, ValueError):
        pass
    if isinstance(value, str):
        s = value.strip()
        return s != "" and s != "N/A"
    return True


def _amalgamate_cluster(cluster: pd.DataFrame) -> pd.DataFrame:
    """Merge a duplicate cluster into a single hybrid row."""
    if len(cluster) <= 1:
        return cluster.iloc[0:0]
    cluster = cluster.sort_values(
        ["_completeness", "_src_pri", "_sort"],
        ascending=[False, True, True],
    )
    hybrid: dict[str, object] = {}
    for col in cluster.columns:
        if col.startswith("_"):
            continue
        picked = None
        for _ri, row in cluster.iterrows():
            val = row.get(col)
            if col == "Source":
                picked = val
                break
            if _is_populated(val):
                picked = val
                break
        hybrid[col] = picked if picked is not None else row.get(col)
    return pd.DataFrame([hybrid], index=[cluster.index[0]])


def _apply_amalgamate_to_kept_frame(
    df: pd.DataFrame,
    dup_df: pd.DataFrame,
    kept_pass1_index: dict[tuple, int],
    kept_for_dup: dict[int, int],
    is_dup: pd.Series,
) -> pd.DataFrame:
    """Reflow ``df`` so each duplicate cluster collapses to a single hybrid kept row."""
    anchor_to_dup_indices: dict[int, list[int]] = {}
    for (_dd, _amt), kept_idx in kept_pass1_index.items():
        anchor_to_dup_indices.setdefault(kept_idx, [])
    for dup_idx, kept_idx in kept_for_dup.items():
        anchor_to_dup_indices.setdefault(kept_idx, []).append(dup_idx)

    hybrid_rows: list[pd.DataFrame] = []
    for anchor_idx, dup_indices in anchor_to_dup_indices.items():
        if not dup_indices:
            continue
        one_cluster = dup_df.loc[dup_indices]
        cluster_df = pd.concat([df.loc[[anchor_idx]], one_cluster])
        hybrid = _amalgamate_cluster(cluster_df)
        if not hybrid.empty:
            hybrid.index = [anchor_idx]
            hybrid_rows.append(hybrid)
    if not hybrid_rows:
        return df
    hybrid_idx_set = {h.index[0] for h in hybrid_rows}
    non_hybrid_kept = df.loc[df[~is_dup].index.difference(hybrid_idx_set)]
    return pd.concat([non_hybrid_kept] + hybrid_rows).reset_index(drop=True)
