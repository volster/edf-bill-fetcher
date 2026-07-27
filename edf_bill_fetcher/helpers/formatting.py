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
    """Return True when ``acc_filter`` appears as a standalone digit run in ``text``.

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
]
