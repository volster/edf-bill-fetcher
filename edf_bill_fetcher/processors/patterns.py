"""Pre-compiled regex patterns used by the amount/reading/period extractors and the multi-regex fallback chain.

Extracted from ``edf_collector.py`` as part of the modularization refactor
(Task 3).  This module is the single source of truth for:

- ``AMOUNT_PATTERNS`` — ordered list of (name, regex) tuples for amount
  extraction; the classifier drives off the *name*, not the list index.
- ``_AMOUNT_PATTERN_NEW_BILL`` / ``_AMOUNT_PATTERN_ONGOING_BALANCE`` —
  pattern-name → entry-type buckets consumed by the classifier.
- ``READING_PATTERNS`` — dict of reading-type → regex; callers iterate
  insertion-order and break on first match.
- ``PERIOD_RE`` — generic period-from / period-to regex.
- ``_POUND_AMOUNT_FALLBACK_RE`` — the "large amount" fallback used by
  ``extract_amount`` and the multi-regex fallback chain.
- ``_COVER_BLOCK_INV_RE`` / ``_COVER_BLOCK_PERIOD_RE`` — cover-body
  fallbacks for invoice number and billing period.
- ``_FALLBACK_INV_RE`` / ``_FALLBACK_AMOUNT_RE`` — last-resort loose
  fallbacks (``_FALLBACK_AMOUNT_RE`` is an alias of
  ``_POUND_AMOUNT_FALLBACK_RE``).

Dependency regexes consumed by ``processors/extraction.py`` also live
here so the package is self-contained (no circular import back into
``edf_collector``):

- ``_INV_NUMBER_RE`` / ``_CREDIT_NUMBER_RE`` — canonical KI / KCR
  invoice-number markers.
- ``_BILLING_PERIOD_RE`` — canonical "Your charges: <from> - <to>" marker.
- ``_PERIOD_CHARGE_RE`` / ``_CREDIT_TOTAL_RE`` — canonical period-charge
  and credit-total markers.
- ``_FROM_HEADER_RE`` / ``_EMAIL_ADDR_RE`` — sender-email extraction.
- ``_PST_PR_ATTACH_LONG_FILENAME`` — MAPI tag constant for the PST
  attachment long-filename walker.

All constants are module-level and fully type-annotated.
"""

from __future__ import annotations

import re

# ---------------------------------------------------------------------------
# Amount-extraction regexes
# ---------------------------------------------------------------------------
#
# Each entry is a (name, regex) tuple. The name maps to a single
# ``Entry Type`` ("New Bill", "Ongoing Balance", ...) — this is the
# contract the classifier consumes. Order matters: earlier entries
# match first, so the more specific patterns (``current_balance_debit``,
# ``total_charges_period``) take priority over the generic fall-through
# ones (``balance_within``).
#
# Adding/removing/reordering a pattern no longer breaks the classifier —
# the classifier drives off the *name* (a stable string), not off the
# list index. If you add a new pattern you must also pick its entry
# type in :data:`AMOUNT_PATTERN_ENTRY_TYPE` so the classifier routes
# it correctly.
AMOUNT_PATTERNS: list[tuple[str, re.Pattern[str]]] = [
    # New-style KI / KCR invoices — "Current balance £X debit"
    (
        "current_balance_debit",
        re.compile(r"current balance\s+£\s?([\d,]+(?:\.\d{2})?)\s*(?:in\s+)?debit", re.IGNORECASE),
    ),
    # New-style KI — "Total charges for this period £X debit"
    (
        "total_charges_period",
        re.compile(
            r"total charges for this period\s+£\s?([\d,]+(?:\.\d{2})?)\s*(?:in\s+)?debit",
            re.IGNORECASE,
        ),
    ),
    # New-style KCR — "Total credits for this bill £X"
    (
        "total_credits_bill",
        re.compile(r"total credits for this bill\s+£\s?([\d,]+(?:\.\d{2})?)", re.IGNORECASE),
    ),
    # Old-style cumulative balance
    (
        "your_new_account_balance",
        re.compile(r"your new account balance\s+£\s?([\d,]+(?:\.\d{2})?)", re.IGNORECASE),
    ),
    # Generic anchors (in priority order — more specific first)
    ("balance_within", re.compile(r"balance[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)", re.IGNORECASE)),
    (
        "total_charges_within",
        re.compile(r"total charges[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)", re.IGNORECASE),
    ),
    (
        "total_amount_due_within",
        re.compile(r"total amount due[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)", re.IGNORECASE),
    ),
    (
        "amount_to_pay_within",
        re.compile(r"amount to pay[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)", re.IGNORECASE),
    ),
    (
        "pound_amount_debit",
        re.compile(r"£\s?([\d,]+(?:\.\d{2})?)\s*(?:in\s+)?debit", re.IGNORECASE),
    ),
    (
        "current_balance_within",
        re.compile(r"current balance[\s\S]{0,30}?£\s?([\d,]+(?:\.\d{2})?)", re.IGNORECASE),
    ),
]
#
# Pattern name → entry type. The classifier looks up a pattern's name
# here to decide whether a match is "New Bill" or "Ongoing Balance".
# Unknown pattern names fall through to the heuristic body of
# :py:meth:`EvidenceEngine._classify_entry_type`.
_AMOUNT_PATTERN_NEW_BILL: frozenset[str] = frozenset(
    {
        "current_balance_debit",
        "total_charges_period",
        "total_credits_bill",
        "total_charges_within",
        "total_amount_due_within",
        "amount_to_pay_within",
        "pound_amount_debit",
    }
)
_AMOUNT_PATTERN_ONGOING_BALANCE: frozenset[str] = frozenset(
    {
        "your_new_account_balance",
        "balance_within",
        "current_balance_within",
    }
)
for name, _ in AMOUNT_PATTERNS:
    assert name in _AMOUNT_PATTERN_NEW_BILL or name in _AMOUNT_PATTERN_ONGOING_BALANCE, (
        f"AMOUNT_PATTERNS entry {name!r} has no entry-type bucket — "
        "add it to either _AMOUNT_PATTERN_NEW_BILL or _AMOUNT_PATTERN_ONGOING_BALANCE."
    )


READING_PATTERNS: dict[str, re.Pattern[str]] = {
    # Order matters: callers iterate insertion-order and break on first
    # match (see EvidenceEngine.process_text / _process_new_invoice).
    # Specific phrases first; the bare word "actual" only matches if
    # nothing more specific does.
    "Estimated": re.compile(r"estimated|est\.|estimate", re.IGNORECASE),
    "Smart": re.compile(r"smart meter|automated reading|smart reading", re.IGNORECASE),
    # "Actual" must NOT match the bare word "actual" — that phrase
    # appears in normal bill prose ("the actual amount you owe is
    # £X"). It only counts when the meter-reading language is present.
    "Actual": re.compile(
        r"actual reading|customer reading|your reading|"  # standard phrase
        r"reading was actual|reading is actual|"  # OK in long prose
        r"actual\s+reading\s*[-:]\s*\d|"
        r"meter\s+reading\s+was\s+actual",
        re.IGNORECASE,
    ),
}


PERIOD_RE = re.compile(
    r"(\d{1,2}(?:\s+\w+\s+\d{4}|\s*/\s*\d{2}\s*/\s*\d{4}|\s*-\s*\d{2}\s*-\s*\d{4}))"
    r"\s*(?:to|to:|–|-)\s*"
    r"(\d{1,2}(?:\s+\w+\s+\d{4}|\s*/\s*\d{2}\s*/\s*\d{4}|\s*-\s*\d{2}\s*-\s*\d{4}))",
    re.IGNORECASE,
)


# Used by the "large amount" fallback in `extract_amount`.  Pre-compiled
# once at module load; this hot-path is hit once per analysed chunk.
_POUND_AMOUNT_FALLBACK_RE = re.compile(r"£\s?(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)")

# =============================================================================
# KI / KCR invoice field regexes — pre-compiled at module load so
# `extract_new_invoice_fields` / `extract_new_credit_fields` don't pay the
# implicit re-compile cost on every PDF page.  All flags are baked into
# the compiled patterns (re.search refuses to combine a flags argument
# with an already-compiled pattern).
# =============================================================================
_INV_NUMBER_RE = re.compile(r"Invoice number:\s*(KI-[\w-]+)", re.IGNORECASE)
_BILLING_PERIOD_RE = re.compile(
    r"Your charges:\s*(\d{1,2}\s+\w+\s+\d{4})\s*[-–]\s*(\d{1,2}\s+\w+\s+\d{4})",
    re.IGNORECASE,
)
_PERIOD_CHARGE_RE = re.compile(
    r"Total charges for this period\s+£([\d,]+\.\d{2})(?:\s+(debit|credit))?",
    re.IGNORECASE,
)
_CREDIT_NUMBER_RE = re.compile(r"Credit note number:\s*(KCR-[\w-]+)", re.IGNORECASE)
# Credit-note accounts can use the same rendering as KI invoices.
_CREDIT_TOTAL_RE = re.compile(r"Total credits for this bill\s+£([\d,]+\.\d{2})", re.IGNORECASE)

# Multi-regex fallback chain (Task 5) -- used when `_INV_NUMBER_RE` /
# `_BILLING_PERIOD_RE` / `_PERIOD_CHARGE_RE` miss. Each fallback returns a
# (value, regex_name) tuple so the Source Excerpt column can show a
# regex-trace ("inv_num via _COVER_BLOCK_INV_RE; period via ...") per spec
# Stream P3.
_COVER_BLOCK_INV_RE = re.compile(r"Invoice\s+number:?\s*([A-Z0-9-]+)", re.IGNORECASE)
_COVER_BLOCK_PERIOD_RE = re.compile(
    r"(?:for\s+the\s+period|covering|bill\s+period)\s*[:]?\s*"
    r"(\d{1,2}\s+\w+\s+\d{4})\s*(?:-|to|--)\s*"
    r"(\d{1,2}\s+\w+\s+\d{4})",
    re.IGNORECASE,
)
_FALLBACK_INV_RE = re.compile(r"\b((?:KI|KCR|T\d{7}|A-\d{8})-\d{3,})\b")
_FALLBACK_AMOUNT_RE = _POUND_AMOUNT_FALLBACK_RE

# `_extract_sender_email` pulls an email out of either the transport
# headers (multi-line From:) or the sender name.  Compile both.
_FROM_HEADER_RE = re.compile(
    r"^From:\s*.*?([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})",
    re.MULTILINE | re.IGNORECASE,
)
_EMAIL_ADDR_RE = re.compile(r"([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})")

# MAPI tag constants from [MS-OXPROPS].
_PST_PR_ATTACH_LONG_FILENAME = 0x3707


__all__ = [
    "AMOUNT_PATTERNS",
    "READING_PATTERNS",
    "_AMOUNT_PATTERN_NEW_BILL",
    "_AMOUNT_PATTERN_ONGOING_BALANCE",
    "_COVER_BLOCK_INV_RE",
    "_COVER_BLOCK_PERIOD_RE",
    "_FALLBACK_AMOUNT_RE",
    "_FALLBACK_INV_RE",
    "PERIOD_RE",
    # Dependency regexes consumed by processors.extraction
    "_INV_NUMBER_RE",
    "_CREDIT_NUMBER_RE",
    "_BILLING_PERIOD_RE",
    "_PERIOD_CHARGE_RE",
    "_CREDIT_TOTAL_RE",
    "_POUND_AMOUNT_FALLBACK_RE",
    "_FROM_HEADER_RE",
    "_EMAIL_ADDR_RE",
    "_PST_PR_ATTACH_LONG_FILENAME",
]
