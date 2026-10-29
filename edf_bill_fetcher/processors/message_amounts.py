"""Explicit email amount contexts that do not imply a completed transaction."""

from __future__ import annotations

import re

_MONEY = r"£\s*([0-9]+(?:,[0-9]{3})*(?:\.[0-9]{2})?)(?![\d,]|\.\d)"
_DEMANDS = (
    (r"your bill for\s*" + _MONEY + r"\s+is overdue", "Overdue bill demand"),
    (
        r"the total outstanding amount of\s*" + _MONEY + r"\s+is now due(?: in full)?",
        "Outstanding amount demanded after plan cancellation",
    ),
)
_SUPPORTING = (
    (
        r"your monthly direct debit payment amount is\s*" + _MONEY,
        "Direct Debit schedule",
        "Scheduled Direct Debit (£)",
    ),
    (
        r"we'll collect your first (?:electricity|gas|energy) payment of\s*" + _MONEY,
        "Direct Debit schedule",
        "Scheduled first Direct Debit (£)",
    ),
    (
        r"your bank(?:\s*/\s*building society)? has(?:n't| not) paid us your latest direct debit of\s*"
        + _MONEY,
        "Failed Direct Debit notice",
        "Failed Direct Debit (£)",
    ),
)


def _stated_amount(text: str, pattern: str) -> float | None:
    """Exclude conditional advice and require a complete, finite currency token."""
    for match in re.finditer(pattern, text, re.I):
        prefix = re.split(r"[.!?;:]", text[max(0, match.start() - 100) : match.start()])[-1]
        if re.search(r"\b(?:if|whether|example)\b", prefix, re.I):
            continue
        amount = float(match[1].replace(",", ""))
        if amount < float("inf"):
            return amount
    return None


def message_amount_context(text: str, subject: str) -> dict | None:
    """Separate a stated debt from a quote or planned/failed collection.

    The caller limits this to email bodies and preserves stronger balance/bill
    anchors. Supporting amounts remain observations with source provenance;
    neither a collection plan nor a failed collection is a paid transaction.
    """
    text = re.sub(r"\s+", " ", text.replace("’", "'"))
    for pattern, basis in _DEMANDS:
        amount = _stated_amount(text, pattern)
        if amount is not None:
            return {"Amount (£)": amount, "Amount Kind": "Balance", "Demand Context": basis}
    patterns = list(_SUPPORTING)
    if re.search(r"\b(?:energy quote|energy quotation)\b", subject, re.I):
        patterns.append(
            (r"total annual cost\s*" + _MONEY, "Energy quote", "Quoted Annual Cost (£)")
        )
    for pattern, role, field in patterns:
        amount = _stated_amount(text, pattern)
        if amount is not None:
            return {
                "Amount (£)": amount,
                "Amount Kind": "Not monetary",
                "Entry Type": "Supporting Document",
                "Document Role": role,
                field: amount,
                "Amount Context": "Supporting amount; not an observed account balance or completed transaction",
            }
    return None
