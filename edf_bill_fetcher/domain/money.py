"""Explicit decimal monetary boundaries without financial-role inference."""

from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal, InvalidOperation


def parse_money(value: str | int | float | Decimal | None) -> Decimal | None:
    """Read finite posted money, round half up to pence, and retain missingness."""
    if value is None or isinstance(value, bool):
        return None
    try:
        text = str(value).strip().replace("£", "").replace(",", "")
        parsed = Decimal(text)
        if not parsed.is_finite():
            return None
        result = parsed.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        return Decimal("0.00") if result == 0 else result
    except (InvalidOperation, ValueError):
        return None


def money_text(value: Decimal | None) -> str | None:
    """Format known finite money canonically; unavailable values stay null."""
    if value is None:
        return None
    parsed = parse_money(value)
    if parsed is None:
        raise ValueError("Cannot format nonfinite money")
    return format(parsed, ".2f")
