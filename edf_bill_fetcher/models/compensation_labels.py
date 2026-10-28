"""Consistent public wording for potential adjustments, separate from reviewed claims."""

CATEGORY_LABELS = {
    "back_billing_excess": "Potential back-billing adjustment",
    "credit_hold_interest": "Credit-hold interest scenario",
    "late_credit_interest": "Late-credit interest scenario",
}
TOTAL_LABEL = "Total potential adjustments (not a reviewed claim)"
NOTICE = (
    "These automatic scenarios are not a reconciled complaint reduction or an award of compensation. "
    "Check gross replacement charges, cancellations, earlier demands and VAT, then use "
    "Dispute ledger and Reconciled claims to record a source-backed proposal."
)


def category_label(value: object) -> str:
    """Render a readable label without exposing an internal category key."""
    text = str(value)
    return CATEGORY_LABELS.get(text, text.replace("_", " ").capitalize())
