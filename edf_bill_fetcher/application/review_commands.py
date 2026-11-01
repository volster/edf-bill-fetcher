"""Transport-independent review commands for desktop, automation and HTTP adapters."""

from __future__ import annotations

import inspect

from edf_bill_fetcher.application.review_session import ReviewSession

COMMANDS = (
    "update",
    "set_revision",
    "set_position",
    "clear_position",
    "set_claim",
    "clear_claim",
    "separate_claims",
    "clear_separation",
    "set_complaint_context",
    "undo",
    "redo",
)


def apply_command(session: ReviewSession, operation: str, arguments: dict) -> None:
    """Apply an explicitly supported command with the existing domain validation."""
    if operation not in COMMANDS or not isinstance(arguments, dict):
        raise ValueError("Unsupported review command")
    method = getattr(session, operation)
    try:
        inspect.signature(method).bind(**arguments)
    except TypeError as error:
        raise ValueError(f"Invalid {operation} arguments: {error}") from error
    method(**arguments)
