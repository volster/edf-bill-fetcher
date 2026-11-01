"""Bounded transport-independent draft history for HTTP requests and CLI invocations."""

from __future__ import annotations

from edf_bill_fetcher.application.review_commands import apply_command
from edf_bill_fetcher.application.review_session import ReviewSession
from edf_bill_fetcher.domain.analysis import AnalysisSnapshot


class ReviewJournal:
    """Restore every selected state through the existing full-result draft binding."""

    def __init__(self, snapshot: AnalysisSnapshot, history: list[dict], position: int):
        """Reject malformed histories without modifying a previously saved draft."""
        if (
            not isinstance(history, list)
            or not 1 <= len(history) <= 101
            or type(position) is not int
            or not 0 <= position < len(history)
        ):
            raise ValueError("Invalid review history")
        self.history = list(history)
        self.position = position
        self.session = ReviewSession.restore(snapshot, history[position])

    def apply(self, operation: str, arguments: dict) -> None:
        """Apply one shared command or move to a validated earlier/later snapshot."""
        if operation in {"undo", "redo"}:
            if arguments:
                raise ValueError("Undo/redo does not accept arguments")
            position = self.position + (-1 if operation == "undo" else 1)
            if not 0 <= position < len(self.history):
                raise ValueError("No change available to undo or redo")
            session = ReviewSession.restore(self.session.snapshot, self.history[position])
            self.position, self.session = position, session
        else:
            apply_command(self.session, operation, arguments)
            self.history = [*self.history[: self.position + 1], self.session.payload()][-101:]
            self.position = len(self.history) - 1
