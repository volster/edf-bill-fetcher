"""Shared run orchestration without UI state or implicit output destinations."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass

from edf_bill_fetcher.domain.analysis import AnalysisSnapshot
from edf_bill_fetcher.domain.json_values import FrozenJSON


class RunCancelled(Exception):
    """Cancellation observed at a stage boundary."""


@dataclass(frozen=True)
class ProgressEvent:
    """A stage update safe for GUI queue delivery."""

    stage: str
    message: str


@dataclass(frozen=True)
class RunOutcome:
    """Truthful terminal result, retaining valid analysis after renderer failure."""

    status: str
    snapshot: AnalysisSnapshot | None
    manifest: FrozenJSON
    error: str | None = None


def check_cancelled(cancelled: Callable[[], bool]) -> None:
    """Stop between stages without discarding already published artifacts."""
    if cancelled():
        raise RunCancelled("Run cancelled")
