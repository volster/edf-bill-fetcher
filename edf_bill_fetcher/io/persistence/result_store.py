"""Strict reads and atomic publication of immutable results."""

from pathlib import Path

from edf_bill_fetcher.domain.analysis import AnalysisSnapshot
from edf_bill_fetcher.domain.json_values import canonical_text, strict_loads
from edf_bill_fetcher.domain.requests import RunRequest
from edf_bill_fetcher.helpers.artifacts import atomic_output
from edf_bill_fetcher.io.persistence.result_codec import (
    decode_result,
    encode_result,
    migrate_legacy,
)


@atomic_output("path")
def save_result(result: AnalysisSnapshot, path: str | Path, *, result_id: str) -> None:
    """Publish beside an explicit destination, preserving previous bytes on failure."""
    Path(path).write_text(canonical_text(encode_result(result, result_id)) + "\n", encoding="utf-8")


def load_result(
    path: str | Path, *, reanalyse: bool = False, legacy_request: RunRequest | None = None
) -> AnalysisSnapshot:
    """Restore v3 as saved; legacy/reanalysis needs an explicit captured context."""
    payload = strict_loads(Path(path).read_text(encoding="utf-8"))
    if isinstance(payload, dict) and "schema_version" in payload:
        result = decode_result(payload)
        if not reanalyse:
            return result
        if legacy_request is None:
            raise ValueError("Explicit reanalysis requires a new captured request")
        payload = {
            **payload["evidence"]["legacy_state"],
            "version": 2,
            "records": payload["evidence"]["prepared_records"],
            "original_records": [row["raw"] for row in payload["evidence"]["observations"]],
        }
    return migrate_legacy(payload, legacy_request)
