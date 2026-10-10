"""Strict reads and atomic publication of immutable results."""

from pathlib import Path

from edf_bill_fetcher.domain.analysis import AnalysisSnapshot
from edf_bill_fetcher.domain.json_values import canonical_text, strict_loads
from edf_bill_fetcher.helpers.artifacts import atomic_output
from edf_bill_fetcher.io.persistence.result_codec import decode_result, encode_result


@atomic_output("path")
def save_result(result: AnalysisSnapshot, path: str | Path, *, result_id: str) -> None:
    """Publish beside an explicit destination, preserving previous bytes on failure."""
    Path(path).write_text(canonical_text(encode_result(result, result_id)) + "\n", encoding="utf-8")


def load_result(path: str | Path) -> AnalysisSnapshot:
    """Load strictly validated v3 analysis without recomputation."""
    return decode_result(strict_loads(Path(path).read_text(encoding="utf-8")))
