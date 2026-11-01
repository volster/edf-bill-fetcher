"""Read accepted results or records exports for identical comparison across interfaces."""

from pathlib import Path

from edf_bill_fetcher.domain.json_values import rows_value, strict_loads
from edf_bill_fetcher.io.persistence.result_codec import decode_result
from edf_bill_fetcher.processors.run_diff import diff_records


def comparison_records(payload: object) -> list[dict]:
    """Preserve accepted prepared rows without rerunning analytical models."""
    if isinstance(payload, dict) and "schema_version" in payload:
        return rows_value(decode_result(payload).prepared.records)
    records = payload.get("records") if isinstance(payload, dict) else payload
    if not isinstance(records, list) or any(not isinstance(row, dict) for row in records):
        raise ValueError("Comparison inputs must be accepted results or records.json lists")
    return records


def compare_files(old: str, new: str) -> dict:
    """Load explicit JSON inputs and use the existing pure comparison processor."""
    return diff_records(
        comparison_records(strict_loads(Path(old).read_text(encoding="utf-8"))),
        comparison_records(strict_loads(Path(new).read_text(encoding="utf-8"))),
    )
