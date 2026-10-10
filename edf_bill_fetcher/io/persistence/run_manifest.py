"""Atomic stage manifests with verifiable artifact identities."""

import hashlib
from pathlib import Path

from edf_bill_fetcher.domain.json_values import canonical_text, strict_loads
from edf_bill_fetcher.helpers.artifacts import atomic_output

TERMINAL = frozenset({"complete", "review_required", "failed", "cancelled", "empty"})


def artifact_record(path: str) -> dict:
    """Record bytes actually published at an explicit artifact destination."""
    source = Path(path)
    digest = hashlib.sha256()
    with source.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return {
        "path": str(source.absolute()),
        "sha256": digest.hexdigest(),
        "bytes": source.stat().st_size,
    }


@atomic_output("path")
def save_manifest(manifest: dict, path: str) -> None:
    """Publish stage state only after validating its terminal outcome."""
    if manifest.get("status") in {"complete", "review_required"} and (
        manifest.get("error")
        or set(manifest.get("requested", {})) - set(manifest.get("artifacts", {}))
    ):
        raise ValueError("Required artifacts are missing; run cannot be complete")
    Path(path).write_text(canonical_text(manifest) + "\n", encoding="utf-8")


def load_manifest(path: str) -> dict:
    """Recognise a saved nonterminal stage as an interrupted run."""
    manifest = strict_loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(manifest, dict) or manifest.get("manifest_version") != 1:
        raise ValueError("Unsupported run manifest")
    if manifest.get("status") not in TERMINAL:
        manifest["status"] = "interrupted"
    return manifest


def is_run_directory(path: str) -> bool:
    """Recognise this application's output folders before recursive source scanning."""
    try:
        manifest = load_manifest(str(Path(path) / "run.json"))
    except (OSError, ValueError):
        return False
    requested = manifest.get("requested")
    return (
        type(manifest.get("manifest_version")) is int
        and isinstance(manifest.get("run_id"), str)
        and bool(manifest["run_id"])
        and isinstance(requested, dict)
        and bool(requested)
        and set(requested) <= {"xlsx", "pdf", "docx", "html"}
        and isinstance(manifest.get("artifacts"), dict)
    )
