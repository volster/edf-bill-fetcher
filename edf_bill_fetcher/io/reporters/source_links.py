"""Contained local evidence links shared by the document renderers."""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from urllib.parse import quote


@dataclass(frozen=True)
class SourceLink:
    """A local manifest path, never an arbitrary HTML string from evidence."""

    href: str
    label: str


def source_links(bundle_dir: str | None) -> dict[str, SourceLink]:
    """Resolve only saved, contained manifest paths beside this report."""
    if not bundle_dir:
        return {}
    try:
        root = Path(bundle_dir).resolve()
        files = json.loads((root / "evidence_manifest.json").read_text(encoding="utf-8"))["files"]
    except (OSError, RuntimeError, ValueError, KeyError, TypeError):
        return {}
    links: dict[str, SourceLink] = {}
    origins: dict[str, tuple[object, object]] = {}
    ambiguous: set[str] = set()
    if not isinstance(files, list):
        return links
    for entry in files:
        if not isinstance(entry, dict):
            continue
        link = _saved_link(root, entry)
        origin = (entry.get("source_id"), entry.get("saved_path"))
        # Unavailable occurrences still make a filename ambiguous. Otherwise a
        # missing original could silently be replaced by a same-named file.
        for key in (entry.get("source_id"), entry.get("attachment_name")):
            if isinstance(key, str) and key:
                if link is None or (key in origins and origins[key] != origin):
                    ambiguous.add(key)
                origins[key] = origin
                if link is not None:
                    links[key] = link
    return {key: link for key, link in links.items() if key not in ambiguous}


def _saved_link(root: Path, entry: dict) -> SourceLink | None:
    """Check one original without letting a filesystem failure hide other links."""
    if entry.get("status") != "saved":
        return None
    path = entry.get("saved_path")
    if not isinstance(path, str) or "\\" in path or ":" in path:
        return None
    relative = PurePosixPath(path)
    if relative.is_absolute() or ".." in relative.parts:
        return None
    try:
        target = (root / path).resolve()
        if root not in target.parents or not target.is_file():
            return None
        expected = entry.get("sha256")
        if isinstance(expected, str) and expected:
            if hashlib.sha256(target.read_bytes()).hexdigest() != expected:
                return None
    except (OSError, RuntimeError, ValueError):
        return None
    return SourceLink(quote(path, safe="/"), "Open source")


def record_source_link(row: dict, links: dict[str, SourceLink]) -> SourceLink | None:
    """Use a preserved source identity and original page, never a guessed filename."""
    identity = row.get("Source ID")
    link = links.get(str(identity)) if identity else links.get(str(row.get("Attachment Name", "")))
    if link is None:
        return None
    page = row.get("Page From")
    if (
        isinstance(page, (int, float))
        and not isinstance(page, bool)
        and 0 < page < 1000000
        and page == int(page)
    ):
        return SourceLink(link.href + f"#page={int(page)}", link.label)
    return link
