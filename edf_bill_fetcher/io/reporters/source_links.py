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
    root = Path(bundle_dir).resolve()
    try:
        files = json.loads((root / "evidence_manifest.json").read_text(encoding="utf-8"))["files"]
    except (OSError, ValueError, KeyError, TypeError):
        return {}
    links: dict[str, SourceLink] = {}
    ambiguous: set[str] = set()
    if not isinstance(files, list):
        return links
    for entry in files:
        if not isinstance(entry, dict) or entry.get("status") != "saved":
            continue
        path = entry.get("saved_path")
        if not isinstance(path, str) or "\\" in path or ":" in path:
            continue
        relative = PurePosixPath(path)
        if relative.is_absolute() or ".." in relative.parts:
            continue
        target = (root / path).resolve()
        if root not in target.parents or not target.is_file():
            continue
        expected = entry.get("sha256")
        if isinstance(expected, str) and expected:
            try:
                if hashlib.sha256(target.read_bytes()).hexdigest() != expected:
                    continue
            except OSError:
                continue
        link = SourceLink(quote(path, safe="/"), "Open source")
        for key in (entry.get("source_id"), entry.get("attachment_name")):
            if isinstance(key, str) and key:
                if key in links and links[key] != link:
                    ambiguous.add(key)
                links[key] = link
    return {key: link for key, link in links.items() if key not in ambiguous}


def record_source_link(row: dict, links: dict[str, SourceLink]) -> SourceLink | None:
    """Use a preserved source identity and original page, never a guessed filename."""
    link = links.get(str(row.get("Source ID", ""))) or links.get(
        str(row.get("Attachment Name", ""))
    )
    if link is None:
        return None
    page = row.get("Page From")
    if isinstance(page, (int, float)) and not isinstance(page, bool) and 0 < page < 1000000:
        return SourceLink(link.href + f"#page={int(page)}", link.label)
    return link
