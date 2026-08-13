"""Shared package-version resolution.

The evidence workbook and both report formats stamp the tool
version on their output (provenance sheet, cover page).  Resolving
it once here — from the repo-root ``pyproject.toml`` — keeps every
surface truthful and avoids each writer re-implementing the lookup.

Deliberately lives in ``helpers`` (the lowest shared layer) so the
writers layer can import it without a layering violation: the old
``_get_package_version`` in ``io/reporters/pdf_report.py`` resolved
``Path(__file__).parent / "pyproject.toml"`` — the reporter's own
directory, which never contains a pyproject.toml — so it silently
returned the ``"0.1.0"`` fallback for every run.  The bug was
invisible because the repo root also declared 0.1.0; the moment the
declared version diverged from the fallback, reports would stamp a
stale version.
"""

from __future__ import annotations

import re
from pathlib import Path

_DEFAULT_VERSION = "0.2.0"
_MODULE_DIR = Path(__file__).resolve().parent


def get_package_version() -> str:
    """Return ``[project] version`` from the repo-root pyproject.toml.

    Walks upward from this module's directory until a
    ``pyproject.toml`` containing a ``version`` line is found, and
    falls back to ``"0.1.0"`` if none can be read (wheeled installs,
    sandboxed runners, frozen builds where the manifest is absent).

    Walking upward (rather than assuming a fixed relative path) keeps
    the lookup correct whether the package is imported from a source
    checkout or from site-packages under a virtualenv, and avoids
    depending on the caller's working directory.
    """
    for parent in (_MODULE_DIR, *_MODULE_DIR.parents):
        candidate = parent / "pyproject.toml"
        if not candidate.is_file():
            continue
        try:
            text = candidate.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        m = re.search(r'^version\s*=\s*"([^"]+)"', text, re.MULTILINE)
        if m:
            return m.group(1)
    return _DEFAULT_VERSION
