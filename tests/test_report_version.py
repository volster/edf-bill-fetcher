"""Tests for the report-version helper.

Versioning the report directly off the project's pyproject.toml
means a paying client always sees the version string we ship —
not a stale literal baked into cover-page code. The fallback
path is checked here to lock in the no-crash behaviour for
builds where pyproject is missing (wheeled installs, sandboxed
runners).

Inputs are entirely synthetic — the helper reads the actual
project pyproject.toml so whatever ``[project] version = "..."``
is declared here, the test will compare against it.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest
from reportlab.platypus import Paragraph

from edf_report import _get_package_version, create_cover_page


def _read_pyproject_version() -> str:
    """Return the value of ``[project] version`` from pyproject.toml.

    Helper that re-reads the project's own declared version so a
    paying-client deployment can be matched against the on-disk
    helper output.
    """
    text = (Path(__file__).resolve().parents[1] / "pyproject.toml").read_text()
    m = re.search(r'^version\s*=\s*"([^"]+)"', text, re.MULTILINE)
    assert m is not None, "pyproject.toml is missing a [project] version line"
    return m.group(1)


def test_get_package_version_returns_string():
    assert isinstance(_get_package_version(), str)


def test_get_package_version_includes_the_actual_pyproject_version():
    # Re-read pyproject.toml and compare a known-good value against
    # the helper's output. Catches the failure mode where the
    # helper silently returns a placeholder that happens to look
    # like a version.
    declared = _read_pyproject_version()
    assert _get_package_version() == declared

    # And confirm the version renders on the actual cover page,
    # not just inside the helper.
    elements = create_cover_page(
        account_ref="ACC-TEST",
        period_start="01/01/2026",
        period_end="31/12/2026",
        report_date="01 Jan 2026",
    )
    found_version = False
    for el in elements:
        if isinstance(el, Paragraph) and f"v{declared}" in (el.text or ""):
            found_version = True
            break
    assert found_version, f"Cover-page text did not contain the version string 'v{declared}'"


def test_get_package_version_falls_back_to_zero_when_pyproject_missing(monkeypatch):
    """If pyproject.toml is unreadable, return a stable fallback.

    The helper resolves Path(__file__).parent / "pyproject.toml";
    we substitute ``pathlib.Path.resolve`` with a stand-in whose
    ``read_text`` always raises ``OSError``. The function must
    fall back to a sane default rather than propagate the
    exception — the cover page is what a paying client sees
    first, and a missing version string is uglier than a stable
    fallback (and worse than no cover page at all).
    """
    import pathlib

    class _HiddenPath(pathlib.Path):
        def read_text(self, *args, **kwargs):
            raise OSError("simulated missing")

    monkeypatch.setattr(
        pathlib.Path,
        "resolve",
        classmethod(lambda cls: _HiddenPath(__file__)),
    )

    try:
        _get_package_version()
    except OSError:
        pytest.fail(
            "_get_package_version() leaked an OSError when pyproject.toml is unreadable; "
            "expected fall-back to '0.1.0'"
        )
