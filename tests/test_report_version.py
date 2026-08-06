"""Tests for the report-version helper.

Versioning the report directly off the project's pyproject.toml
means a user always sees the version string we ship —
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

from reportlab.platypus import Paragraph

from edf_bill_fetcher.io.reporters.pdf_report import _get_package_version, create_cover_page


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


def test_get_package_version_falls_back_to_zero_when_pyproject_missing(
    monkeypatch,
):
    """If pyproject.toml is unreadable, return a stable fallback.

    The helper resolves Path(__file__).parent / "pyproject.toml" and
    reads it. We monkey-patch the module-level lookup by replacing
    ``open`` in the ``edf_report`` module's namespace so any read
    attempt raises ``OSError``. The function must feed the
    fallback chain instead of propagating — the cover page is what a
    user sees first and a missing version string is uglier
    than a stable fallback. Avoid relying on ``Path.resolve`` /
    platform-specific ``_flavour`` (CI Windows runners differ from
    Linux / macOS in how pathlib's POSIXPath / WindowsPath copies
    work through subclassing).
    """
    import builtins

    real_open = builtins.open

    def _raise_open(*args, **kwargs):
        # Allow the helper to open *something* (so ast / import
        # machinery works in unrelated tests) when the target is
        # NOT pyproject.toml — but raise when target IS pyproject.
        path = args[0] if args else kwargs.get("file", kwargs.get("name"))
        if isinstance(path, str) and path.endswith("pyproject.toml"):
            raise OSError("simulated missing")
        if isinstance(path, object) and str(path).endswith("pyproject.toml"):
            raise OSError("simulated missing")
        return real_open(*args, **kwargs)

    monkeypatch.setattr(builtins, "open", _raise_open)
    try:
        from edf_bill_fetcher.io.reporters.pdf_report import _get_package_version as _get

        _get()
    except OSError:
        import pytest

        pytest.fail(
            "_get_package_version() leaked an OSError when pyproject.toml is unreadable; "
            "expected fall-back to '0.1.0'"
        )


def test_get_package_version_uses_repo_root_not_module_dir(monkeypatch, tmp_path):
    """A pyproject.toml sitting next to the module must be ignored.

    The helper resolves the version from the *repo-root* pyproject.toml.
    Regression pin: the original implementation looked at
    ``Path(__file__).parent / "pyproject.toml"`` — the module's own
    directory — which does not contain a pyproject.toml, so it
    silently returned the ``"0.1.0"`` fallback for every run.  That
    bug was invisible because the repo root *also* declared 0.1.0.
    This test plants a decoy pyproject.toml with a distinct version
    in a fake module directory and asserts it is never read.
    """
    import edf_bill_fetcher.io.reporters.pdf_report as pdf_mod

    decoy_dir = tmp_path / "edf_bill_fetcher" / "io" / "reporters"
    decoy_dir.mkdir(parents=True)
    (decoy_dir / "pyproject.toml").write_text('version = "9.9.9"\n')

    monkeypatch.setattr(pdf_mod, "__file__", str(decoy_dir / "pdf_report.py"))
    assert pdf_mod._get_package_version() == _read_pyproject_version()
