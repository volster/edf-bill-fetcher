"""Report provenance uses the EDF checkout or installed distribution version."""

from __future__ import annotations

import re
from pathlib import Path

import pytest
from reportlab.platypus import Paragraph

from edf_bill_fetcher.helpers import version as version_helper
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


def test_checkout_uses_own_project_section_and_ignores_nearby_decoys(monkeypatch, tmp_path):
    root = tmp_path / "checkout"
    helper = root / "edf_bill_fetcher" / "helpers"
    helper.mkdir(parents=True)
    (root / "pyproject.toml").write_text(
        '[tool.other]\nversion = "8.8.8"\n'
        '[project]\nname = "edf-bill-fetcher"\nversion = "1.2.3"\n',
        encoding="utf-8",
    )
    (helper / "pyproject.toml").write_text('version = "9.9.9"\n', encoding="utf-8")
    (tmp_path / "pyproject.toml").write_text('version = "7.7.7"\n', encoding="utf-8")
    monkeypatch.setattr(version_helper, "_MODULE_DIR", helper)
    monkeypatch.setattr(version_helper, "version", lambda _name: "6.6.6")
    assert _get_package_version() == "1.2.3"


@pytest.mark.parametrize(
    "manifest",
    [
        None,
        b'[project]\nname = "foreign-project"\nversion = "9.9.9"\n',
        b'[tool.other]\nname = "edf-bill-fetcher"\nversion = "9.9.9"\n',
        b'[project]\nname = "edf-bill-fetcher"\nversion = 999\n',
        b'[project]\nname = "edf-bill-fetcher"\nversion = ""\n',
        b'[project]\nname = "edf-bill-fetcher"\nversion = "unterminated\n',
        b"\xff",
    ],
)
def test_installed_version_does_not_borrow_foreign_or_invalid_manifests(
    monkeypatch, tmp_path, manifest
):
    helper = tmp_path / "site-packages" / "edf_bill_fetcher" / "helpers"
    helper.mkdir(parents=True)
    # This farther ancestor used to be incorrectly attributed to EDF.
    (tmp_path / "pyproject.toml").write_text(
        '[project]\nname = "foreign-project"\nversion = "9.9.9"\n', encoding="utf-8"
    )
    if manifest is not None:
        (helper.parent.parent / "pyproject.toml").write_bytes(manifest)
    monkeypatch.setattr(version_helper, "_MODULE_DIR", helper)
    seen = []

    def installed_version(name):
        seen.append(name)
        return "2.3.4"

    monkeypatch.setattr(version_helper, "version", installed_version)
    assert _get_package_version() == "2.3.4"
    assert seen == ["edf-bill-fetcher"]


def test_unreadable_manifest_uses_installed_version(monkeypatch, tmp_path):
    helper = tmp_path / "edf_bill_fetcher" / "helpers"
    helper.mkdir(parents=True)
    candidate = tmp_path / "pyproject.toml"
    candidate.mkdir()  # Opening a directory as a TOML manifest raises OSError on every platform.
    monkeypatch.setattr(version_helper, "_MODULE_DIR", helper)
    monkeypatch.setattr(version_helper, "version", lambda _name: "2.3.4")
    assert _get_package_version() == "2.3.4"


def test_distribution_metadata_is_used_without_a_source_manifest(monkeypatch, tmp_path):
    metadata = tmp_path / "edf_bill_fetcher-4.5.6.dist-info"
    metadata.mkdir()
    (metadata / "METADATA").write_text(
        "Metadata-Version: 2.1\nName: edf-bill-fetcher\nVersion: 4.5.6\n", encoding="utf-8"
    )
    monkeypatch.syspath_prepend(str(tmp_path))
    monkeypatch.setattr(version_helper, "_MODULE_DIR", tmp_path / "edf_bill_fetcher" / "helpers")
    assert _get_package_version() == "4.5.6"


def test_last_declared_default_is_only_used_when_metadata_is_unavailable(monkeypatch, tmp_path):
    monkeypatch.setattr(version_helper, "_MODULE_DIR", tmp_path / "edf_bill_fetcher" / "helpers")

    def missing_distribution(_name):
        raise version_helper.PackageNotFoundError("edf-bill-fetcher")

    monkeypatch.setattr(version_helper, "version", missing_distribution)
    assert _get_package_version() == version_helper._DEFAULT_VERSION
