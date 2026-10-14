"""Resolve the EDF package version without borrowing another project's metadata."""

from __future__ import annotations

import sys
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

if sys.version_info >= (3, 11):
    import tomllib
else:
    import tomli as tomllib

_DEFAULT_VERSION = "0.5.0"
_DISTRIBUTION = "edf-bill-fetcher"
_MODULE_DIR = Path(__file__).resolve().parent


def get_package_version() -> str:
    """Return the checkout version or the installed EDF distribution version.

    The source manifest belongs exactly two directories above this helper.
    Installed and frozen packages use their distribution metadata instead;
    searching further ancestors could attribute another project's version
    to the evidence reports. An unavailable manifest or distribution falls
    back to the last declared package version.
    """
    candidate = _MODULE_DIR.parent.parent / "pyproject.toml"
    try:
        with candidate.open("rb") as manifest:
            project = tomllib.load(manifest).get("project", {})
        if isinstance(project, dict) and project.get("name") == _DISTRIBUTION:
            declared = project.get("version")
            if isinstance(declared, str) and declared.strip():
                return declared
    except (OSError, ValueError):
        pass
    try:
        installed = version(_DISTRIBUTION)
        if isinstance(installed, str) and installed.strip():
            return installed
    except (PackageNotFoundError, OSError):
        pass
    return _DEFAULT_VERSION
