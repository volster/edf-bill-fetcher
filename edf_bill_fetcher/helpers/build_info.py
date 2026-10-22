"""Identify source checkouts and frozen builds without requiring Git at runtime."""

from __future__ import annotations

import json
import re
import subprocess
import sys
from datetime import datetime, timezone
from importlib.resources import files
from pathlib import Path

from edf_bill_fetcher.helpers.version import get_package_version


def source_build_info(root: Path) -> dict:
    """Capture the exact source revision, including whether tracked files changed."""
    info = {"version": get_package_version(), "commit": "unknown", "dirty": False}
    try:
        revision = subprocess.run(
            ["git", "-C", str(root), "rev-parse", "--verify", "HEAD"],
            capture_output=True,
            text=True,
            check=True,
            timeout=5,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        ).stdout.strip()
        if re.fullmatch(r"[0-9a-f]{40,64}", revision):
            info["commit"] = revision
            status = subprocess.run(
                ["git", "-C", str(root), "status", "--porcelain", "--untracked-files=no"],
                capture_output=True,
                text=True,
                check=True,
                timeout=5,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            info["dirty"] = bool(status.stdout.strip())
    except (OSError, subprocess.SubprocessError):
        pass
    return info


def build_metadata(root: Path) -> dict:
    """Stamp a build once so distributing it does not change its identity."""
    return {**source_build_info(root), "built_at": datetime.now(timezone.utc).isoformat()}


def get_build_info() -> dict:
    """Read embedded provenance in packaged apps and Git only in source checkouts."""
    if getattr(sys, "frozen", False):
        try:
            data = json.loads(
                files("edf_bill_fetcher.data").joinpath("build-info.json").read_text()
            )
            if isinstance(data, dict) and isinstance(data.get("commit"), str):
                return {**data, "mode": "packaged"}
        except (OSError, ValueError):
            pass
        return {"version": get_package_version(), "commit": "unknown", "mode": "packaged"}
    return {**source_build_info(Path(__file__).resolve().parents[2]), "mode": "source"}
