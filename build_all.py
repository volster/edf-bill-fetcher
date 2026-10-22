"""Build Windows, macOS and Linux on GitHub, then download this request's artifacts."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path
from uuid import uuid4

WORKFLOW = "matrix_build.yml"


def gh(arguments: list[str], *, capture: bool = False) -> str:
    """Use the user's existing GitHub CLI authentication, never a shell or token file."""
    result = subprocess.run(["gh", *arguments], text=True, check=True, capture_output=capture)
    return result.stdout or ""


def build_all(repo: str, ref: str, output: Path) -> Path:
    """Correlate a dispatch with a unique request ID rather than choosing the latest run."""
    request = uuid4().hex
    gh(["workflow", "run", WORKFLOW, "--repo", repo, "--ref", ref, "-f", f"request_id={request}"])
    deadline = time.monotonic() + 120
    run = None
    while time.monotonic() < deadline:
        runs = json.loads(
            gh(
                [
                    "run",
                    "list",
                    "--repo",
                    repo,
                    "--workflow",
                    WORKFLOW,
                    "--event",
                    "workflow_dispatch",
                    "--limit",
                    "50",
                    "--json",
                    "databaseId,displayTitle,url,headSha",
                ],
                capture=True,
            )
        )
        run = next((item for item in runs if item["displayTitle"] == f"EDF build {request}"), None)
        if run:
            break
        time.sleep(3)
    if run is None:
        raise RuntimeError(
            "GitHub did not expose this build request within two minutes. Check Actions; no older artifacts were downloaded."
        )
    sys.stdout.write(f"Building revision {run['headSha']}: {run['url']}\n")
    run_id = str(run["databaseId"])
    gh(["run", "watch", run_id, "--repo", repo, "--exit-status"])
    destination = output.resolve() / run_id
    if destination.exists():
        raise RuntimeError(f"Refusing to mix artifacts into an existing directory: {destination}")
    destination.mkdir(parents=True)
    gh(
        [
            "run",
            "download",
            run_id,
            "--repo",
            repo,
            "--pattern",
            "EDF_Evidence_Collector-*",
            "--dir",
            str(destination),
        ]
    )
    expected = {
        "windows-2025": "EDF_Evidence_Collector.exe",
        "macos-15": "EDF_Evidence_Collector-macos-arm64-ad-hoc.zip",
        "ubuntu-24.04": "EDF_Evidence_Collector",
    }
    for platform, filename in expected.items():
        matches = list(destination.glob(f"EDF_Evidence_Collector-{platform}*"))
        if len(matches) != 1 or not any(path.is_file() for path in matches[0].rglob(filename)):
            raise RuntimeError(f"Missing or ambiguous {platform} artifact in {destination}")
    sys.stdout.write(f"All three platform artifacts downloaded to {destination}\n")
    return destination


def main(argv: list[str] | None = None) -> int:
    """Dispatch a native build matrix; this does not cross-compile on the local machine."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", default="volster/edf-bill-fetcher")
    parser.add_argument(
        "--ref",
        default="how-its-going",
        help="Remote branch or tag to build; local uncommitted edits are not uploaded",
    )
    parser.add_argument("--output", type=Path, default=Path("dist-all"))
    options = parser.parse_args(argv)
    if shutil.which("gh") is None:
        sys.stderr.write("Install GitHub CLI and run gh auth login, then rerun this command.\n")
        return 1
    try:
        build_all(options.repo, options.ref, options.output)
        return 0
    except (OSError, RuntimeError, ValueError, subprocess.SubprocessError) as error:
        sys.stderr.write(f"Build or artifact download failed: {error}\n")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
