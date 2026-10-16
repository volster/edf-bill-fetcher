"""Package and validate the Mac download, preserving the signed app bundle."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

APP_NAME = "EDF_Evidence_Collector.app"
BINARY_NAME = "EDF_Evidence_Collector"


def verify_app(app: Path) -> str:
    """Reject missing, damaged or incompatible Apple silicon bundles."""
    binary = app / "Contents" / "MacOS" / BINARY_NAME
    if not binary.is_file() or not os.access(binary, os.X_OK):
        raise ValueError(f"Missing executable Mac app: {binary}")
    subprocess.run(
        [
            "/usr/bin/codesign",
            "--verify",
            "--deep",
            "--strict",
            "--all-architectures",
            "--verbose=4",
            str(app),
        ],
        check=True,
    )
    architectures = subprocess.run(
        ["/usr/bin/lipo", "-archs", str(binary)],
        check=True,
        capture_output=True,
        text=True,
    ).stdout.strip()
    if "arm64" not in architectures.split():
        raise ValueError(f"Apple silicon architecture missing: {architectures}")
    return architectures


def check_gui(app: Path, report_path: Path) -> None:
    """Open the extracted app through LaunchServices and require a real Cocoa window."""
    report_path.unlink(missing_ok=True)
    subprocess.run(
        [
            "/usr/bin/open",
            "-W",
            "-n",
            str(app),
            "--args",
            "--check-gui",
            "--report",
            str(report_path),
        ],
        check=True,
        timeout=90,
    )
    # `open` can succeed even when its child crashes: require evidence from Qt.
    report = json.loads(report_path.read_text(encoding="utf-8"))
    expected = {"schema": 1, "visible": True, "pages": 4, "platform": "cocoa", "frozen": True}
    if report != expected:
        raise ValueError(f"Extracted Mac app GUI check failed: {report}")


def package_macos(archive: Path, *, dist: Path, diagnostics: Path) -> None:
    """Verify, archive, extract and exercise exactly the app being distributed."""
    archive = archive.resolve()
    diagnostics.mkdir(parents=True, exist_ok=True)
    architectures = verify_app(dist / APP_NAME)
    trust = {
        "architecture": architectures,
        "signature": "ad-hoc",
        "notarized": False,
        "gatekeeper": "Requires one-time approval in Privacy & Security; see START_HERE.txt",
    }
    with tempfile.TemporaryDirectory(prefix="edf-macos-package-") as directory:
        root = Path(directory)
        staging = root / "EDF Evidence Collector"
        staging.mkdir()
        subprocess.run(
            ["/usr/bin/ditto", str(dist / APP_NAME), str(staging / APP_NAME)], check=True
        )
        for name in ("build-dependencies.txt", "build-provenance.json"):
            shutil.copyfile(dist / name, staging / name)
        shutil.copyfile(Path(__file__).parent / "docs" / "MACOS.md", staging / "START_HERE.txt")
        (staging / "macos-trust.json").write_text(
            json.dumps(trust, indent=2) + "\n", encoding="utf-8"
        )
        subprocess.run(
            [
                "/usr/bin/ditto",
                "-c",
                "-k",
                "--sequesterRsrc",
                "--keepParent",
                str(staging),
                str(archive),
            ],
            check=True,
        )
        extracted = root / "extracted"
        subprocess.run(["/usr/bin/ditto", "-x", "-k", str(archive), str(extracted)], check=True)
        app = extracted / staging.name / APP_NAME
        verify_app(app)
        check_gui(app, (diagnostics / "macos-extracted-gui.json").resolve())
        subprocess.run(
            [
                sys.executable,
                str(Path(__file__).parent / ".github" / "scripts" / "check_frozen_ocr.py"),
                "--executable",
                str(app / "Contents" / "MacOS" / BINARY_NAME),
            ],
            check=True,
            timeout=420,
        )
    (diagnostics / "macos-trust.json").write_text(
        json.dumps(trust, indent=2) + "\n", encoding="utf-8"
    )


def main() -> None:
    """Create an Apple silicon archive on a native Mac runner."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--archive", type=Path, required=True)
    parser.add_argument("--dist", type=Path, default=Path("dist"))
    parser.add_argument("--diagnostics", type=Path, default=Path(".ci-diagnostics"))
    options = parser.parse_args()
    if sys.platform != "darwin":
        parser.error("Mac packaging requires macOS (codesign, ditto and LaunchServices)")
    package_macos(options.archive, dist=options.dist, diagnostics=options.diagnostics)


if __name__ == "__main__":
    main()
