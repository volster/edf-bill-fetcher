"""Mac distribution must fail on bad signatures and failed extracted-app launches."""

import json
import shutil
import subprocess
import zipfile
from pathlib import Path

import pytest

import ci_build
import ci_macos


def make_app(root):
    app = root / ci_macos.APP_NAME
    binary = app / "Contents" / "MacOS" / ci_macos.BINARY_NAME
    binary.parent.mkdir(parents=True)
    binary.write_bytes(b"synthetic Mach-O placeholder")
    binary.chmod(0o755)
    return app


def test_missing_app_fails_before_native_tools(tmp_path, monkeypatch):
    monkeypatch.setattr(subprocess, "run", lambda *a, **k: pytest.fail("Missing app was accepted"))
    with pytest.raises(ValueError, match="Missing executable Mac app"):
        ci_macos.verify_app(tmp_path / ci_macos.APP_NAME)


def test_damaged_signature_is_fatal(tmp_path, monkeypatch):
    app = make_app(tmp_path)
    calls = []

    def damaged(command, **kwargs):
        calls.append(command)
        raise subprocess.CalledProcessError(1, command)

    monkeypatch.setattr(subprocess, "run", damaged)
    with pytest.raises(subprocess.CalledProcessError):
        ci_macos.verify_app(app)
    assert len(calls) == 1
    assert {"--verify", "--deep", "--strict", "--all-architectures"} <= set(calls[0])


@pytest.mark.parametrize("architecture", ["arm64", "x86_64 arm64", "x86_64", ""])
def test_apple_silicon_download_requires_arm64(tmp_path, monkeypatch, architecture):
    app = make_app(tmp_path)
    monkeypatch.setattr(
        subprocess, "run", lambda *a, **k: subprocess.CompletedProcess(a[0], 0, architecture)
    )
    if "arm64" in architecture.split():
        assert ci_macos.verify_app(app) == architecture
    else:
        with pytest.raises(ValueError, match="Apple silicon architecture missing"):
            ci_macos.verify_app(app)


@pytest.mark.parametrize(
    "report",
    [
        None,
        {"schema": 1, "visible": False, "pages": 4, "platform": "cocoa", "frozen": True},
        {"schema": 1, "visible": True, "pages": 4, "platform": "offscreen", "frozen": True},
        {"schema": 1, "visible": True, "pages": 4, "platform": "cocoa", "frozen": False},
        {"schema": 1, "visible": True, "pages": 4, "platform": "cocoa", "frozen": True},
    ],
)
def test_launchservices_success_requires_fresh_cocoa_window_evidence(tmp_path, monkeypatch, report):
    path = tmp_path / "gui.json"
    # A stale successful report must not let a crashed app pass.
    good = {"schema": 1, "visible": True, "pages": 4, "platform": "cocoa", "frozen": True}
    path.write_text(json.dumps(good))
    calls = []

    def opened(command, **kwargs):
        calls.append((command, kwargs))
        assert not path.exists()
        if report is not None:
            path.write_text(json.dumps(report))
        return subprocess.CompletedProcess(command, 0)

    monkeypatch.setattr(subprocess, "run", opened)
    if report == good:
        ci_macos.check_gui(tmp_path / ci_macos.APP_NAME, path)
    else:
        with pytest.raises((FileNotFoundError, ValueError)):
            ci_macos.check_gui(tmp_path / ci_macos.APP_NAME, path)
    assert calls[0][0][:3] == ["/usr/bin/open", "-W", "-n"]
    assert calls[0][1]["timeout"] == 90


@pytest.mark.parametrize("platform", ["darwin", "win32", "linux"])
def test_build_uses_standard_mac_bundle_and_propagates_pyinstaller_failure(monkeypatch, platform):
    calls = []
    monkeypatch.setattr(ci_build.sys, "platform", platform)
    provenance = {"version": "1.2.3", "commit": "a" * 40, "dirty": False}
    monkeypatch.setattr(ci_build, "build_metadata", lambda root: provenance)

    def failed(command, **kwargs):
        calls.append((command, kwargs))
        data_argument = next(value for value in command if value.startswith("--add-data="))
        metadata = data_argument.removeprefix("--add-data=").rsplit(ci_build.os.pathsep, 1)[0]
        assert json.loads(Path(metadata).read_text()) == provenance
        raise subprocess.CalledProcessError(1, command)

    monkeypatch.setattr(subprocess, "run", failed)
    with pytest.raises(subprocess.CalledProcessError):
        ci_build.main()
    command, options = calls[0]
    assert command[:3] == [ci_build.sys.executable, "-m", "PyInstaller"]
    assert options["check"] is True
    if platform == "darwin":
        assert "--onedir" in command
        assert "--onefile" not in command
        assert "--osx-bundle-identifier=io.github.volster.edf-bill-fetcher" in command
        assert options["env"]["PYINSTALLER_STRICT_BUNDLE_CODESIGN_ERROR"] == "1"
        assert options["env"]["PYINSTALLER_VERIFY_BUNDLE_SIGNATURE"] == "1"
    else:
        assert "--onefile" in command


def test_bad_source_signature_prevents_archive_creation(tmp_path, monkeypatch):
    make_app(tmp_path / "dist")

    def failed(*args, **kwargs):
        raise subprocess.CalledProcessError(1, ["codesign"])

    monkeypatch.setattr(subprocess, "run", failed)
    archive = tmp_path / "download.zip"
    with pytest.raises(subprocess.CalledProcessError):
        ci_macos.package_macos(
            archive, dist=tmp_path / "dist", diagnostics=tmp_path / "diagnostics"
        )
    assert not archive.exists()


@pytest.mark.parametrize("damaged_archive", [False, True])
def test_packaging_checks_the_extracted_app_and_excludes_raw_build_tree(
    tmp_path, monkeypatch, damaged_archive
):
    dist = tmp_path / "dist"
    original = make_app(dist)
    (dist / "build-dependencies.txt").write_text("example==1\n")
    (dist / "build-provenance.json").write_text('{"sha": "example"}')
    (dist / "raw-build-only.txt").write_text("Do not distribute the build directory")
    verified = []
    opened = []
    ocr = []

    def verify(app):
        assert (app / "Contents" / "MacOS" / ci_macos.BINARY_NAME).is_file()
        verified.append(app)
        if len(verified) == 2 and damaged_archive:
            raise ValueError("Damaged extracted signature")
        return "arm64"

    def run(command, **kwargs):
        # Stand in for native ditto only to exercise orchestration and layout.
        # Actual Apple archive/signature behavior is checked on the Mac runner.
        if command[0] == "/usr/bin/ditto":
            if "-c" in command:
                source, destination = map(Path, command[-2:])
                with zipfile.ZipFile(destination, "w") as archive:
                    for path in source.rglob("*"):
                        if path.is_file():
                            archive.write(path, path.relative_to(source.parent))
            elif "-x" in command:
                with zipfile.ZipFile(command[-2]) as archive:
                    archive.extractall(command[-1])
            else:
                shutil.copytree(command[-2], command[-1], symlinks=True)
        else:
            assert "check_frozen_ocr.py" in command[1]
            ocr.append(Path(command[-1]))
        return subprocess.CompletedProcess(command, 0)

    monkeypatch.setattr(ci_macos, "verify_app", verify)
    monkeypatch.setattr(ci_macos, "check_gui", lambda app, report: opened.append(app))
    monkeypatch.setattr(subprocess, "run", run)
    archive = tmp_path / "download.zip"
    diagnostics = tmp_path / "diagnostics"
    if damaged_archive:
        with pytest.raises(ValueError, match="Damaged extracted signature"):
            ci_macos.package_macos(archive, dist=dist, diagnostics=diagnostics)
        assert not opened
        assert not ocr
    else:
        ci_macos.package_macos(archive, dist=dist, diagnostics=diagnostics)
        assert opened == [verified[1]]
        assert ocr == [verified[1] / "Contents" / "MacOS" / ci_macos.BINARY_NAME]
        assert json.loads((diagnostics / "macos-trust.json").read_text())["notarized"] is False
    assert len(verified) == 2
    assert verified[0] == original
    assert verified[1] != original
    with zipfile.ZipFile(archive) as bundle:
        names = bundle.namelist()
        assert not any("raw-build-only" in name for name in names)
        assert "EDF Evidence Collector/START_HERE.txt" in names
        assert "EDF Evidence Collector/build-provenance.json" in names
        trust = json.loads(bundle.read("EDF Evidence Collector/macos-trust.json"))
        assert trust["signature"] == "ad-hoc"
