# Run, develop and build from source

Use **64-bit Python 3.11** and Git. Python 3.10 and 3.12 are also tested; the current dependency set does not support Python 3.13 or newer. Ordinary CPython from python.org is recommended for release builds, matching CI. Conda is not required.

Clone the `how-its-going` branch and open a terminal in the checkout. On Windows, use `py -3.11` below; on macOS/Linux use `python3.11`.

```powershell
git clone --branch how-its-going https://github.com/volster/edf-bill-fetcher.git
cd edf-bill-fetcher
py -3.11 bootstrap.py run --install-system
```

This creates `.venv`, installs the Python dependencies including MSG support, checks native libraries and English OCR, then launches the app. No environment activation or PowerShell execution-policy change is needed. `--install-system` uses winget (or Chocolatey), Homebrew, or apt and may request administrator access. Package signatures and normal package-manager checks remain enabled. On other Linux distributions, install the equivalent packages yourself and omit this option.

For subsequent runs:

```powershell
.\.venv\Scripts\python.exe main.py
```

On macOS/Linux, the environment interpreter is `.venv/bin/python`.

## Development and checks

```powershell
py -3.11 bootstrap.py develop --install-system
.\.venv\Scripts\python.exe -m pytest
```

The `develop` mode installs the `[dev,build,msg]` extras, including pytest, lint/type-check tools and PyInstaller. The regular `run` mode does not install development tools. Both use `.github/ci-constraints.txt`.

To inspect an existing environment without installing or launching anything:

```powershell
py -3.11 bootstrap.py --check
.\.venv\Scripts\python.exe main.py --doctor --output diagnostics.json
```

Diagnostics return a nonzero exit code when a capability is missing. The JSON identifies the build and checks PDF, Excel, Qt, PST/OST, MSG and OCR support. It contains no selected evidence files or account details. Native failure messages may contain local installation paths; review them before sharing.

Tesseract is a native program, not a pip dependency. A full source installation needs its executable **and English `eng.traineddata`**. Windows' default `C:\Program Files\Tesseract-OCR` location is supported. On macOS use `brew install tesseract`; on Debian/Ubuntu use `sudo apt install tesseract-ocr tesseract-ocr-eng`. Merely installing `requirements.txt` does not install this engine.

The desktop shows extraction readiness before a run. A selected PST/OST source with no working reader blocks extraction. Missing OCR or MSG support requires explicit acknowledgement of the displayed limitations. Refresh checks after repairing an installation. Generating reports from saved analysis does not require OCR.

## Build for this computer

```powershell
py -3.11 bootstrap.py build --install-system
```

Or, after development setup:

```powershell
.\.venv\Scripts\python.exe ci_build.py
.\.venv\Scripts\python.exe .github/scripts/check_frozen_ocr.py
```

`ci_build.py` handles all three platforms. PyInstaller builds for the operating system and architecture of the Python interpreter running it:

| Build machine | Output |
| --- | --- |
| Windows | `dist/EDF_Evidence_Collector.exe` |
| macOS | `dist/EDF_Evidence_Collector.app` |
| Linux | `dist/EDF_Evidence_Collector` |

Each build embeds its source revision, local-change flag and build time. **About this build** displays that information without requiring Git on the user's machine. Tesseract, its native dependencies and English data are bundled. CI runs diagnostics and a real scanned-bill extraction through the executable before uploading it. The Mac workflow also checks the extracted distribution archive; see [MACOS.md](MACOS.md) for signing and first-open instructions.

## Build all three platforms

PyInstaller cannot produce Windows, macOS and Linux executables from one Windows installation. The GitHub workflow builds them on separate native runners.

Install GitHub CLI and authenticate with `gh auth login`, then run:

```powershell
py -3.11 build_all.py
```

This builds the **remote `how-its-going` branch**, waits for this exact request to finish, and downloads its three platform artifacts into `dist-all/<run-id>/`. Push your changes first; the script does not upload local or uncommitted edits. A failed platform build stops the command rather than downloading an older successful run. The workflow must include the `request_id` input used by the script.

Use `--ref <remote-branch-or-tag>`, `--repo <owner/repository>` or `--output <folder>` to change those defaults. You need permission to dispatch workflows and download their artifacts. These manually dispatched builds retain their `manual-unvalidated` label; automatic builds follow successful CI. Downloading all three does not mean all three can run on your computer.

Without GitHub CLI, open **Actions → Build Executables → Run workflow**, select the branch and download the three application artifacts when the run succeeds.

## Windows Tcl/Tk test failures

Use `python -m pytest` from the intended environment so an unrelated pytest installation is not selected. The Tcl/Tk startup-recovery tests exercise the retained Tk interface and CI recovery helper; the current desktop uses Qt. Synthetic library paths used by those tests must not replace the real interpreter's configuration, including Conda's `TCL_LIBRARY` and `TK_LIBRARY` values. Missing real libraries and failed second attempts remain test failures; tests are not skipped or retried to hide them.
