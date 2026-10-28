# Run, develop and build from source

[Documentation index](README.md) · [Download a build](../README.md#download-a-build)

Install Git and either Conda or **64-bit Python 3.11**. Application environments
using Python 3.10 and 3.12 are also tested; the current dependencies do not support
Python 3.13 or newer. The bootstrap launcher can use a newer Python when Conda
will provide the supported project interpreter.

Clone the working branch and open a terminal in the checkout:

```console
git clone --branch how-its-going https://github.com/volster/edf-bill-fetcher.git
cd edf-bill-fetcher
```

## Automatic environment setup

In a Conda-enabled terminal, such as an Anaconda/Miniforge prompt:

```console
python bootstrap.py run --install-system
```

Without Conda, use `py -3.11 bootstrap.py run --install-system` on Windows or
`python3.11 bootstrap.py run --install-system` on macOS/Debian/Ubuntu.

The bootstrapper:

1. Looks for Conda through `CONDA_EXE`, `PATH` and common installation locations.
2. If found, looks for a registered environment named **edf-bill-fetcher**. It
   reuses that environment, or creates it with Python 3.11 and pip if absent.
3. If Conda is absent, reuses or creates **.venv** in the checkout using the
   launcher's Python. Creating this venv requires Python 3.10–3.12.
4. Installs the requested Python extras using `.github/ci-constraints.txt`, checks
   native capabilities including English OCR, then runs the selected action.

Conda commands run through `conda run --prefix …`, so environment activation and
Windows DLL search paths apply to pip, diagnostics, the app and builds. The app
is never installed into Conda's `base` environment. Existing project environments
must contain a supported 64-bit Python; incompatible environments produce an
error rather than having their Python silently replaced.

### Choose an environment explicitly

```console
python bootstrap.py develop --conda-env my-edf-project
python bootstrap.py run --venv .venv
```

`--conda-env NAME` reuses or creates that named environment and requires Conda.
`--venv PATH` bypasses Conda detection and retains the existing venv workflow.
The options are mutually exclusive. Neither option selects an unrelated active
environment automatically. If you have an existing project environment with a
custom name, supply it explicitly.

If Conda is found but its inventory, solver, downloads or commands fail, setup
stops with the error. It does not silently switch to venv. Repair Conda or
explicitly choose `--venv .venv`. Conda uses your configured channels and normal
package policies; the bootstrapper does not accept channel terms on your behalf.
Duplicate registered environments with the requested name are rejected; use a
unique name. For multiple checkouts, give each its own environment name because
an editable pip installation points to the checkout most recently installed.

### Subsequent launches

With the default Conda environment:

```console
conda activate edf-bill-fetcher
python main.py
```

Without activation:

```console
conda run --no-capture-output -n edf-bill-fetcher python main.py
```

For venv on Windows:

```powershell
.\.venv\Scripts\python.exe main.py
```

For venv on macOS/Linux:

```console
.venv/bin/python main.py
```

The remaining examples use `python` to mean the selected environment's Python.
Activate it first, use the full venv interpreter path, or prefix the command with
`conda run --no-capture-output -n edf-bill-fetcher` as appropriate. Running the
bootstrapper again is also supported, but it repeats dependency installation and
readiness checks before launching.

## Native prerequisites

`--install-system` uses winget (or Chocolatey), Homebrew, or apt and may request
administrator access. Package signatures and normal package-manager checks stay
enabled. On other Linux distributions, install the equivalent packages yourself
and omit that flag. Conda environment selection does not remove native requirements.

Tesseract is a native program, not a pip dependency. A complete source installation
needs its executable **and English `eng.traineddata`**:

- Windows: install Tesseract with English data; the default
  `C:\Program Files\Tesseract-OCR` location is supported.
- macOS: `brew install tesseract`.
- Debian/Ubuntu: `sudo apt install tesseract-ocr tesseract-ocr-eng`.

Linux desktop use also needs Qt libraries (`libxcb-cursor0`,
`libxkbcommon-x11-0`, `libegl1` on Debian/Ubuntu). A libpff source build can require
compiler/autotools development packages when a compatible wheel is unavailable;
`--install-system` includes the relevant Linux/macOS build prerequisites.

The desktop checks readiness before extraction. A selected PST/OST source without
a working reader blocks extraction. Missing OCR or MSG support requires explicit
acknowledgement of the limitations. Refresh checks after repairing an installation.
Generating reports from saved analysis does not require OCR.

## Manual pip setup

The bootstrapper is optional. After installing native prerequisites and creating
or activating your own supported environment, run:

```console
python -m pip install -c .github/ci-constraints.txt -e ".[msg]"
python main.py --doctor
python main.py
```

For a manual venv, first run `py -3.11 -m venv .venv` on Windows or
`python3.11 -m venv .venv` on macOS/Linux, then use its Python for the commands
above. Use `.[dev,build,msg]` in place of `.[msg]` for the development toolchain.
Installing `requirements.txt` alone does not install the native OCR engine.

## Development and checks

```console
python bootstrap.py develop --install-system
```

`develop` installs `[dev,build,msg]`, including pytest, lint/type checks and
PyInstaller. `run` installs runtime and MSG dependencies; `build` additionally
installs PyInstaller. All use the constrained pip requirements.

Inspect the automatically selected environment without installing or launching:

```console
python bootstrap.py --check
```

Add `--conda-env NAME` or `--venv PATH` to check an explicitly chosen environment.
`--check` fails if that environment does not exist; it does not create one or
install system packages, even if `--install-system` is also present.

From the selected environment, run:

```console
python -m pytest
python main.py --doctor --output diagnostics.json
```

Diagnostics return nonzero when a capability is missing. They identify the build
and check PDF, Excel, Qt, PST/OST, MSG and OCR support. No selected evidence files
or account details are included. Native failure messages may contain installation
paths; review them before sharing. See [Development](DEVELOPMENT.md) for daily checks.

## Build for this computer

```console
python bootstrap.py build --install-system
```

Or, after development setup:

```console
python ci_build.py
python .github/scripts/check_frozen_ocr.py
```

Official CI builds use standard CPython from `actions/setup-python`. A local Conda
build reflects its own resolved runtime; run the frozen checks before distributing it.

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

```console
python build_all.py
```

This builds the **remote `how-its-going` branch**, waits for this exact request to finish, and downloads its three platform artifacts into `dist-all/<run-id>/`. Push your changes first; the script does not upload local or uncommitted edits. A failed platform build stops the command rather than downloading an older successful run. The workflow must include the `request_id` input used by the script.

Use `--ref <remote-branch-or-tag>`, `--repo <owner/repository>` or `--output <folder>` to change those defaults. You need permission to dispatch workflows and download their artifacts. These manually dispatched builds retain their `manual-unvalidated` label; automatic builds follow successful CI. Downloading all three does not mean all three can run on your computer.

Without GitHub CLI, open **Actions → Build Executables → Run workflow**, select the branch and download the three application artifacts when the run succeeds.

For day-to-day checks and contribution guidance, continue to [Development](DEVELOPMENT.md).

## GUI test runtime

Use `python -m pytest` from the intended environment so an unrelated pytest installation is not selected. The supported desktop and runtime preflight use Qt. Headless Linux tests select Qt's offscreen platform automatically; CI sets it explicitly on every platform. The retired Tk desktop, Tcl/Tk startup probes and library recovery wrapper have been removed, and frozen builds explicitly exclude Tk.
