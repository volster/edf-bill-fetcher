# CI and executable build troubleshooting

[Documentation index](README.md) · [Development](DEVELOPMENT.md) · [Builds](BUILDING.md)

The required CI matrix runs Python 3.10, 3.11 and 3.12 on Ubuntu 24.04,
Windows Server 2025 and macOS 15. Executable builds use Python 3.11 from
the same `actions/setup-python` distribution. Runner major versions are
explicit so a `*-latest` image migration does not change the operating
system underneath a source change.

All nine test jobs retain lint, format, type and full-suite coverage checks.
The coverage requirement remains 90%. Automatic executable builds require
successful CI for the source revision recorded in `build-provenance.json`.
Manual builds remain labelled `manual-unvalidated`. Tagged releases validate
and build the tagged revision before publication.

## Workflow and artifact map

| Workflow | Trigger and revision | Result |
| --- | --- | --- |
| [CI](../.github/workflows/ci.yml) | Push/PR on configured branches, or reusable invocation for a tagged revision. | Nine OS/Python legs and the aggregate Required validation job. |
| [Build Executables](../.github/workflows/matrix_build.yml) | Successful CI completion on `how-its-going`, or manual dispatch. | Three native application artifacts and separate diagnostics. Automatic builds check out the completed CI SHA. |
| [Release](../.github/workflows/release.yml) | A `v*` tag. | Reuses CI for that tag, builds its validated SHA and publishes only after validation and all native builds succeed. |

An automatic artifact's name ends in `ci-validated`; manual dispatch uses
`manual-unvalidated`, even if the selected revision previously passed CI. The Mac
artifact also contains `arm64-ad-hoc`. Provenance inside the download records the
SHA and validation mode. The artifact belongs to that revision, not necessarily
the latest branch tip by the time it is downloaded.

Branch application artifacts expire after 7 days; diagnostics after 14 days.
Release workflow artifacts are retained for 90 days. The published release assets
are distinct from temporary Actions artifacts. GitHub sign-in can be required to
download an Actions artifact.

## Diagnosing a new failure

Start with the failing job and step, rather than rerunning the whole workflow.
Every test job uploads `ci-<os>-py<version>-<attempt>` for 14 days, including:

- `python.json`: the selected interpreter, Python version and runner image.
- `runtime.json`: the real Qt workspace startup result, Qt version/platform, or a startup traceback.
- `dependencies.txt`: exact installed package versions.
- `tests/attempt-*/test-results.xml` and `coverage.xml`, plus `attempts.json`:
  the full suite's original reports and exit code.
- Top-level JUnit and coverage reports for the final complete suite.

Artifacts include only allowlisted environment fields, not a dump of tokens or
other runner environment variables. Setup failures before Python is installed
may have no diagnostic files; the setup action's log is then the primary evidence.

The GUI preflight opens the actual Qt workspace with temporary preferences.
Headless checks use Qt’s offscreen platform. The retired Tk interface, Tcl/Tk
startup probes and native-constructor retry wrapper have been removed.

`ci_pytest.py` runs the full supported suite **once** and preserves its original
exit code and reports. There is no whole-suite retry or skipped GUI check to
hide a failure. Source-control and PDF-preview actions must remain reachable
at normal and narrow window sizes, including larger native font metrics.

Build jobs upload separate `build-diagnostics-*` artifacts even on failure.
These retain runtime/dependency reports and PyInstaller warning/import reports.
They are excluded from release asset downloads. The delivered binary must also
pass the image-only PDF OCR and amended-Excel round-trip smoke check. The source
GUI preflight does not by itself certify a frozen executable's interactive GUI.

Mac builds additionally verify deep, strict code-signature integrity, archive
with `ditto`, and verify the extracted `.app`. The extracted app must open a
real Cocoa window through LaunchServices, then pass OCR and Excel correction
checks. Signing failures are fatal during the PyInstaller build. Both build
workflows use `ci_build.py` and `ci_macos.py` so release packaging cannot drift
from the ordinary build. See [Mac opening and trust status](MACOS.md): ad-hoc
signing does not provide Gatekeeper trust or Apple notarization.

## Triage by failing step

| Failure location | First evidence and next check |
| --- | --- |
| Setup Python / native packages | The setup action/package-manager log. Python diagnostics may not exist yet. |
| Pip installation | Dependency resolver/build error and `dependencies.txt` if written; compare supported Python/architecture and constrained tool versions. |
| Runtime preflight | `runtime.json` and `python.json`; distinguish an importable package from a working native library or real workspace. |
| Pytest | First failing test, JUnit traceback and original attempt exit code. Reproduce that case before rerunning the full suite. |
| Coverage gate | The failing leg's full-suite report and command, not a different local subset. See [coverage scope](COVERAGE.md#measurement-scope). |
| PyInstaller | Warning/import reports, exact dependency inventory and build log. A successful source preflight does not establish the frozen imports. |
| Frozen OCR/review smoke | The executable's diagnostics and smoke failure. Check native OCR/data inclusion and generated artifacts rather than bypassing the check. |
| Mac extracted app | Signature/archive verification and Cocoa-window diagnostics. See [Mac integrity checks](MACOS.md#inspect-an-opening-problem). |
| Release publication | Validation/build dependency results and release asset selection. Do not publish an older or manually built binary as if it passed for this tag. |

On Windows, use explicit UTF-8 when reading the app's JSON/HTML/text artifacts.
Keep JSON decoding failures, layout failures and native library startup failures
separate: they have different causes and fixes. The diagnostics are evidence of
what happened, not a reason to add retries or disable assertions.

## Reproduce and maintain

Use the project environment after [native prerequisite setup](BUILDING.md).
The following uses Bash (including Git Bash on Windows); in PowerShell, set
`$env:QT_QPA_PLATFORM = "offscreen"` before `python ci_runtime.py`.

```bash
python -m pip install -c .github/ci-constraints.txt -e ".[dev,msg]"
python ci_runtime.py --inventory
QT_QPA_PLATFORM=offscreen python ci_runtime.py
python ci_pytest.py
```

The preflight runs setup-python’s explicit `python-path` output directly.
macOS framework Python can have a different `sys.prefix` from `pythonLocation`;
those directory strings must not be compared as an identity check. Invoke pip,
pytest, Ruff, mypy and PyInstaller through
`python -m` to keep tools attached to that interpreter.

`.github/ci-constraints.txt` pins CI/build tools to the reviewed versions;
application dependency ranges remain in `pyproject.toml`. Update constraints
deliberately and validate all nine jobs and all three executable builds. Pip
caches are keyed by the project and constraints files. Installed dependency
reports help identify changes in runtime packages, which are not fully locked.
Review runner-image updates separately from application changes where practical.

## Maintainer verification after a workflow change

Review the source ref at checkout, the aggregate dependency graph, tool/interpreter
selection, diagnostic retention and asset selection together. A successful sibling
matrix job cannot stand in for a failed one. Confirm that cancellation of an older
branch run does not relabel a newer or unrelated build as validated.

After the relevant remote checks finish, record the exact run URL and SHA. State
which jobs succeeded, failed or were not run. Do not describe a parsed YAML file,
a queued job or a local Linux check as a completed three-platform build.

## Historical Windows investigations

- [CI #201](https://github.com/volster/edf-bill-fetcher/actions/runs/38071063546)
  and #202 caught actual Qt layout defects under Windows font metrics. These
  were addressed by the later layout changes; such assertions must keep failing
  when the interface regresses.
- [CI #205](https://github.com/volster/edf-bill-fetcher/actions/runs/38801700446)
  failed during one Tk fixture's setup because Tcl could not read
  `D:/a/_temp/tklibs/tk8.6/tk.tcl`. That was the workflow's copied runtime,
  rather than the Python installation's own library tree.
- The previous workflow mixed an incompletely activated Conda environment,
  inherited runner DLL paths and a manually copied Tcl/Tk tree. Several native
  command failures could be overwritten by subsequent commands' exit codes.
  Its supposed file warmup enumerated paths without reading file contents.
  Defender changes and repeated mirror checks did not eliminate the failures.
- Available failure annotations establish the missing/unreadable runtime path.
  They do **not** prove that Defender deleted files or that the official Python
  installation was corrupt. The old speculative scanner workarounds were removed.
- [CI #214](https://github.com/volster/edf-bill-fetcher/actions/runs/38908627975)
  reproduced the failure under the installed CPython 3.10 tree after a successful
  preflight: a Tk root created **inside a test body** could not read `button.tcl`.
  The other eight jobs passed. Thus neither a temp mirror nor fixture setup is
  essential to the failure, and the old setup-error-only suite retry missed it.
- A historical run reached the six-hour platform limit. Test/build jobs now have
  a 30-minute limit; superseded branch CI and executable builds are cancelled.

A later Windows regression involved reading UTF-8 saved results with the default
Windows encoding in an integration test. Production loading already used explicit
UTF-8. The test now uses the production loader and exercises both UTF-8 and cp1252
defaults. When JSON identities or pound-sign field names differ across platforms,
check file decoding before weakening the identity validation.
