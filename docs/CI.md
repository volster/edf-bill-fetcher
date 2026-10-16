# CI and executable build troubleshooting

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

## What the Windows investigation established

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
- A historical run reached the six-hour platform limit. Test/build jobs now have
  a 30-minute limit; superseded branch CI and executable builds are cancelled.

## Diagnosing a new failure

Start with the failing job and step, rather than rerunning the whole workflow.
Every test job uploads `ci-<os>-py<version>-<attempt>` for 14 days, including:

- `python.json`: the selected interpreter, Python version and runner image.
- `runtime.json`: a real GUI preflight, Tcl/Tk locations, SHA-256 hashes of
  readable `.tcl` scripts, Qt version/platform, or a startup traceback.
- `dependencies.txt`: exact installed package versions.
- `tests/attempt-*/test-results.xml` and `coverage.xml`, plus `attempts.json`:
  the original evidence from each completed test attempt.
- Top-level JUnit and coverage reports for the final complete suite.

Artifacts include only allowlisted environment fields, not a dump of tokens or
other runner environment variables. Setup failures before Python is installed
may have no diagnostic files; the setup action's log is then the primary evidence.

The GUI preflight creates/destroys 20 Tk roots, loads all available ttk themes,
and opens the actual Qt workspace with temporary preferences. Linux runs it
under Xvfb. Do not add a global `TCL_LIBRARY`, `TK_LIBRARY`, copied library tree
or security-service change without evidence identifying a necessary fix.

`ci_pytest.py` retains one narrowly bounded Windows GitHub Actions retry for
JUnit **setup errors solely identifying an unusable `tk.tcl` or `init.tcl`**.
It never retries assertions, mixed failures, collection errors or local runs.
A retry runs the full suite again, records a visible job-summary entry, and
keeps both attempts' evidence. Its second exit code is authoritative; coverage
and JUnit from a failed first attempt are never presented as the final result.
This fallback is not evidence that a runtime failure has been resolved.

Build jobs upload separate `build-diagnostics-*` artifacts even on failure.
These retain runtime/dependency reports and PyInstaller warning/import reports.
They are excluded from release asset downloads. The delivered binary must also
pass the image-only PDF OCR and amended-Excel round-trip smoke check. The source
GUI preflight does not by itself certify a frozen executable's interactive GUI.

## Reproduce and maintain

```bash
python -m pip install -c .github/ci-constraints.txt -e ".[dev,msg]"
python ci_runtime.py --inventory
python ci_runtime.py       # Linux: xvfb-run -a python ci_runtime.py
python ci_pytest.py        # Linux: xvfb-run -a python ci_pytest.py
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
