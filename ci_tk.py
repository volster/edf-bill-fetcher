"""Bound recovery to transient native Tcl/Tk file reads on Windows CI runners."""

from __future__ import annotations

import hashlib
import json
import os
import re
import sys
import time
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any


def _script_evidence(error: Exception) -> list[dict]:
    """Verify each script Tcl reports unreadable, without copying or modifying it."""
    paths = list(dict.fromkeys(re.findall(r'couldn\'t read file "([^"\r\n]+)"', str(error))))
    bootstrap = str(error).strip() == 'invalid command name "tcl_findLibrary"'
    if not paths and bootstrap:
        import _tkinter

        # setup-python's Windows distribution keeps these scripts beside the
        # runtime. Do not guess a different library when an override is set.
        if os.environ.get("TCL_LIBRARY") or os.environ.get("TK_LIBRARY"):
            return []
        directory = Path(sys.base_prefix) / "tcl"
        tcl = directory / f"tcl{_tkinter.TCL_VERSION}"
        paths = [str(tcl / name) for name in ("init.tcl", "auto.tcl", "tclIndex")]
        paths.append(str(directory / f"tk{_tkinter.TK_VERSION}" / "tk.tcl"))
    evidence = []
    for filename in paths:
        item: dict = {"path": filename}
        try:
            data = Path(filename).read_bytes()
            item.update(
                size=len(data), sha256=hashlib.sha256(data).hexdigest(), readable=bool(data)
            )
            if bootstrap:
                item["origin"] = "Default CPython Tcl/Tk installation"
                if Path(filename).name == "auto.tcl":
                    item["defines_bootstrap_command"] = bool(
                        re.search(rb"(?m)^proc\s+(?:::)?tcl_findLibrary\s", data)
                    )
                elif Path(filename).name == "tclIndex":
                    item["indexes_bootstrap_command"] = b"auto_index(tcl_findLibrary)" in data
        except OSError as problem:
            item.update(readable=False, error=str(problem))
        evidence.append(item)
    return evidence


def _create_with_retry(create: Callable[..., Any], args: tuple, kwargs: dict, report: Path) -> Any:
    """Retry only the native constructor when its allegedly missing scripts are readable."""
    import _tkinter

    try:
        return create(*args, **kwargs)
    except _tkinter.TclError as error:
        detail = str(error)
        direct_read = re.fullmatch(
            r'couldn\'t read file "([^"\r\n]+\.tcl)": [^\r\n]+', detail.strip()
        )
        installed_script = direct_read is not None and Path(
            direct_read[1]
        ).resolve().is_relative_to((Path(sys.base_prefix) / "tcl").resolve())
        if (
            not installed_script
            and detail.strip() != 'invalid command name "tcl_findLibrary"'
            and not any(
                marker in detail
                for marker in ("Can't find a usable tk.tcl", "Can't find a usable init.tcl")
            )
        ):
            raise
        # Let a transient read lock settle before verifying the files; a
        # permanently absent or unreadable script still prevents any retry.
        time.sleep(0.25)
        evidence = _script_evidence(error)
        retryable = bool(evidence) and all(
            item["readable"]
            and item.get("defines_bootstrap_command", True)
            and item.get("indexes_bootstrap_command", True)
            for item in evidence
        )
        record = {
            "executable": sys.executable,
            "error": detail,
            "scripts": evidence,
            "retryable": retryable,
            "recovered": False,
        }
        report.parent.mkdir(parents=True, exist_ok=True)

        def save() -> None:
            with report.open("a", encoding="utf-8") as handle:
                handle.write(json.dumps(record) + "\n")
            if not record["recovered"] and os.environ.get("GITHUB_ACTIONS") == "true":
                annotation = json.dumps(record)[-6000:]
                annotation = (
                    annotation.replace("%", "%25").replace("\r", "%0D").replace("\n", "%0A")
                )
                sys.stdout.write("::error title=Windows Tcl/Tk startup::" + annotation + "\n")

        if not retryable:
            save()
            raise
        try:
            # Catching the first error does not rerun a fixture, test or app
            # operation. Only this native initialization call is repeated.
            result = create(*args, **kwargs)
        except Exception as repeated:
            record["retry_error"] = str(repeated)
            save()
            raise
        record["recovered"] = True
        save()
        message = (
            "Native Tk startup recovered after one retry; verified Tcl/Tk script "
            "reads and hashes are retained in tk-startup.jsonl."
        )
        if os.environ.get("GITHUB_ACTIONS") == "true":
            sys.stdout.write("::warning title=Windows Tcl/Tk startup recovered::" + message + "\n")
        else:
            sys.stderr.write(message + "\n")
        return result


@contextmanager
def windows_tk_startup_retry() -> Iterator[None]:
    """Guard the native constructor only on Windows GitHub Actions, then restore it."""
    if sys.platform != "win32" or os.environ.get("GITHUB_ACTIONS") != "true":
        yield
        return
    import _tkinter

    original = _tkinter.create
    report = Path(".ci-diagnostics/tk-startup.jsonl").resolve()

    def create(*args: Any, **kwargs: Any) -> Any:
        return _create_with_retry(original, args, kwargs, report)

    _tkinter.create = create
    try:
        yield
    finally:
        _tkinter.create = original
