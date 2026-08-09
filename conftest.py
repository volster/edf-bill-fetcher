"""Root conftest — auto-start a virtual display for headless test runs.

When no ``$DISPLAY`` is available (SSH session, container, CI without
an ``xvfb-run`` wrapper), pytest fails or errors on every UI/Tkinter
test that needs a display server.  Rather than requiring callers to
wrap ``pytest`` in ``xvfb-run -a``, this conftest transparently starts
a virtual framebuffer via ``PyVirtualDisplay`` when ``$DISPLAY`` is
unset so the full suite — UI and non-UI alike — runs green from a
bare ``pytest`` invocation.

The display lives for the entire process and is torn down on exit.
If PyVirtualDisplay or Xvfb is unavailable, the suite still runs —
UI tests will error with their normal ``_tkinter.TclError`` messages,
but the non-UI majority passes cleanly.  Install ``pyvirtualdisplay``
(``pip install pyvirtualdisplay``) and ``xvfb`` (system package) to
get the auto-bootstrap behaviour.
"""

from __future__ import annotations

import atexit
import os

_virtual_display = None  # type: ignore[var-annotated]


def _start_virtual_display() -> None:
    """Start a virtual X display if no ``$DISPLAY`` is set."""
    global _virtual_display
    if os.environ.get("DISPLAY"):
        return  # a display is already active
    try:
        from pyvirtualdisplay import Display
    except ImportError:
        return  # PyVirtualDisplay not installed
    try:
        _virtual_display = Display(visible=False, size=(1920, 1080))
        _virtual_display.start()
    except Exception:
        _virtual_display = None


_start_virtual_display()


@atexit.register
def _stop_virtual_display() -> None:
    """Tear down the virtual display on interpreter exit."""
    global _virtual_display
    if _virtual_display is not None:
        try:
            _virtual_display.stop()
        except Exception:
            pass
        _virtual_display = None
