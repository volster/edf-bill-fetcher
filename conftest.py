"""Use Qt's offscreen platform for tests on Linux without a desktop session."""

import os
import sys

if (
    sys.platform == "linux"
    and not os.environ.get("DISPLAY")
    and not os.environ.get("WAYLAND_DISPLAY")
):
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
