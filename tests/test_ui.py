"""The supported launch and UI imports must work without the retired Tk runtime."""

import subprocess
import sys


def test_cli_and_qt_import_without_loading_tk():
    result = subprocess.run(
        [
            sys.executable,
            "-c",
            "import sys; import main; import edf_bill_fetcher.ui.qt_app; assert 'tkinter' not in sys.modules; assert '_tkinter' not in sys.modules",
        ],
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert result.returncode == 0, result.stderr
