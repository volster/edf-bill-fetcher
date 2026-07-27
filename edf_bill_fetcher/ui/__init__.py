"""UI submodule — Tkinter application classes.

Placeholder for the modularization refactor (Task 6). The full UI
class extraction (ReportOptionsDialog, App, build_ui) was deferred
because both classes are heavily coupled to module-level functions
in ``edf_collector.py`` (config loaders, evidence_engine factories,
run_cli_* entry points) that would need to move with them for
standalone operation.

For now, this module re-exports the UI classes from ``edf_collector``
so callers can use
``from edf_bill_fetcher.ui import App, ReportOptionsDialog``
without changing their code. The compat shim is removed by Task 7.
"""

from edf_collector import App, ReportOptionsDialog

__all__ = ["App", "ReportOptionsDialog"]
