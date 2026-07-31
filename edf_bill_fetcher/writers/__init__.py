"""Writer functions for the EDF evidence workbook.

Extracted from ``edf_collector.py`` as part of the modularization
refactor (Task 5).  Each function writes one or more Excel sheets
using openpyxl.
"""

from __future__ import annotations

import gc
import glob
import json
import os
import pickle
import threading
import traceback
from datetime import date
from typing import Any, cast

import pandas as pd

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    HAS_TK = True
except ImportError:
    HAS_TK = False

try:
    import pypff

    HAS_PYPFF = True
except ImportError:
    HAS_PYPFF = False

try:
    import importlib.util

    HAS_SCIPY = importlib.util.find_spec("scipy") is not None
except ImportError:
    HAS_SCIPY = False

try:
    importlib.util.find_spec("statsmodels.tsa.holtwinters")

    HAS_STATSMODELS = True
except ImportError:
    HAS_STATSMODELS = False

try:
    HAS_PDF_REPORT = importlib.util.find_spec("edf_report") is not None
    HAS_DOCX_REPORT = importlib.util.find_spec("edf_report_docx") is not None
except ImportError:
    HAS_PDF_REPORT = False
    HAS_DOCX_REPORT = False

from openpyxl.chart import BarChart, LineChart, Reference  # noqa: F401
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: F401

from edf_bill_fetcher.helpers.date_utils import (  # noqa: E402,F401,I001
    _ISO_DATE_RE,
    _safe_to_datetime,
    parse_to_display_date,
    parse_to_sort_date,
    to_excel_date,
)
from edf_bill_fetcher.helpers.excel_utils import (  # noqa: E402,F401,I001
    _TEXT_SUPPRESSION_QUEUE,
    CELL_BORDER,
)
from edf_bill_fetcher.helpers.excel_utils import (  # noqa: F401
    open_pdf_hyperlink_cell as _open_pdf_hyperlink_cell,
)
from edf_bill_fetcher.helpers.formatting import (  # noqa: E402,F401,I001
    account_number_matches as _account_number_matches,
)
from edf_bill_fetcher.io.adapters.pdf import legal_context  # noqa: E402,F401,I001
from edf_bill_fetcher.io.writers.evidence import (  # noqa: E402,F401
    write_evidence_sheet,
    write_summary_sheet,
)
from edf_bill_fetcher.writers._helpers import (  # noqa: E402,F401,I001
    _SOURCE_PRECEDENCE,
    DUP_GREY,
    EDF_NAVY,
    EDF_OFFWHITE,
    EDF_ORANGE,
    EST_YELLOW,
    JUMP_RED,
    MEDIUM_GREY,
    _analyze_tariff_impact,
    _compute_volatility,
    _data_quality_report,
    _detect_payment_patterns,
    _disclosed_label,
    _holt_winters_forecast,
    _holt_winters_forecast_pair,
    _iqr_anomalies,
    _linear_forecast,
    _linear_forecast_pair,
    _parse_amount_for_event,
    _reading_type_to_aem,
    _recon_hyperlink,
    _zscore_anomalies,
    build_evidence_index,
    compute_dispute_flags,
    detect_sap_back_billing_events,
    match_sap_events_to_edf,
)
from evidence_bundle import build_bundle_index, save_evidence_files  # noqa: E402,F401

__all__ = [
    "_analyze_tariff_impact",
    "_compute_volatility",
    "_data_quality_report",
    "_detect_payment_patterns",
    "_disclosed_label",
    "_holt_winters_forecast",
    "_holt_winters_forecast_pair",
    "_iqr_anomalies",
    "_linear_forecast",
    "_linear_forecast_pair",
    "_recon_hyperlink",
    "_write_sap_bb_events_sheet",
    "_write_sap_bb_matches_sheet",
    "_write_sap_header_row",
    "build_evidence_index",
    "compute_dispute_flags",
    "detect_back_billing",
    "detect_sap_back_billing_events",
    "export_to_excel",
    "match_sap_events_to_edf",
    "write_back_billing_sheet",
    "write_contract_history_sheet",
    "write_data_quality_sheet",
    "write_evidence_sheet",
    "write_forecast_sheet",
    "write_meter_readings_sheet",
    "write_payment_analysis_sheet",
    "write_rebilling_sheet",
    "write_reconciliation_sheet",
    "write_sap_back_billing_sheets",
    "write_sap_contract_history_sheet",
    "write_sap_financial_transactions_sheet",
    "write_sap_meter_readings_sheet",
    "write_statistical_analysis_sheet",
    "write_summary_sheet",
    "write_tariff_analysis_sheet",
]


# ---------------------------------------------------------------------------
# A ``dict[str, int]`` mapping per-row signatures to the 1-indexed Excel row
# on the ``EDF Evidence Report`` sheet so the 4 analyser writers can emit a
# ``View on Evidence Report`` hotlink on each row. Two signatures per row are
# emitted:
#   - ``inv:<Invoice #>`` -- exact Invoice # match, used when the analyser
#     DataFrame carries the Invoice # in its key column.
#   - ``amt_days:<Amount £ with 2dp>|<days>`` -- fallback signature keyed on
#     the diagnostic pair `` (£, days-billed)``, used when Invoice # is N/A
#     or otherwise unparseable. ``setdefault`` keeps the first hit.














def run_analysers(df: pd.DataFrame) -> dict[str, Any]:
    """Run all Phase-2 detection analyses on the deduplicated
    DataFrame and return their outputs in a dict.

    The orchestrator is a thin wrapper so :func:`export_to_excel` can
    call four detectors with one line and downstream tests can
    inspect the full set without re-running each individually.

    Returns:
        dict with keys ``back_billing``, ``rebilling``,
        ``meter_rollover``, ``contracts``, ``evidence_index``. The
        first four are tidy DataFrames; ``evidence_index`` is a
        ``dict[str, int]`` mapping per-row signatures to the Excel row
        on the ``EDF Evidence Report`` sheet so the analyser tabs can
        emit a ``View on Evidence Report`` hotlink.
    """
    import edf_collector as _edc

    return {
        "back_billing": _edc.detect_back_billing(df),
        "rebilling": _edc.detect_rebilling(df, evidence_df=df),
        "meter_rollover": _edc.detect_meter_rollover(df),
        "contracts": _edc.infer_contracts(df),
        "evidence_index": build_evidence_index(df, header_row_offset=1),
    }






from edf_bill_fetcher.io.writers.sap import (  # noqa: E402,F401,I001
    _bb_invoice_value,
    _write_sap_bb_events_sheet,
    _write_sap_bb_matches_sheet,
    _write_sap_header_row,
    _write_sap_contract_history_sheet_impl,
    write_sap_back_billing_sheets,
    write_sap_contract_history_sheet,
    write_sap_financial_transactions_sheet,
    write_sap_meter_readings_sheet,
)

from edf_bill_fetcher.io.writers.reconciliation import (  # noqa: E402,F401,I001
    _recon_amount_to_float,
    _recon_parse_iso_date,
    write_reconciliation_sheet,
)

# GUI
# ---------------------------------------------------------------------------


class ReportOptionsDialog:
    """Modern report options dialog with format selection and section checkboxes."""

    SECTIONS = [
        ("cover", "Cover Page", True),
        ("toc", "Table of Contents", True),
        ("exec_summary", "Executive Summary", True),
        ("key_findings", "Key Findings", True),
        ("evidence_index", "Evidence Index", True),
        ("detailed_findings", "Detailed Findings", True),
        ("timeline", "Timeline", True),
        ("ofgem", "OFGEM Price Cap Comparison", True),
        ("statistical", "Statistical Analysis", True),
        ("payment", "Payment Analysis", True),
        ("forecast", "Forecast", True),
        ("data_quality", "Data Quality", True),
        ("tariff", "Tariff Impact Analysis", True),
        ("appendix_methodology", "Appendix: Methodology", True),
        ("appendix_glossary", "Appendix: Glossary", True),
        ("appendix_full_evidence", "Appendix: Full Evidence Table", True),
    ]

    def __init__(self, parent):
        self.parent = parent
        self.result = None
        self.dialog = None

    def show(self):
        """Show the dialog and return the selected options."""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Report Options")
        # Default size for 1080p: visible buttons without scrolling
        self.dialog.geometry("600x900")
        self.dialog.minsize(500, 500)
        self.dialog.resizable(True, True)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()

        # Center on parent
        self.dialog.update_idletasks()
        x = self.parent.winfo_rootx() + (self.parent.winfo_width() // 2) - 300
        y = self.parent.winfo_rooty() + (self.parent.winfo_height() // 2) - 450
        self.dialog.geometry(f"+{x}+{y}")

        self._build_ui()
        self.dialog.wait_window()
        return self.result

    def _build_ui(self):
        """Build the dialog UI."""
        # Create scrollable main area
        canvas = tk.Canvas(self.dialog, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.dialog, orient="vertical", command=canvas.yview)
        main = ttk.Frame(canvas, padding=20)

        main.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=main, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Bind mousewheel
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        if self.dialog is not None:
            self.dialog.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Also allow resizing canvas window width
        def _on_canvas_configure(event):
            canvas.itemconfig(canvas.find_all()[0], width=event.width)

        canvas.bind("<Configure>", _on_canvas_configure)

        # Header
        hdr = ttk.Frame(main)
        hdr.pack(fill=tk.X, pady=(0, 20))

        title_lbl = ttk.Label(
            hdr,
            text="Generate Ombudsman Report",
            font=("Calibri", 18, "bold"),
            foreground=EDF_NAVY,
        )
        title_lbl.pack(anchor=tk.W)

        subtitle = ttk.Label(
            hdr,
            text="Choose format and select sections to include",
            font=("Calibri", 10),
            foreground=MEDIUM_GREY,
        )
        subtitle.pack(anchor=tk.W, pady=(4, 0))

        ttk.Separator(main, orient="horizontal").pack(fill=tk.X, pady=(0, 16))

        # Format selection
        fmt_frame = ttk.LabelFrame(main, text=" Output Format ", padding=12)
        fmt_frame.pack(fill=tk.X, pady=(0, 16))

        self.format_var = tk.StringVar(value="both")
        formats = [
            ("both", "Both (PDF + Word)", "Generate both PDF and DOCX reports"),
            ("pdf", "PDF Only", "Professional PDF report (reportlab)"),
            ("docx", "Word Document Only", "Editable Word document (python-docx)"),
        ]

        for val, label, desc in formats:
            r = ttk.Frame(fmt_frame)
            r.pack(fill=tk.X, pady=3)
            rb = ttk.Radiobutton(r, variable=self.format_var, value=val)
            rb.pack(side=tk.LEFT)
            lbl_frame = ttk.Frame(r)
            lbl_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
            ttk.Label(lbl_frame, text=label, font=("Calibri", 10, "bold")).pack(anchor=tk.W)
            ttk.Label(lbl_frame, text=desc, font=("Calibri", 8), foreground=MEDIUM_GREY).pack(
                anchor=tk.W
            )

        ttk.Separator(main, orient="horizontal").pack(fill=tk.X, pady=(8, 16))

        # Section checkboxes
        sec_frame = ttk.LabelFrame(main, text=" Report Sections ", padding=12)
        sec_frame.pack(fill=tk.X, pady=(0, 16))

        # Select All / None buttons
        btn_frame = ttk.Frame(sec_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Button(btn_frame, text="Select All", command=self._select_all, width=12).pack(
            side=tk.LEFT
        )
        ttk.Button(btn_frame, text="Select None", command=self._select_none, width=12).pack(
            side=tk.LEFT, padx=(8, 0)
        )
        ttk.Button(btn_frame, text="Defaults", command=self._select_defaults, width=12).pack(
            side=tk.LEFT, padx=(8, 0)
        )

        # Checkboxes (main dialog is now scrollable, so no nested scrollbar needed)
        self.section_vars = {}
        for key, label, default in self.SECTIONS:
            var = tk.BooleanVar(value=default)
            self.section_vars[key] = var
            cb = ttk.Checkbutton(sec_frame, text=label, variable=var)
            cb.pack(anchor=tk.W, pady=1)

        ttk.Separator(main, orient="horizontal").pack(fill=tk.X, pady=(8, 16))

        # Action buttons
        action_frame = ttk.Frame(main)
        action_frame.pack(fill=tk.X)

        cancel_btn = ttk.Button(action_frame, text="Cancel", command=self._cancel, width=14)
        cancel_btn.pack(side=tk.RIGHT)

        ok_btn = tk.Button(
            action_frame,
            text="OK — Generate Report",
            bg=EDF_ORANGE,
            fg="white",
            font=("Calibri", 11, "bold"),
            command=self._generate,
            relief="flat",
            width=22,
        )
        ok_btn.pack(side=tk.RIGHT, padx=(0, 12))

        # Bind Enter key to OK, Escape to Cancel
        if self.dialog:
            self.dialog.bind("<Return>", lambda e: self._generate())
            self.dialog.bind("<Escape>", lambda e: self._cancel())

    def _select_all(self):
        for var in self.section_vars.values():
            var.set(True)

    def _select_none(self):
        for var in self.section_vars.values():
            var.set(False)

    def _select_defaults(self):
        for key, var in self.section_vars.items():
            # Find default from SECTIONS
            for k, _, default in self.SECTIONS:
                if k == key:
                    var.set(default)
                    break

    def _generate(self):
        """Collect results and close dialog."""
        selected_sections = [key for key, var in self.section_vars.items() if var.get()]
        if not selected_sections:
            messagebox.showwarning("No Sections", "Please select at least one report section.")
            return

        self.result = {
            "format": self.format_var.get(),
            "sections": selected_sections,
        }
        if self.dialog is not None:
            self.dialog.destroy()

    def _cancel(self):
        self.result = None
        if self.dialog is not None:
            self.dialog.destroy()


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("EDF Master Evidence Collector")
        self.root.geometry("780x860")
        self.root.configure(bg=EDF_OFFWHITE)

        self.pst_path = tk.StringVar()
        self.pdf_dir = tk.StringVar()
        self.htm_path = tk.StringVar()
        self.acc_num = tk.StringVar(value="")
        self.status = tk.StringVar(value="Ready.")
        self.progress_v = tk.DoubleVar(value=0)

        self.use_anchors = tk.BooleanVar(value=True)
        self.use_large = tk.BooleanVar(value=True)
        self.use_reading_class = tk.BooleanVar(value=True)
        self.use_pdf_fields = tk.BooleanVar(value=True)
        self.use_acc_filt = tk.BooleanVar(value=False)
        self.filter_below = tk.BooleanVar(value=True)
        self.save_filtered = tk.BooleanVar(value=True)
        self.use_dedup = tk.BooleanVar(value=True)
        self.save_dups = tk.BooleanVar(value=True)
        self.use_domain_filter = tk.BooleanVar(value=True)
        self.domain_filter = tk.StringVar(value="edfenergy.com")
        self.min_amount = tk.DoubleVar(value=500.0)
        self.analysis_min = tk.DoubleVar(value=500.0)
        self.output_name = tk.StringVar(value="EDF_Dispute_Evidence.xlsx")
        self.report_account_ref = tk.StringVar(value="")

        # New vars for UI refresh (see spec 2026-07-10-ui-refresh-design.md)
        self.output_folder = tk.StringVar(value="")
        self.amalgamate_duplicates = tk.BooleanVar(value=False)
        self.auto_generate_report = tk.BooleanVar(value=False)
        # Stream P5: save evidence files referenced by the workbook into a
        # flat ``output/evidence_files/`` folder and a themed DOCX index.
        # Defaults to True so the bundle is produced alongside the workbook
        # by default; reviewer can uncheck if they only want the XLSX.
        self.save_evidence_files_var = tk.BooleanVar(value=True)
        # Stream P1/P2 GUI toggles. SAP CSV-in-PDF data dumps render
        # their own dedicated sheets when "scan_sap_dumps" is set; the
        # cross-source Reconciliation sheet is independently controllable
        # via "generate_reconciliation_sheet" so a reviewer can keep the
        # SAP data without the cross-sheet matching view if desired.
        # Both default to True so the new sheets appear in the standard
        # extraction output; toggle off if the reviewer doesn't want
        # the legacy SAP dump analysis at all (e.g. on a clean run with
        # only invoice PDFs).
        self.scan_sap_dumps_var = tk.BooleanVar(value=True)
        self.generate_reconciliation_sheet_var = tk.BooleanVar(value=True)
        self._report_options: dict = {}
        self._CONFIG_PATH = os.path.expanduser("~/.edf_collector/config.json")

        # Load persisted config (may override the var defaults above)
        self._load_config()

        self.cancel_event = threading.Event()
        self.build_ui()

    # -- Config persistence --

    def _load_config(self):
        """Read config file and mutate tk vars via .set().

        Silently falls back to hardcoded defaults when the file is
        missing, unreadable, or malformed.
        """
        try:
            with open(self._CONFIG_PATH) as f:
                data = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            return

        gui = data.get("gui_state", {})
        _bool_keys: dict[str, tk.Variable] = {
            "use_anchors": self.use_anchors,
            "use_large": self.use_large,
            "use_reading_class": self.use_reading_class,
            "use_pdf_fields": self.use_pdf_fields,
            "use_acc_filt": self.use_acc_filt,
            "filter_below": self.filter_below,
            "save_filtered": self.save_filtered,
            "use_dedup": self.use_dedup,
            "save_dups": self.save_dups,
            "amalgamate_duplicates": self.amalgamate_duplicates,
            "use_domain_filter": self.use_domain_filter,
            "auto_generate_report": self.auto_generate_report,
            "save_evidence_files": self.save_evidence_files_var,
            "scan_sap_dumps": self.scan_sap_dumps_var,
            "generate_reconciliation_sheet": self.generate_reconciliation_sheet_var,
        }
        for key, var in _bool_keys.items():
            if key in gui:
                var.set(bool(gui[key]))

        _str_keys: dict[str, tk.Variable] = {
            "acc_num": self.acc_num,
            "domain_filter": self.domain_filter,
            "output_name": self.output_name,
            "report_account_ref": self.report_account_ref,
            "output_folder": self.output_folder,
        }
        for key, var in _str_keys.items():
            if key in gui:
                var.set(str(gui[key]))

        _float_keys: dict[str, tk.Variable] = {
            "min_amount": self.min_amount,
            "analysis_min": self.analysis_min,
        }
        for key, var in _float_keys.items():
            if key in gui:
                try:
                    var.set(float(gui[key]))
                except (ValueError, TypeError):
                    pass

        ro = data.get("report_options", {})
        if ro:
            self._report_options = ro

    def _save_config(self):
        """Persist GUI state + report options to config file atomically.

        Write to <path>.tmp, fsync, os.replace.  Permissions 0o600.
        """
        config_dir = os.path.dirname(self._CONFIG_PATH)
        os.makedirs(config_dir, exist_ok=True)

        gui = {
            "use_anchors": self.use_anchors.get(),
            "use_large": self.use_large.get(),
            "use_reading_class": self.use_reading_class.get(),
            "use_pdf_fields": self.use_pdf_fields.get(),
            "use_acc_filt": self.use_acc_filt.get(),
            "acc_num": self.acc_num.get(),
            "min_amount": self.min_amount.get(),
            "analysis_min": self.analysis_min.get(),
            "filter_below": self.filter_below.get(),
            "save_filtered": self.save_filtered.get(),
            "use_dedup": self.use_dedup.get(),
            "save_dups": self.save_dups.get(),
            "amalgamate_duplicates": self.amalgamate_duplicates.get(),
            "use_domain_filter": self.use_domain_filter.get(),
            "domain_filter": self.domain_filter.get(),
            "output_name": self.output_name.get(),
            "report_account_ref": self.report_account_ref.get(),
            "auto_generate_report": self.auto_generate_report.get(),
            "output_folder": self.output_folder.get(),
            "save_evidence_files": self.save_evidence_files_var.get(),
            "scan_sap_dumps": self.scan_sap_dumps_var.get(),
            "generate_reconciliation_sheet": self.generate_reconciliation_sheet_var.get(),
        }

        payload = {
            "output_folder": self.output_folder.get(),
            "report_options": getattr(self, "_report_options", {}),
            "gui_state": gui,
        }

        tmp_path = self._CONFIG_PATH + ".tmp"
        with open(tmp_path, "w") as f:
            json.dump(payload, f, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.chmod(tmp_path, 0o600)
        os.replace(tmp_path, self._CONFIG_PATH)

    def _resolve_output_path(
        self,
        stem: str,
        ext: str,
        batch_n: int | None = None,
        is_report: bool = False,
    ) -> str:
        """Build a sequential non-overwriting output path.

        Naming: {folder}/{stem}_{date}_{N}[{_Report}].{ext}
        If batch_n is passed, use it (shared counter for a batch).
        If None, scan folder for max existing N and use N+1.
        If output_folder is empty, falls back to os.getcwd().
        """
        folder = self.output_folder.get().strip() or os.getcwd()
        date_stamp = date.today().isoformat()
        suffix = "_Report" if is_report else ""

        if batch_n is not None:
            n = batch_n
        else:
            pattern = os.path.join(folder, f"{stem}_{date_stamp}_*{suffix}.{ext}")
            existing = glob.glob(pattern)
            max_n = 0
            for f in existing:
                basename = os.path.basename(f)
                prefix = f"{stem}_{date_stamp}_"
                rest = basename[len(prefix) :]
                if suffix:
                    rest = rest[: rest.index(suffix)]
                rest = rest.rsplit(".", 1)[0]
                if rest.isdigit():
                    max_n = max(max_n, int(rest))
            n = max_n + 1

        filename = f"{stem}_{date_stamp}_{n}{suffix}.{ext}"
        return os.path.join(folder, filename)

    def build_ui(self):
        hdr = tk.Frame(self.root, bg=EDF_ORANGE, height=60)
        hdr.pack(fill=tk.X)
        tk.Label(
            hdr,
            text="EDF BILLING EVIDENCE COLLECTOR",
            bg=EDF_ORANGE,
            fg="white",
            font=("Calibri", 14, "bold"),
        ).pack(pady=15)

        container = ttk.Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container, bg=EDF_OFFWHITE, highlightthickness=0)
        yscroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=yscroll.set)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        main = ttk.Frame(canvas, padding=16)
        cw = canvas.create_window((0, 0), window=main, anchor="nw")

        def _reconfig(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(cw, width=canvas.winfo_width())

        main.bind("<Configure>", _reconfig)
        canvas.bind("<Configure>", _reconfig)

        # --- Section 1: Source Data ---
        s1 = ttk.LabelFrame(main, text=" 1. Source Data ", padding=10)
        s1.pack(fill=tk.X, pady=5)

        def browse_row(parent, label, var, cmd):
            r = ttk.Frame(parent)
            r.pack(fill=tk.X, pady=2)
            ttk.Label(r, text=label, width=14).pack(side=tk.LEFT)
            ttk.Entry(r, textvariable=var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            ttk.Button(r, text="Browse", command=cmd).pack(side=tk.LEFT)

        browse_row(s1, "PST/OST File:", self.pst_path, self._pick_pst)
        browse_row(s1, "PDF Folder:", self.pdf_dir, self._pick_pdf_dir)
        browse_row(
            s1,
            "HTM Export:",
            self.htm_path,
            lambda: self.htm_path.set(
                filedialog.askopenfilename(filetypes=[("HTM/HTML", "*.htm *.html")])
            ),
        )

        # Output Folder picker (new - spec Design Section 1)
        browse_row(s1, "Output Folder:", self.output_folder, self._pick_output_folder)

        # Output filename row relocated from Section 2 to Section 1
        r_out = ttk.Frame(s1)
        r_out.pack(fill=tk.X, pady=2)
        ttk.Label(r_out, text="Output filename:", width=14).pack(side=tk.LEFT)
        ttk.Entry(r_out, textvariable=self.output_name, width=30).pack(side=tk.LEFT, padx=5)

        # --- Section 2: Extraction options ---
        s2 = ttk.LabelFrame(main, text=" 2. Search & Filter Options ", padding=10)
        s2.pack(fill=tk.X, pady=5)
        for text, var in [
            ("Smart Context Search", self.use_anchors),
            ("Large Number Fallback", self.use_large),
            ("Classify Reading Type", self.use_reading_class),
            ("Deep PDF Mine (kWh, standing charge, invoice #)", self.use_pdf_fields),
        ]:
            tk.Checkbutton(s2, text=text, variable=var, bg=EDF_OFFWHITE).pack(anchor=tk.W)

        r3 = ttk.Frame(s2)
        r3.pack(fill=tk.X, pady=4)
        tk.Checkbutton(
            r3, text="Filter by Account #:", variable=self.use_acc_filt, bg=EDF_OFFWHITE
        ).pack(side=tk.LEFT)
        ttk.Entry(r3, textvariable=self.acc_num, width=16).pack(side=tk.LEFT, padx=5)

        r3d = ttk.Frame(s2)
        r3d.pack(fill=tk.X, pady=4)
        tk.Checkbutton(
            r3d,
            text="Filter PST emails by sender domain:",
            variable=self.use_domain_filter,
            bg=EDF_OFFWHITE,
        ).pack(side=tk.LEFT)
        ttk.Entry(r3d, textvariable=self.domain_filter, width=40).pack(side=tk.LEFT, padx=5)
        ttk.Label(r3d, text="(comma-separated domains/addresses)", font=("Calibri", 8)).pack(
            side=tk.LEFT
        )

        r4 = ttk.Frame(s2)
        r4.pack(fill=tk.X, pady=2)
        chk_filt = tk.Checkbutton(
            r4, text="Filter results below minimum £:", variable=self.filter_below, bg=EDF_OFFWHITE
        )
        chk_filt.pack(side=tk.LEFT)
        ttk.Entry(r4, textvariable=self.min_amount, width=8).pack(side=tk.LEFT, padx=5)

        r4c = ttk.Frame(s2)
        r4c.pack(fill=tk.X, pady=2)
        ttk.Label(r4c, text="Analysis threshold (£):", width=24).pack(side=tk.LEFT)
        ttk.Entry(r4c, textvariable=self.analysis_min, width=8).pack(side=tk.LEFT, padx=5)

        r4d = ttk.Frame(s2)
        r4d.pack(fill=tk.X, pady=2)
        ttk.Label(r4d, text="Report account reference:", width=24).pack(side=tk.LEFT)
        ttk.Entry(r4d, textvariable=self.report_account_ref, width=20).pack(side=tk.LEFT, padx=5)

        chk_sf = tk.Checkbutton(
            s2,
            text="Keep filtered-out records on side sheet (Filtered (Below Min))",
            variable=self.save_filtered,
            bg=EDF_OFFWHITE,
        )
        chk_sf.pack(anchor=tk.W, padx=20)

        def _update_sf_state() -> None:
            chk_sf.config(state="normal" if self.filter_below.get() else "disabled")

        chk_filt.config(command=_update_sf_state)
        _update_sf_state()

        # Auto-generate report after extraction (spec Design Section 2)
        tk.Checkbutton(
            s2,
            text="Auto-generate report after extraction",
            variable=self.auto_generate_report,
            bg=EDF_OFFWHITE,
            command=self._save_config,
        ).pack(anchor=tk.W)

        # Stream P5: save evidence files + themed DOCX bundle index alongside
        # the workbook (spec Design Section 2 + §7). Defaults True.
        tk.Checkbutton(
            s2,
            text="Save evidence files + bundle index (output/evidence_files + evidence_index.docx)",
            variable=self.save_evidence_files_var,
            bg=EDF_OFFWHITE,
            command=self._save_config,
        ).pack(anchor=tk.W)

        # Stream P1: detect + render the three SAP CSV-in-PDF data
        # dumps (Contract / Meter-Read / Financial-Transactions) on
        # their own dedicated sheets.
        tk.Checkbutton(
            s2,
            text="Scan SAP CSV-in-PDF data dumps (adds SAP Contract History / Meter Readings / Financial Transactions sheets)",
            variable=self.scan_sap_dumps_var,
            bg=EDF_OFFWHITE,
            command=self._save_config,
        ).pack(anchor=tk.W)

        # Stream P2: cross-source Reconciliation sheet (SAP rows vs
        # inferred analyser data). Independent of the SAP-scan
        # toggle so a reviewer can keep the SAP data without the
        # cross-source match view if they want only the raw SAP
        # signals.
        tk.Checkbutton(
            s2,
            text="Generate cross-source Reconciliation sheet (SAP vs inferred analyser rows)",
            variable=self.generate_reconciliation_sheet_var,
            bg=EDF_OFFWHITE,
            command=self._save_config,
        ).pack(anchor=tk.W)

        self.report_options_section2_btn = tk.Button(
            s2,
            text="Report Options...",
            bg=EDF_NAVY,
            fg="white",
            font=("Calibri", 10),
            command=self._open_report_options,
            relief="flat",
        )
        self.report_options_section2_btn.pack(anchor=tk.W, padx=20, pady=4)

        # --- Section 3: Deduplication (relabelled + amalgamate child) ---
        s3 = ttk.LabelFrame(main, text=" 3. Deduplication ", padding=10)
        s3.pack(fill=tk.X, pady=5)
        chk_dup = tk.Checkbutton(
            s3,
            text="Drop duplicates found across sources",
            variable=self.use_dedup,
            bg=EDF_OFFWHITE,
        )
        chk_dup.pack(anchor=tk.W)
        chk_sd = tk.Checkbutton(
            s3,
            text="Record dropped duplicates on side sheet (Duplicate Entries)",
            variable=self.save_dups,
            bg=EDF_OFFWHITE,
        )
        chk_sd.pack(anchor=tk.W, padx=20)
        chk_am = tk.Checkbutton(
            s3,
            text="Build hybrid row per duplicate cluster (merge columns from every sibling)",
            variable=self.amalgamate_duplicates,
            bg=EDF_OFFWHITE,
            command=self._save_config,
        )
        chk_am.pack(anchor=tk.W, padx=40)

        def _update_dedup_state() -> None:
            dedup_on = self.use_dedup.get()
            chk_sd.config(state="normal" if dedup_on else "disabled")
            chk_am.config(state="normal" if (dedup_on and self.save_dups.get()) else "disabled")

        def _update_amalgamate_state() -> None:
            chk_am.config(
                state="normal" if (self.use_dedup.get() and self.save_dups.get()) else "disabled"
            )

        chk_dup.config(command=_update_dedup_state)
        chk_sd.config(command=_update_amalgamate_state)
        _update_dedup_state()

        # --- Progress ---
        self.pb = ttk.Progressbar(main, mode="determinate", maximum=100, variable=self.progress_v)
        self.pb.pack(fill=tk.X, pady=10)
        ttk.Label(
            main, textvariable=self.status, foreground=EDF_NAVY, font=("Calibri", 11, "bold")
        ).pack()

        btns = ttk.Frame(main)
        btns.pack(fill=tk.X, pady=8)
        self.run_btn = tk.Button(
            btns,
            text="EXTRACT TO EXCEL",
            bg=EDF_ORANGE,
            fg="white",
            font=("Calibri", 12, "bold"),
            command=self.start_thread,
            relief="flat",
        )
        self.run_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)

        self.report_options_btn = tk.Button(
            btns,
            text="Report Options",
            bg=EDF_NAVY,
            fg="white",
            font=("Calibri", 12, "bold"),
            command=self._open_report_options,
            relief="flat",
            state="normal" if (HAS_PDF_REPORT or HAS_DOCX_REPORT) else "disabled",
        )
        self.report_options_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, padx=(8, 0))

        # Load Spreadsheet & Generate Report button
        self.load_report_btn = tk.Button(
            btns,
            text="LOAD & REPORT",
            bg=EDF_ORANGE,
            fg="white",
            font=("Calibri", 12, "bold"),
            command=self.load_spreadsheet_and_report,
            relief="flat",
        )
        self.load_report_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, padx=(8, 0))

    # -- Helpers --

    def _pick_pst(self):
        p = filedialog.askopenfilename(filetypes=[("Mail Stores", "*.pst *.ost")])
        if p:
            self.pst_path.set(p)

    def _pick_pdf_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.pdf_dir.set(p)

    def _pick_output_folder(self):
        p = filedialog.askdirectory()
        if p:
            self.output_folder.set(p)
            self._save_config()

    def _open_report_options(self):
        """Open ReportOptionsDialog and persist selection on OK."""
        dialog = ReportOptionsDialog(self.root)
        options = dialog.show()
        if options:
            self._report_options = options
            self._save_config()

    def set_status(self, text):
        def _apply():
            self.status.set(text)
            self.root.update_idletasks()

        if threading.current_thread() is threading.main_thread():
            _apply()
        else:
            self.root.after(0, _apply)

    def set_progress(self, current, total, text=None):
        pct = max(0, min(100, (current / total) * 100)) if total else 0

        def _apply():
            self.progress_v.set(pct)
            if text:
                self.status.set(text)

        if threading.current_thread() is threading.main_thread():
            _apply()
        else:
            self.root.after(0, _apply)

    def _show(self, level, title, text):
        def _s():
            if level == "info":
                messagebox.showinfo(title, text)
            elif level == "warning":
                messagebox.showwarning(title, text)
            else:
                messagebox.showerror(title, text)

        if threading.current_thread() is threading.main_thread():
            _s()
        else:
            self.root.after(0, _s)

    def _finish(self):
        self._set_extract_idle()
        self.progress_v.set(0)
        self.set_status("Cancelled." if self.cancel_event.is_set() else "Ready.")
        gc.collect()

    def _set_extract_idle(self):
        """Flip run_btn to Idle: orange, EXTRACT TO EXCEL."""
        self.run_btn.config(
            text="EXTRACT TO EXCEL",
            bg=EDF_ORANGE,
            fg="white",
            command=self.start_thread,
            state="normal",
        )

    def _set_extract_running(self):
        """Flip run_btn to Running: navy, Cancel."""
        self.run_btn.config(
            text="Cancel",
            bg=EDF_NAVY,
            fg="white",
            command=self._cancel,
            state="normal",
        )

    def _set_extract_cancelling(self):
        """Flip run_btn to Cancelling: grey, Cancelling..."""
        self.run_btn.config(
            text="Cancelling...",
            bg=MEDIUM_GREY,
            fg="white",
            state="disabled",
        )

    def load_spreadsheet_and_report(self):
        """Load records from an existing spreadsheet and auto-generate reports.

        Assumes the spreadsheet has standard EDF Evidence Report format with
        an 'EDF Evidence Report' sheet.  Sequential-named reports written
        into output_folder (or the picked file's directory if unset).
        """
        if not HAS_PDF_REPORT and not HAS_DOCX_REPORT:
            self._show(
                "error",
                "Report Unavailable",
                "Report generation requires 'reportlab' (PDF) and/or 'python-docx' (Word).\n"
                "Install with: pip install reportlab python-docx",
            )
            return

        file_path = filedialog.askopenfilename(
            initialdir=self.output_folder.get() or os.getcwd(),
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            title="Select EDF Evidence Report Spreadsheet",
        )
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path, sheet_name="EDF Evidence Report")
            if df.empty:
                self._show(
                    "warning",
                    "No Data",
                    "The selected spreadsheet has no records in 'EDF Evidence Report' sheet.",
                )
                return

            records = df.to_dict("records")
            ro = getattr(self, "_report_options", {})
            fmt = ro.get("format", "both")
            sections = ro.get("sections", [s[0] for s in ReportOptionsDialog.SECTIONS])

            base_dir = self.output_folder.get().strip() or os.path.dirname(file_path)
            self.output_folder.set(base_dir)
            stem = os.path.basename(file_path).replace(".xlsx", "")

            output_paths: dict[str, str] = {}
            if fmt in ("pdf", "both") and HAS_PDF_REPORT:
                output_paths["pdf"] = self._resolve_output_path(stem, "pdf", is_report=True)
            if fmt in ("docx", "both") and HAS_DOCX_REPORT:
                output_paths["docx"] = self._resolve_output_path(
                    stem, "docx", batch_n=1, is_report=True
                )

            if not output_paths:
                self._show(
                    "warning",
                    "No Reports",
                    "No report paths resolved (check pdf/docx availability).",
                )
                return

            self.set_status("Generating report…")
            self.load_report_btn.config(state="disabled")

            config = {
                "min_amount": self.min_amount.get(),
                "analysis_min": self.analysis_min.get(),
                "acc_num": self.acc_num.get(),
                "report_account_ref": self.report_account_ref.get().strip(),
                "report_sections": sections,
            }

            from dataclasses import dataclass

            @dataclass
            class MockEngine:
                records: list
                filtered_records: list
                pdf_count: int
                email_count: int
                error_log: list

            engine = MockEngine(
                records=records, filtered_records=[], pdf_count=0, email_count=0, error_log=[]
            )

            def _generate():
                from edf_report import generate_pdf_from_gui
                from edf_report_docx import generate_docx_from_gui

                try:
                    msgs = []
                    if "pdf" in output_paths:
                        s, m = generate_pdf_from_gui(
                            records=records,
                            output_path=output_paths["pdf"],
                            config=config,
                            engine=engine,
                            filtered=[],
                        )
                        msgs.append(("PDF", s, m))
                    if "docx" in output_paths:
                        s, m = generate_docx_from_gui(
                            records=records,
                            output_path=output_paths["docx"],
                            config=config,
                            engine=engine,
                            filtered=[],
                        )
                        msgs.append(("DOCX", s, m))

                    combined = []
                    all_ok = True
                    for label, ok, m in msgs:
                        if ok:
                            combined.append(
                                f"✓ {label}: {m.split(chr(10))[-1] if m else 'Generated'}\n{output_paths[label.lower()]}"
                            )
                        else:
                            all_ok = False
                            self.root.after(
                                0,
                                lambda mn=m, lb=label: self._show(
                                    "error", f"{lb} Generation Failed", mn
                                ),
                            )
                    if all_ok and combined:
                        self.root.after(
                            0,
                            lambda c=combined: self._show(
                                "info", "Reports Generated", "\n\n".join(c)
                            ),
                        )
                except Exception as e:
                    self.root.after(
                        0,
                        lambda err=e: self._show("error", "Error", f"An error occurred:\n\n{err}"),
                    )
                finally:
                    self.root.after(
                        0,
                        lambda: (
                            self.load_report_btn.config(state="normal"),
                            self.set_status("Ready."),
                        ),
                    )

            threading.Thread(target=_generate, daemon=True).start()

        except Exception as e:
            self._show("error", "Load Error", f"Failed to load spreadsheet:\n\n{e}")

    def _cancel(self):
        self.cancel_event.set()
        self._set_extract_cancelling()
        self.set_status("Cancelling…")

    def start_thread(self):
        try:
            self.min_amount.get()
            self.analysis_min.get()
        except Exception:
            messagebox.showerror(
                "Error", "Minimum amount and analysis threshold must be valid numbers."
            )
            return

        has_sources = any(
            [
                self.pst_path.get().strip(),
                self.pdf_dir.get().strip(),
                self.htm_path.get().strip(),
            ]
        )
        if not has_sources:
            messagebox.showerror(
                "Error",
                "Please select at least one source:\nPST/OST file, PDF folder, or HTM export.",
            )
            return
        self.cancel_event.clear()
        self._set_extract_running()
        self.progress_v.set(0)
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        config = {
            "use_anchors": self.use_anchors.get(),
            "use_large": self.use_large.get(),
            "use_reading_classification": self.use_reading_class.get(),
            "use_pdf_fields": self.use_pdf_fields.get(),
            "use_acc_filter": self.use_acc_filt.get(),
            "acc_num": self.acc_num.get(),
            "min_amount": self.min_amount.get(),
            "analysis_min": self.analysis_min.get(),
            "report_account_ref": self.report_account_ref.get().strip(),
            "filter_below": self.filter_below.get(),
            "save_filtered": self.save_filtered.get(),
            "use_dedup": self.use_dedup.get(),
            "save_dups": self.save_dups.get(),
            "amalgamate_duplicates": self.amalgamate_duplicates.get(),
            "use_domain_filter": self.use_domain_filter.get(),
            "domain_filter": self.domain_filter.get().strip(),
            # Stream P1/P2 toggles -- threaded through to
            # export_to_excel which gates SAP sheet writes + the
            # Reconciliation sheet.
            "save_evidence_files": self.save_evidence_files_var.get(),
            "scan_sap_dumps": self.scan_sap_dumps_var.get(),
            "generate_reconciliation_sheet": self.generate_reconciliation_sheet_var.get(),
        }

        from edf_collector import EvidenceEngine  # noqa: F401,E402

        engine = EvidenceEngine(config, self.set_status, self.set_progress, self.cancel_event)
        self.engine = engine

        try:
            pst_path = self.pst_path.get().strip()
            if pst_path and os.path.exists(pst_path) and not self.cancel_event.is_set():
                if not HAS_PYPFF:
                    self._show("warning", "PST", "pypff not installed — PST/OST scanning skipped.")
                else:
                    self.set_status("Scanning PST/OST…")
                    try:
                        pff = pypff.file()
                    except AttributeError:
                        pff = getattr(pypff, "File", None)
                        if pff is None:
                            raise AttributeError(
                                "pypff module has no 'file' or 'File' attribute"
                            ) from None
                        pff = pff()
                    pff.open(os.path.abspath(pst_path))
                    try:
                        engine.crawl_pst(pff.get_root_folder())
                    finally:
                        pff.close()

            htm_path = self.htm_path.get().strip()
            if htm_path and os.path.exists(htm_path) and not self.cancel_event.is_set():
                self.set_status("Parsing HTM account history…")
                engine.process_htm_file(htm_path)

            pdf_path = self.pdf_dir.get().strip()
            if pdf_path and os.path.exists(pdf_path) and not self.cancel_event.is_set():
                engine.crawl_local_pdfs(pdf_path)

            if self.cancel_event.is_set():
                self._show("warning", "Cancelled", "Extraction cancelled.")
                return

            if engine.records:
                self.set_status("Writing Excel report…")
                # Fall back to source dir when output_folder unset
                if not self.output_folder.get().strip():
                    base_dir = (
                        os.path.dirname(pst_path)
                        if pst_path
                        else pdf_path
                        if pdf_path
                        else os.path.dirname(htm_path)
                        if htm_path
                        else os.getcwd()
                    )
                    self.output_folder.set(base_dir)
                stem = self.output_name.get().strip() or "EDF_Dispute_Evidence"
                if stem.lower().endswith(".xlsx"):
                    stem = stem[:-5]
                xlsx_path = self._resolve_output_path(stem, "xlsx")
                export_to_excel(  # type: ignore  # noqa: F821
                    engine.records,
                    xlsx_path,
                    engine.error_log,
                    config,
                    filtered=engine.filtered_records,
                    sap_rows={
                        "contract": engine.sap_contract_rows,
                        "meter": engine.sap_meter_rows,
                        "financial": engine.sap_financial_rows,
                    },
                )
                self._save_config()
                summary = (
                    f"Extraction complete.\n\n"
                    f"  Emails matched: {engine.email_count}\n"
                    f"  PDFs processed: {engine.pdf_count}\n"
                    f"  Records found:  {len(engine.records)}\n"
                )
                if engine.error_log:
                    summary += f"\n  Parse errors: {len(engine.error_log)} (see Parse Errors tab)"
                summary += f"\n\nSaved to:\n{xlsx_path}"

                # Stream P5: save evidence files + themed DOCX bundle index
                # into a sibling ``evidence_files/`` folder when the toggle is
                # set on (default True).
                if self.save_evidence_files_var.get():
                    try:
                        import pandas as pd

                        out_dir = os.path.dirname(xlsx_path) or os.getcwd()
                        ev_dir = os.path.join(out_dir, "evidence_files")
                        dfc = pd.DataFrame(engine.records)
                        # Build the source-paths reverse-lookup from the
                        # crawl attribute the engine carries internally.
                        source_paths = getattr(engine, "source_paths", {}) or {}
                        saved = save_evidence_files(dfc, source_paths, ev_dir)
                        index_docx = os.path.join(out_dir, "evidence_index.docx")
                        build_bundle_index(
                            dfc, saved, index_docx, account=str(config.get("acc_num", ""))
                        )
                        summary += f"\n\nSaved {len(saved)} evidence files to:\n{ev_dir}"
                        summary += f"\nBundle index:\n{index_docx}"
                    except Exception as bundle_err:
                        # Don't lose the run if the bundle step blows up --
                        # log it loudly but still keep the XLSX.
                        self._show(
                            "warning",
                            "Bundle step failed",
                            (
                                f"Evidence file save failed:\n{bundle_err}"
                                f"\n\nThe XLSX workbook is still saved at:\n{xlsx_path}"
                            ),
                        )

                if self.auto_generate_report.get():
                    report_paths = self._run_auto_report(engine, stem, 1)
                    if report_paths:
                        summary += "\n\nReports:\n" + "\n".join(report_paths)

                self._show("info", "Success", summary)
            else:
                self._show(
                    "warning",
                    "No Data",
                    "No billing amounts found.\n\nTips:\n"
                    "• Uncheck the Account Filter\n"
                    "• Lower the minimum threshold\n"
                    "• Check your source files contain EDF billing data",
                )

        except Exception:
            self._show("error", "Error", f"An error occurred:\n\n{traceback.format_exc()}")
        finally:
            self.root.after(0, self._finish)

    def _run_auto_report(self, engine, stem, batch_n):
        """Run report generation for the auto-generate flow.

        Uses persisted _report_options; writes to output_folder;
        returns list of written paths.
        """
        from edf_report import generate_pdf_from_gui
        from edf_report_docx import generate_docx_from_gui

        ro = getattr(self, "_report_options", {})
        fmt = ro.get("format", "both")
        sections = ro.get("sections", [s[0] for s in ReportOptionsDialog.SECTIONS])

        config = {
            "min_amount": self.min_amount.get(),
            "analysis_min": self.analysis_min.get(),
            "acc_num": self.acc_num.get(),
            "report_account_ref": self.report_account_ref.get().strip(),
            "report_sections": sections,
        }

        written: list[str] = []
        if fmt in ("pdf", "both") and HAS_PDF_REPORT:
            pdf_path = self._resolve_output_path(stem, "pdf", batch_n=batch_n, is_report=True)
            success, _ = generate_pdf_from_gui(
                records=engine.records,
                output_path=pdf_path,
                config=config,
                engine=engine,
                filtered=engine.filtered_records,
            )
            if success:
                written.append(pdf_path)

        if fmt in ("docx", "both") and HAS_DOCX_REPORT:
            docx_path = self._resolve_output_path(stem, "docx", batch_n=batch_n, is_report=True)
            success, _ = generate_docx_from_gui(
                records=engine.records,
                output_path=docx_path,
                config=config,
                engine=engine,
                filtered=engine.filtered_records,
            )
            if success:
                written.append(docx_path)

        return written


# ---------------------------------------------------------------------------
# Safe pickle deserialiser — prevents arbitrary code execution when loading
# engine-data pickle files from disk.  Only standard built-in types and
# the project's own EvidenceEngine class are allowed through; anything
# else raises UnpicklingError.
# ---------------------------------------------------------------------------


class _RestrictedUnpickler(pickle.Unpickler):
    """Unpickler that only allows known-safe types.

    Permits: built-in scalars, dicts, lists, tuples, sets, frozensets,
    bytes/bytearray, and the project's own ``EvidenceEngine``.  Everything
    else triggers ``pickle.UnpicklingError`` so a crafted pickle can never
    import and call arbitrary code.
    """

    # Module→class whitelist.  Only classes listed here can be rebuilt.
    # A whitelist value of ``None`` (as opposed to the usual
    # ``set[str]`` of permitted class names) is interpreted as
    # "the entire module is trusted".  We only use this for
    # ``pyarrow.lib`` whose exposed pickle surface is purely
    # restoration-callable ``_something`` functions, never
    # ``os.system`` / ``subprocess.Popen``.  Every other
    # whitelist entry is an explicit set of class names.
    #
    # Note ``dict.get(key)`` returns ``None`` for both "key
    # absent" and "key present with value None" — we therefore
    # distinguish via the sentinel object below rather than
    # raw ``is None`` comparison.
    _SAFE_CLASSES: dict[str, set[str] | None] = {
        "builtins": {
            "dict",
            "list",
            "tuple",
            "set",
            "frozenset",
            "int",
            "float",
            "str",
            "bool",
            "bytes",
            "bytearray",
            "NoneType",
            "type",
            "slice",
        },
        "collections": {"OrderedDict", "defaultdict", "Counter", "deque"},
        "collections.__init__": {"OrderedDict", "defaultdict", "Counter", "deque"},
        "pandas.core.series": {"Series"},
        "pandas.core.frame": {"DataFrame"},
        # NOTE: newer pandas releases have relocated these classes
        # under ``pandas.*.frame`` / ``pandas.*.series`` submodules
        # depending on the wheel build.  Whitelist both the original
        # canonical paths and the ``pandas.*`` alias so a round-trip
        # works regardless of which path the running pandas 2.x
        # resolves the class through.
        "pandas": {"DataFrame", "Series", "Index", "StringDtype", "RangeIndex"},
        # Pandas 2.x stores string columns as ``ArrowStringArray``
        # via the Arrow backend (the legacy ``numpy.object_`` path
        # was deprecated).  The pickle protocol resolves this
        # through ``pandas.arrays`` rather than ``pandas.core.*``,
        # so we whitelist the runtime module path explicitly.
        "pandas.arrays": {"ArrowStringArray"},
        # The Arrow backend itself (``pyarrow.lib``) is a transitive
        # dependency of pandas 2.x and is not a sandboxing risk —
        # allowing arbitrary Python objects to land via pyarrow
        # would require the user to have actively installed
        # pyarrow *and* crafted a malicious data file, after
        # which the unpickler still has to resolve the class.
        # We grant the *entire* ``pyarrow.lib`` surface here so
        # any pandas 2.x Arrow-backed string column round-trips
        # cleanly without our having to keep this list current
        # every time the pyarrow release rotates a private name.
        # The cost is a slightly-bigger whitelist; the safety is
        # unchanged because pyarrow.lib's exposed API is only
        # ``_scalar_to_array``/``_restore_array``-style restoration
        # routines, never ``os.system`` or ``subprocess.Popen``.
        "pyarrow.lib": None,
        # Phase 1.4: ``BlockManager`` is the internal layout primitive
        # that pandas 2.x uses to back every ``DataFrame`` /
        # ``Series``.  Without it, a pickle of a ``records`` list
        # containing a DataFrame falls back to "Can't pickle local
        # object" or "Blocked unsafe class ... BlockManager"
        # depending on whether the unpickler bails before/after
        # resolving the type.  Phase 1.4 acceptance: pin the round-trip
        # of a real engine whose ``engine.records`` includes a
        # ``pandas.DataFrame`` — see tests/test_pickle_roundtrip.py.
        "pandas.core.internals.managers": {"BlockManager"},
        # Phase 1.4: pandas's ``_unpickle_block`` is the C-extension
        # helper that ``BlockManager.__setstate__`` falls through to
        # when materialising ``Block`` objects from a pickled stream.
        # Without it the BlockManager round-trip falls back to
        # "Blocked unsafe class ... _unpickle_block".  BlockManager
        # itself is a thin Python wrapper around this C-level loader,
        # so both are required for a clean DataFrame round-trip.
        "pandas._libs.internals": {"_unpickle_block"},
        # Phase 1.4: numpy's ``_frombuffer`` is the C-extension helper
        # used by ``ndarray.__reduce__`` to round-trip the raw byte
        # buffer that holds the array alongside a type descriptor.
        # ``ndarray`` was already on the whitelist; this lets the
        # byte-buffer half survive the round-trip.
        # NOTE: numpy >= 2.0 moved this to ``numpy._core.numeric``; keep
        # both paths for backward/forward compatibility.
        "numpy.core.numeric": {"_frombuffer"},
        "numpy._core.numeric": {"_frombuffer"},
        # Phase 1.4: ``numpy.dtype`` is the scalar-type descriptor
        # every ndarray carries — without it a round-tripped
        # ndarray raises "Object has no attribute 'itemsize'".
        # Whitelist the dedicated ``dtype`` module too, alongside
        # the existing ndarray entries.
        "numpy.dtype": {"dtype"},
        "numpy": {"ndarray", "dtype"},
        "numpy.ndarray": {"ndarray"},
        # Phase 1.4: ``_reconstruct`` is the C-extension helper that
        # rebuilds an ``ndarray`` of a given shape/dtype from the
        # pickle-encoded byte buffer.  Without this entry, a
        # round-tripped 2D ``numpy.ndarray`` (under the bonnet of
        # every ``pandas.DataFrame``) fails with
        # "Blocked unsafe class 'numpy.core.multiarray'.'_reconstruct'".
        # NOTE: numpy >= 2.0 moved this to ``numpy._core.multiarray``;
        # keep both paths for compatibility.
        "numpy.core.multiarray": {"_reconstruct"},
        "numpy._core.multiarray": {"_reconstruct"},
        # Phase 1.4: ``_new_Index`` rebuilds a pandas Index from a
        # pickled (dtype, kind) tuple — needed because the persistent
        # RangeIndex(DataFrame.index) carries a ``kind`` token.  The
        # public ``Index`` class is the parent that ``_new_Index``
        # instantiates; both need to be on the whitelist for the
        # round-trip to construct a fully-fledged ``Index`` after
        # the C-extension helper has built its layout.
        "pandas.core.indexes.base": {"_new_Index", "Index"},
        # Phase 1.4: ``RangeIndex`` is the integer-only Index
        # subclass that pandas DataFrames grow by default.  Without
        # it the round-trip works for ``Index`` but raises
        # "Blocked unsafe class 'pandas.core.indexes.range'.'RangeIndex'"
        # on the most common case.  ``RangeIndex.__init__`` is a thin
        # wrapper so this single entry is sufficient.
        "pandas.core.indexes.range": {"RangeIndex"},
        # Our own classes — needed for persisted engine objects
        # NOTE: "__main__" was previously allowed but is a security risk —
        # it would permit any user script named EvidenceEngine to be
        # unpickled.  The proper module path "edf_collector" is the only
        # legitimate source for this class.
        "edf_collector": {"EvidenceEngine"},
    }

    def find_class(self, module: str, name: str) -> type:
        """Resolve ``module.name`` from the explicit whitelist only.

        A whitelist value of ``None`` (as opposed to the usual
        ``set[str]`` of permitted class names) is interpreted as
        "the entire module is trusted".  We only use this for
        ``pyarrow.lib`` whose exposed pickle surface is purely
        restoration-callable ``_something`` functions, never
        ``os.system`` / ``subprocess.Popen``.  Every other
        whitelist entry is an explicit set of class names.

        Note ``dict.get(key)`` returns ``None`` for both "key
        absent" and "key present with value None" — we therefore
        distinguish via the sentinel object below rather than
        raw ``is None`` comparison.
        """
        _SENTINEL = object()  # used purely to disambiguate "absent" vs "None"
        allowed = self._SAFE_CLASSES.get(module, _SENTINEL)
        # Module not in whitelist → blocked.
        if allowed is _SENTINEL:
            raise pickle.UnpicklingError(
                f"Blocked unsafe class {module!r}.{name!r} in pickle stream"
            )
        # Whole-module permission (``None`` value) → allow.
        # Per-name permission (``set`` value) → check membership.
        # Use ``allow_everything = allowed is None`` to drive the
        # control flow explicitly so mypy can narrow the type
        # from ``set[str] |`` to ``None`` at the call sites
        # without resorting to ``cast``.
        allow_everything = allowed is None
        if allow_everything or (isinstance(allowed, set) and name in allowed):
            if module == "edf_collector":
                import importlib

                mod: Any = importlib.import_module("edf_collector")
                cls: Any = getattr(mod, name)
                if not isinstance(cls, type):
                    raise pickle.UnpicklingError(
                        f"Resolved edf_collector attribute {name!r} is not a class"
                    )
                return cls
            return cast(type, super().find_class(module, name))
        raise pickle.UnpicklingError(f"Blocked unsafe class {module!r}.{name!r} in pickle stream")


def _safe_pickle_load(path: str) -> Any:
    """Load a pickle file through the restricted unpickler.

    Usage:  obj = _safe_pickle_load("engine.pkl")
    Raises pickle.UnpicklingError for disallowed types.
    """
    with open(path, "rb") as f:
        return _RestrictedUnpickler(f).load()


def run_cli_extract(args: list[str]) -> None:
    """Run extraction from command line (headless mode)."""
    import argparse
    import json
    import os
    import sys

    parser = argparse.ArgumentParser(
        description="Extract EDF billing data from PST/OST, PDF folder, or HTM export",
        prog="edf-collector --extract",
    )
    parser.add_argument("--pst", help="Path to PST/OST file")
    parser.add_argument("--pdf-dir", help="Path to directory containing PDF bills")
    parser.add_argument("--htm", help="Path to HTM account history export")
    parser.add_argument("--output", "-o", required=True, help="Output Excel file path")
    parser.add_argument("--records-json", help="Also save extracted records as JSON")
    parser.add_argument("--config", "-c", help="Path to config JSON file (optional)")
    parser.add_argument("--acc-filter", help="Filter by account number (e.g., A-12345678)")
    parser.add_argument(
        "--domain-filter",
        default="edfenergy.com",
        help="Comma-separated sender domains for PST filtering",
    )
    parser.add_argument("--min-amount", type=float, default=500.0, help="Minimum amount threshold")
    parser.add_argument("--no-dedup", action="store_true", help="Disable deduplication")
    parser.add_argument("--no-anchors", action="store_true", help="Disable smart context search")
    parser.add_argument("--no-large", action="store_true", help="Disable large amount fallback")
    parser.add_argument(
        "--no-reading-class", action="store_true", help="Disable reading classification"
    )
    parser.add_argument(
        "--no-pdf-fields", action="store_true", help="Disable deep PDF field extraction"
    )
    parser.add_argument(
        "--no-filter-below", action="store_true", help="Don't filter records below minimum amount"
    )
    parsed = parser.parse_args(args)

    # Check at least one source
    if not any([parsed.pst, parsed.pdf_dir, parsed.htm]):
        sys.stderr.write("ERROR: At least one source required (--pst, --pdf-dir, or --htm)\n")
        sys.exit(1)

    # Load config from file if provided
    config = {}
    if parsed.config:
        try:
            with open(parsed.config, encoding="utf-8") as f:
                config = json.load(f)
        except Exception as e:
            sys.stderr.write(f"ERROR: Failed to load config: {e}\n")
            sys.exit(1)

    # Override with CLI args
    config.update(
        {
            "use_acc_filter": bool(parsed.acc_filter),
            "acc_num": parsed.acc_filter or "",
            "use_domain_filter": True,
            "domain_filter": parsed.domain_filter,
            "min_amount": parsed.min_amount,
            "filter_below": not parsed.no_filter_below,
            "use_dedup": not parsed.no_dedup,
            "use_anchors": not parsed.no_anchors,
            "use_large": not parsed.no_large,
            "use_reading_classification": not parsed.no_reading_class,
            "use_pdf_fields": not parsed.no_pdf_fields,
            "save_filtered": True,
            "save_dups": True,
        }
    )

    # Check PST dependency
    if parsed.pst and not HAS_PYPFF:
        sys.stderr.write(
            "ERROR: PST/OST support requires 'libpff-python'. Install with: pip install libpff-python\n"
        )
        sys.exit(1)

    from edf_collector import EvidenceEngine  # noqa: F401,E402

    engine = EvidenceEngine(config, print, None, None)

    try:
        if parsed.pst and os.path.exists(parsed.pst):
            print(f"Scanning PST/OST: {parsed.pst}")
            try:
                pff = pypff.file()
            except AttributeError:
                pff = getattr(pypff, "File", None)
                if pff is None:
                    raise AttributeError("pypff module has no 'file' or 'File' attribute") from None
                pff = pff()
            pff.open(os.path.abspath(parsed.pst))
            try:
                engine.crawl_pst(pff.get_root_folder())
            finally:
                pff.close()

        if parsed.htm and os.path.exists(parsed.htm):
            print(f"Parsing HTM: {parsed.htm}")
            engine.process_htm_file(parsed.htm)

        if parsed.pdf_dir and os.path.exists(parsed.pdf_dir):
            print(f"Scanning PDF folder: {parsed.pdf_dir}")
            engine.crawl_local_pdfs(parsed.pdf_dir)

        if not engine.records:
            sys.stderr.write("WARNING: No billing records found\n")
            sys.exit(1)

        # Export to Excel
        print(f"Writing Excel report: {parsed.output}")
        export_to_excel(  # type: ignore  # noqa: F821
            engine.records,
            parsed.output,
            engine.error_log,
            config,
            filtered=engine.filtered_records,
            sap_rows={
                "contract": engine.sap_contract_rows,
                "meter": engine.sap_meter_rows,
                "financial": engine.sap_financial_rows,
            },
        )

        # Optionally save records as JSON
        if parsed.records_json:
            import datetime

            output_data = {
                "extracted_at": datetime.datetime.now().isoformat(),
                "config": config,
                "records": engine.records,
                "filtered_records": engine.filtered_records,
                "error_log": engine.error_log,
            }
            with open(parsed.records_json, "w", encoding="utf-8") as f:
                json.dump(output_data, f, indent=2, default=str)
            print(f"Records saved as JSON: {parsed.records_json}")

        print("Extraction complete!")
        print(f"  PDFs processed: {engine.pdf_count}")
        print(f"  Emails matched: {engine.email_count}")
        print(f"  Records found:  {len(engine.records)}")
        if engine.error_log:
            print(f"  Parse errors:   {len(engine.error_log)}")

    except Exception as e:
        sys.stderr.write(f"ERROR: {e}\n")
        import traceback

        traceback.print_exc()
        sys.exit(1)


def run_cli_pdf_report(args: list[str]) -> None:
    """Run PDF report generation from command line."""
    import argparse
    import json
    import sys

    from edf_report import generate_pdf_from_gui

    parser = argparse.ArgumentParser(
        description="Generate PDF report from extracted records",
        prog="edf-collector --pdf-report",
    )
    parser.add_argument(
        "--records",
        "-i",
        required=True,
        help="Path to extracted records JSON file (exported from GUI or script)",
    )
    parser.add_argument("--output", "-o", required=True, help="Output PDF file path")
    parser.add_argument("--config", "-c", help="Path to config JSON file (optional)")
    parser.add_argument(
        "--engine-data",
        "-e",
        help="Path to engine data pickle file (optional, for filtered records)",
    )
    parsed = parser.parse_args(args)

    try:
        with open(parsed.records, encoding="utf-8") as f:
            loaded = json.load(f)

        # Accept either a bare list of records (preferred) or the wrapper
        # object emitted by ``--extract --records-json``.  The wrapper
        # shape is ``{"records": [...], ...meta}`` — unwrap it so both
        # CLI entry points behave identically.
        if isinstance(loaded, dict) and "records" in loaded:
            records = loaded["records"]
        else:
            records = loaded

        config = {}
        if parsed.config:
            with open(parsed.config, encoding="utf-8") as f:
                config = json.load(f)

        engine = None
        filtered = None
        if parsed.engine_data:
            # Use the restricted unpickler to prevent arbitrary code
            # execution from crafted pickle files (see C1 fix).
            engine = _safe_pickle_load(parsed.engine_data)
            filtered = getattr(engine, "filtered_records", None)

        success, msg = generate_pdf_from_gui(
            records=records,
            output_path=parsed.output,
            config=config,
            engine=engine,
            filtered=filtered,
        )
        if success:
            sys.stdout.write(msg + "\n")
            sys.exit(0)
        else:
            sys.stderr.write(f"ERROR: {msg}\n")
            sys.exit(1)

    except Exception as e:
        sys.stderr.write(f"ERROR: {e}\n")
        sys.exit(1)


def run_cli_docx_report(args: list[str]) -> None:
    """Run DOCX report generation from command line."""
    import argparse
    import json
    import sys

    from edf_report_docx import generate_docx_from_gui

    parser = argparse.ArgumentParser(
        description="Generate DOCX report from extracted records",
        prog="edf-collector --docx-report",
    )
    parser.add_argument(
        "--records",
        "-i",
        required=True,
        help="Path to extracted records JSON file (exported from GUI or script)",
    )
    parser.add_argument("--output", "-o", required=True, help="Output DOCX file path")
    parser.add_argument("--config", "-c", help="Path to config JSON file (optional)")
    parser.add_argument(
        "--engine-data",
        "-e",
        help="Path to engine data pickle file (optional, for filtered records)",
    )
    parsed = parser.parse_args(args)

    try:
        with open(parsed.records, encoding="utf-8") as f:
            loaded = json.load(f)

        # Accept either a bare list of records (preferred) or the wrapper
        # object emitted by ``--extract --records-json``.  Mirrors the
        # PDF CLI loader so both formats round-trip without extra steps.
        if isinstance(loaded, dict) and "records" in loaded:
            records = loaded["records"]
        else:
            records = loaded

        config = {}
        if parsed.config:
            with open(parsed.config, encoding="utf-8") as f:
                config = json.load(f)

        engine = None
        filtered = None
        if parsed.engine_data:
            # Use the restricted unpickler to prevent arbitrary code
            # execution from crafted pickle files (see C1 fix).
            engine = _safe_pickle_load(parsed.engine_data)
            filtered = getattr(engine, "filtered_records", None)

        success, msg = generate_docx_from_gui(
            records=records,
            output_path=parsed.output,
            config=config,
            engine=engine,
            filtered=filtered,
        )
        if success:
            sys.stdout.write(msg + "\n")
            sys.exit(0)
        else:
            sys.stderr.write(f"ERROR: {msg}\n")
            sys.exit(1)
    except Exception as e:
        sys.stderr.write(f"ERROR: {e}\n")
        sys.exit(1)


def main() -> None:
    """Entry point for the EDF Evidence Collector CLI."""
    import sys

    if len(sys.argv) > 1:
        if sys.argv[1] in ("--pdf-report", "--report", "-r"):
            run_cli_pdf_report(sys.argv[2:])
            return
        elif sys.argv[1] in ("--docx-report", "--word-report", "-w"):
            run_cli_docx_report(sys.argv[2:])
            return
        elif sys.argv[1] in ("--extract", "-e"):
            run_cli_extract(sys.argv[2:])
            return

    if not HAS_TK:
        sys.stderr.write(
            "ERROR: tkinter is not available in this Python build. "
            "Launch a CLI command instead (e.g. --extract, --pdf-report, "
            "--docx-report) or run on a system with Tk installed."
        )
        sys.stderr.write("\n")
        sys.exit(2)

    root = tk.Tk()
    App(root)
    root.mainloop()


def __getattr__(name: str) -> object:
    if name in ("write_evidence_sheet", "write_summary_sheet"):
        from edf_bill_fetcher.io.writers.evidence import (
            write_evidence_sheet,
            write_summary_sheet,
        )

        return {"write_evidence_sheet": write_evidence_sheet, "write_summary_sheet": write_summary_sheet}[name]
    if name == "write_statistical_analysis_sheet":
        from edf_bill_fetcher.io.writers.statistical import write_statistical_analysis_sheet

        return write_statistical_analysis_sheet
    if name == "write_payment_analysis_sheet":
        from edf_bill_fetcher.io.writers.payment import write_payment_analysis_sheet

        return write_payment_analysis_sheet
    if name == "write_forecast_sheet":
        from edf_bill_fetcher.io.writers.forecast import write_forecast_sheet

        return write_forecast_sheet
    if name == "write_data_quality_sheet":
        from edf_bill_fetcher.io.writers.data_quality import write_data_quality_sheet

        return write_data_quality_sheet
    if name == "write_tariff_analysis_sheet":
        from edf_bill_fetcher.io.writers.tariff import write_tariff_analysis_sheet

        return write_tariff_analysis_sheet
    if name == "export_to_excel":
        from edf_bill_fetcher.io.writers.export import export_to_excel

        return export_to_excel
    if name == "write_reconciliation_sheet":
        from edf_bill_fetcher.io.writers.reconciliation import write_reconciliation_sheet

        return write_reconciliation_sheet
    if name in (
        "write_back_billing_sheet",
        "_assess_reason",
        "detect_back_billing",
    ):
        import edf_bill_fetcher.io.writers.back_billing as _m_back_billing

        return getattr(_m_back_billing, name)
    if name in (
        "write_rebilling_sheet",
        "_reversal_match",
        "detect_rebilling",
    ):
        import edf_bill_fetcher.io.writers.rebilling as _m_rebilling

        return getattr(_m_rebilling, name)
    if name in (
        "detect_meter_rollover",
        "infer_contracts",
        "write_meter_readings_sheet",
        "write_contract_history_sheet",
    ):
        import edf_bill_fetcher.io.writers.meter as _m_meter

        return getattr(_m_meter, name)
    if name in (
        "write_sap_back_billing_sheets",
        "_bb_invoice_value",
        "_write_sap_bb_events_sheet",
        "_write_sap_bb_matches_sheet",
        "_write_sap_header_row",
        "_write_sap_contract_history_sheet_impl",
        "write_sap_contract_history_sheet",
        "write_sap_financial_transactions_sheet",
        "write_sap_meter_readings_sheet",
    ):
        import edf_bill_fetcher.io.writers.sap as _m_sap

        return getattr(_m_sap, name)
    if name in (
        "_recon_amount_to_float",
        "_recon_parse_iso_date",
    ):
        import edf_bill_fetcher.io.writers.reconciliation as _m_recon

        return getattr(_m_recon, name)
    if name in (
        "detect_sap_back_billing_events",
        "match_sap_events_to_edf",
        "compute_dispute_flags",
        "build_evidence_index",
        "_SOURCE_PRECEDENCE",
    ):
        import edf_bill_fetcher.writers._helpers as _m_helpers

        return getattr(_m_helpers, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


if __name__ == "__main__":
    main()
