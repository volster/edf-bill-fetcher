"""Tkinter UI classes — App and ReportOptionsDialog."""

from __future__ import annotations

import copy
import gc
import glob
import json
import os
import queue
import tempfile
import threading
import traceback
from collections.abc import Callable
from datetime import date, datetime
from functools import partial
from typing import Any, cast

import pandas as pd

from edf_bill_fetcher.models.config import ConfigDict

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    HAS_TK = True
except ImportError:
    HAS_TK = False

try:
    import pypff

    HAS_PYPFF = pypff is not None
except ImportError:
    HAS_PYPFF = False

try:
    import importlib.util

    HAS_PDF_REPORT = (
        importlib.util.find_spec("edf_bill_fetcher.io.reporters.pdf_report") is not None
    )
    HAS_DOCX_REPORT = (
        importlib.util.find_spec("edf_bill_fetcher.io.reporters.docx_report") is not None
    )
    HAS_HTML_REPORT = (
        importlib.util.find_spec("edf_bill_fetcher.io.reporters.html_report") is not None
    )
except ImportError:
    HAS_PDF_REPORT = False
    HAS_DOCX_REPORT = False
    HAS_HTML_REPORT = False

# Branding constants — kept locally to avoid circular imports through
# edf_collector.py during the modularization refactor.
EDF_ORANGE = "#FE5716"
EDF_NAVY = "#10367A"
EDF_OFFWHITE = "#F5F5F5"
MEDIUM_GREY = "#666666"


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
        ("compensation", "Compensation Analysis", True),
        ("appendix_methodology", "Appendix: Methodology", True),
        ("appendix_glossary", "Appendix: Glossary", True),
        ("appendix_full_evidence", "Appendix: Full Evidence Table", True),
    ]

    def __init__(self, parent):
        """Create the report-options dialog bound to its parent window."""
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
            (
                "both",
                "All Formats (PDF + Word + HTML)",
                "Generate PDF, DOCX, and HTML reports",
            ),
            ("pdf", "PDF Only", "Professional PDF report (reportlab)"),
            ("docx", "Word Document Only", "Editable Word document (python-docx)"),
            (
                "html",
                "HTML Only",
                "Self-contained HTML report (opens in any browser, no external assets)",
            ),
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
    """Tkinter GUI for the EDF Evidence Collector — orchestrates the user-facing workflow."""

    def __init__(self, root):
        """Set up the application window, widgets, and state."""
        self.root = root
        self.root.title("EDF Master Evidence Collector")
        # Default size + minsize chosen so the window fits 768px-tall
        # screens; the scrollable canvas in build_ui handles the overflow.
        self.root.geometry("780x700")
        self.root.minsize(720, 600)
        self.root.configure(bg=EDF_OFFWHITE)

        self.pst_path = tk.StringVar()
        self.pdf_dir = tk.StringVar()
        self.htm_path = tk.StringVar()
        self.eml_msg_dir = tk.StringVar()
        self.acc_num = tk.StringVar(value="")
        self.status = tk.StringVar(value="Ready.")
        self.progress_v = tk.DoubleVar(value=0)

        self.use_anchors = tk.BooleanVar(value=True)
        self.use_large = tk.BooleanVar(value=True)
        self.use_reading_class = tk.BooleanVar(value=True)
        self.use_pdf_fields = tk.BooleanVar(value=True)
        self.use_acc_filt = tk.BooleanVar(value=False)
        self.filter_below = tk.BooleanVar(value=False)
        self.save_filtered = tk.BooleanVar(value=True)
        self.use_dedup = tk.BooleanVar(value=True)
        self.save_dups = tk.BooleanVar(value=True)
        self.use_domain_filter = tk.BooleanVar(value=True)
        self.domain_filter = tk.StringVar(value="edfenergy.com")
        self.min_amount = tk.DoubleVar(value=50.0)
        self.analysis_min = tk.DoubleVar(value=0.0)
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
        self._events: queue.Queue[Callable[[], Any]] = queue.Queue()
        self._worker = None
        self._closing = False
        self.build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._close)
        self._poll_id = self.root.after(50, self._poll_events)
        self.root.bind("<Destroy>", self._on_destroy, add="+")

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
            "use_reading_classification": self.use_reading_class,
            "use_pdf_fields": self.use_pdf_fields,
            "use_acc_filter": self.use_acc_filt,
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

        # Legacy names (pre-typed-config) — apply only when the canonical
        # key is absent, so old config files keep working.
        if "use_reading_classification" not in gui and "use_reading_class" in gui:
            self.use_reading_class.set(bool(gui["use_reading_class"]))
        if "use_acc_filter" not in gui and "use_acc_filt" in gui:
            self.use_acc_filt.set(bool(gui["use_acc_filt"]))

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
            "use_reading_classification": self.use_reading_class.get(),
            "use_pdf_fields": self.use_pdf_fields.get(),
            "use_acc_filter": self.use_acc_filt.get(),
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
            pattern = os.path.join(folder, f"{glob.escape(stem)}_{date_stamp}_*{suffix}.{ext}")
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
        """Construct the full tkinter widget tree for the main window."""
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

        # Wheel scrolling: <MouseWheel> for Windows/macOS, <Button-4>/<Button-5>
        # for X11. Bound on the toplevel bindtag (not bind_all) so every
        # descendant of the main window scrolls, and so ReportOptionsDialog's
        # bind_all/unbind_all lifecycle cannot strip the main window's binding.
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.root.bind("<MouseWheel>", _on_mousewheel)
        self.root.bind("<Button-4>", lambda _e: canvas.yview_scroll(-1, "units"))
        self.root.bind("<Button-5>", lambda _e: canvas.yview_scroll(1, "units"))

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
        browse_row(s1, "EML/MSG Folder:", self.eml_msg_dir, self._pick_eml_msg_dir)

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
            state="normal"
            if (HAS_PDF_REPORT or HAS_DOCX_REPORT or HAS_HTML_REPORT)
            else "disabled",
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

    def _pick_eml_msg_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.eml_msg_dir.set(p)

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
        """Update the status-bar text, marshalling to the tk main thread."""

        def _apply():
            self.status.set(text)
            self.root.update_idletasks()

        if threading.current_thread() is threading.main_thread():
            _apply()
        else:
            self._events.put(_apply)

    def set_progress(self, current, total, text=None):
        """Update the progress bar to current/total percent, marshalling to the tk main thread."""
        pct = max(0, min(100, (current / total) * 100)) if total else 0

        def _apply():
            self.progress_v.set(pct)
            if text:
                self.status.set(text)

        if threading.current_thread() is threading.main_thread():
            _apply()
        else:
            self._events.put(_apply)

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
            self._events.put(_s)

    def _finish(self):
        self._set_extract_idle()
        self.load_report_btn.config(state="normal")
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
        if self._worker is not None and self._worker.is_alive():
            return
        if not HAS_PDF_REPORT and not HAS_DOCX_REPORT and not HAS_HTML_REPORT:
            self._show(
                "error",
                "Report Unavailable",
                "Report generation requires 'reportlab' (PDF) and/or 'python-docx' (Word);\n"
                "HTML reports are always available once this package is installed.\n"
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

            run_dir = tempfile.mkdtemp(prefix=stem + "_report_", dir=base_dir)
            output_paths: dict[str, str] = {}
            if fmt in ("pdf", "both") and HAS_PDF_REPORT:
                output_paths["pdf"] = os.path.join(run_dir, stem + "_Report.pdf")
            if fmt in ("docx", "both") and HAS_DOCX_REPORT:
                output_paths["docx"] = os.path.join(run_dir, stem + "_Report.docx")
            if fmt in ("html", "both") and HAS_HTML_REPORT:
                output_paths["html"] = os.path.join(run_dir, stem + "_Report.html")

            if not output_paths:
                self._show(
                    "warning",
                    "No Reports",
                    "No report paths resolved (check pdf/docx/html availability).",
                )
                return

            self.set_status("Generating report…")
            self.load_report_btn.config(state="disabled")

            config: ConfigDict = {
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
                from edf_bill_fetcher.io.reporters.docx_report import generate_docx_from_gui
                from edf_bill_fetcher.io.reporters.html_report import generate_html_from_gui
                from edf_bill_fetcher.io.reporters.pdf_report import generate_pdf_from_gui

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
                    if "html" in output_paths:
                        s, m = generate_html_from_gui(
                            records=records,
                            output_path=output_paths["html"],
                            config=config,
                            engine=engine,
                            filtered=[],
                        )
                        msgs.append(("HTML", s, m))

                    combined = []
                    all_ok = True
                    for label, ok, m in msgs:
                        if ok:
                            combined.append(
                                f"✓ {label}: {m.split(chr(10))[-1] if m else 'Generated'}\n{output_paths[label.lower()]}"
                            )
                        else:
                            all_ok = False
                            self._events.put(
                                partial(self._show, "error", f"{label} Generation Failed", m)
                            )
                    if all_ok and combined:
                        self._events.put(
                            partial(self._show, "info", "Reports Generated", "\n\n".join(combined))
                        )
                except Exception as e:
                    self._events.put(
                        partial(self._show, "error", "Error", f"An error occurred:\n\n{e}")
                    )
                finally:
                    self._events.put(
                        lambda: (
                            self.load_report_btn.config(state="normal"),
                            self.set_status("Ready."),
                        )
                    )

            self._worker = threading.Thread(target=_generate, daemon=False)
            self._worker.start()

        except Exception as e:
            self._show("error", "Load Error", f"Failed to load spreadsheet:\n\n{e}")

    def _cancel(self):
        self.cancel_event.set()
        self._set_extract_cancelling()
        self.set_status("Cancelling…")

    def start_thread(self):
        """Validate inputs and launch the extraction worker thread."""
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
                self.eml_msg_dir.get().strip(),
            ]
        )
        if not has_sources:
            messagebox.showerror(
                "Error",
                "Please select at least one source:\n"
                "PST/OST file, PDF folder, HTM export, or EML/MSG folder.",
            )
            return
        if self._worker is not None and self._worker.is_alive():
            return
        try:
            self._run_request = self._snapshot_request()
        except Exception as error:
            self._show("error", "Output Error", str(error))
            return
        try:
            self._save_config()
        except OSError as error:
            self._show("warning", "Settings", f"Settings could not be saved: {error}")
        self.cancel_event.clear()
        self.load_report_btn.config(state="disabled")
        self._set_extract_running()
        self.progress_v.set(0)
        self._worker = threading.Thread(target=self._run, daemon=False)
        self._worker.start()

    def _snapshot_request(self):
        """Capture configuration and reserve a directory before starting a worker."""
        config: ConfigDict = {
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

        config["as_of"] = datetime.now().date().isoformat()
        from edf_bill_fetcher.processors.preparation import prepare_evidence

        config = prepare_evidence([], config).config
        paths = {
            name: getattr(self, name).get().strip()
            for name in ("pst_path", "pdf_dir", "htm_path", "eml_msg_dir")
        }
        base_dir = self.output_folder.get().strip()
        if not base_dir:
            source = next((value for value in paths.values() if value), os.getcwd())
            base_dir = source if os.path.isdir(source) else os.path.dirname(source)
        stem = os.path.basename(self.output_name.get().strip() or "EDF_Dispute_Evidence.xlsx")
        stem = os.path.splitext(stem)[0] or "EDF_Dispute_Evidence"
        run_dir = tempfile.mkdtemp(
            prefix=stem + "_" + datetime.now().strftime("%Y%m%d") + "_", dir=base_dir
        )
        from datetime import timezone
        from uuid import uuid4

        from edf_bill_fetcher.domain.json_values import freeze_json
        from edf_bill_fetcher.domain.requests import RunRequest, SourceSelection

        source_pairs = (
            ("pst", paths["pst_path"]),
            ("pdf_dir", paths["pdf_dir"]),
            ("htm", paths["htm_path"]),
            ("eml_dir", paths["eml_msg_dir"]),
            ("msg_dir", paths["eml_msg_dir"]),
        )
        request = RunRequest(
            uuid4().hex,
            datetime.now(timezone.utc),
            date.fromisoformat(config["as_of"]),
            tuple(SourceSelection(kind, path) for kind, path in source_pairs if path),
            freeze_json(dict(config)),
        )
        return {
            "config": copy.deepcopy(config),
            "request": request,
            "paths": paths,
            "stem": stem,
            "run_dir": run_dir,
            "auto_report": self.auto_generate_report.get(),
            "report_options": copy.deepcopy(self._report_options),
        }

    def _poll_events(self):
        """Execute UI callbacks only on the main thread, including during shutdown."""
        if self._poll_id is not None:
            self.root.after_cancel(self._poll_id)
            self._poll_id = None
        try:
            while True:
                callback = self._events.get_nowait()
                if not self._closing:
                    callback()
        except queue.Empty:
            pass
        if self._closing and (self._worker is None or not self._worker.is_alive()):
            self.root.destroy()
            return
        self._poll_id = self.root.after(50, self._poll_events)

    def _on_destroy(self, event):
        if event.widget is self.root and self._poll_id is not None:
            self.root.after_cancel(self._poll_id)
            self._poll_id = None

    def _close(self):
        """Cancel work and keep its cleanup alive before destroying the Tk interpreter."""
        self._closing = True
        self.cancel_event.set()
        self.root.withdraw()

    def _run(self):
        """Translate a captured request and shared coordinator outcome onto the UI queue."""
        from edf_bill_fetcher.application.run_service import (
            bundle_engine_sources,
            collect_engine,
            run,
        )
        from edf_bill_fetcher.collectors.engine import EvidenceEngine
        from edf_bill_fetcher.models.report_view import RenderSelection

        captured = self._run_request
        request = captured["request"]
        config = copy.deepcopy(captured["config"])
        engine = EvidenceEngine(config, self.set_status, self.set_progress, self.cancel_event)
        self.engine = engine
        stem, run_dir = captured["stem"], captured["run_dir"]
        outputs = {"xlsx": os.path.join(run_dir, stem + ".xlsx")}
        options = captured["report_options"]
        sections = tuple(options.get("sections", [s[0] for s in ReportOptionsDialog.SECTIONS]))
        if captured["auto_report"]:
            fmt = options.get("format", "both")
            for name, available in (
                ("pdf", HAS_PDF_REPORT),
                ("docx", HAS_DOCX_REPORT),
                ("html", HAS_HTML_REPORT),
            ):
                if fmt in (name, "both") and available:
                    outputs[name] = os.path.join(run_dir, stem + "_Report." + name)
        try:
            outcome = run(
                request,
                outputs,
                result_path=os.path.join(run_dir, "result.json"),
                manifest_path=os.path.join(run_dir, "run.json"),
                collect=lambda: collect_engine(engine, request),
                copy_sources=lambda prepared: bundle_engine_sources(
                    engine, prepared, run_dir, legacy_paths=(os.path.join(run_dir, "records.json"),)
                ),
                cleanup=engine.cleanup,
                cancelled=self.cancel_event.is_set,
                progress=lambda event: self.set_status(event.message),
                selection=RenderSelection(tuple(outputs), sections),
            )
            self._last_result = outcome.snapshot
            if outcome.status in {"complete", "review_required"}:
                summary = (
                    f"Extraction {outcome.status}.\nPDFs: {engine.pdf_count} | Emails: {engine.email_count} | Records: {len(engine.records)}\n\n"
                    + "\n".join(outputs.values())
                )
                summary += f"\n\nSaved analysis and run status:\n{run_dir}"
                self._show(
                    "warning" if outcome.status == "review_required" else "info",
                    "Review Required" if outcome.status == "review_required" else "Success",
                    summary,
                )
            elif outcome.status == "empty":
                self._show(
                    "warning",
                    "No Data",
                    "No billing records found. Check your sources and filters.",
                )
            elif outcome.status == "cancelled":
                self._show(
                    "warning",
                    "Cancelled",
                    "Run cancelled. Completed artifacts remain in the run folder.",
                )
            else:
                self._show(
                    "error",
                    "Run Failed",
                    f"{outcome.error}\n\nCompleted artifacts and run status:\n{run_dir}",
                )
        except Exception:
            engine.cleanup()
            self._show("error", "Error", traceback.format_exc())
        finally:
            self._events.put(self._finish)

    def _run_auto_report(self, engine, stem, batch_n):
        """Run report generation for the auto-generate flow.

        Uses persisted _report_options; writes to output_folder;
        returns list of written paths.
        """
        from edf_bill_fetcher.io.reporters.docx_report import generate_docx_from_gui
        from edf_bill_fetcher.io.reporters.html_report import generate_html_from_gui
        from edf_bill_fetcher.io.reporters.pdf_report import generate_pdf_from_gui

        ro = self._run_request["report_options"]
        fmt = ro.get("format", "both")
        sections = ro.get("sections", [s[0] for s in ReportOptionsDialog.SECTIONS])

        config = cast(ConfigDict, {**copy.deepcopy(engine.config), "report_sections": sections})

        written: list[str] = []
        if fmt in ("pdf", "both") and HAS_PDF_REPORT:
            pdf_path = os.path.join(self._run_request["run_dir"], stem + "_Report.pdf")
            success, message = generate_pdf_from_gui(
                records=engine.records,
                output_path=pdf_path,
                config=config,
                engine=engine,
                filtered=engine.filtered_records,
            )
            if not success:
                raise RuntimeError(message)
            if success:
                written.append(pdf_path)

        if fmt in ("docx", "both") and HAS_DOCX_REPORT:
            docx_path = os.path.join(self._run_request["run_dir"], stem + "_Report.docx")
            success, message = generate_docx_from_gui(
                records=engine.records,
                output_path=docx_path,
                config=config,
                engine=engine,
                filtered=engine.filtered_records,
            )
            if not success:
                raise RuntimeError(message)
            if success:
                written.append(docx_path)

        if fmt in ("html", "both") and HAS_HTML_REPORT:
            html_path = os.path.join(self._run_request["run_dir"], stem + "_Report.html")
            success, message = generate_html_from_gui(
                records=engine.records,
                output_path=html_path,
                config=config,
                engine=engine,
                filtered=engine.filtered_records,
            )
            if not success:
                raise RuntimeError(message)
            if success:
                written.append(html_path)

        return written
