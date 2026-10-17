"""HTML Report Generator for EDF Energy Ombudsman Submissions.

Builds a self-contained HTML document from the same ``REPORT_SECTIONS``
registry the PDF and DOCX renderers use (see
``edf_bill_fetcher.io.reporters.pdf_report``).  Section titles and
numbering are derived from that registry via ``RenderContext``, so all
three surfaces agree on headings and the table of contents.

The document carries inline CSS only — no external stylesheets, scripts
or remote assets — so it opens correctly from an offline file share.

Each registry section is rendered by a ``create_<name>(...)`` builder
returning an HTML fragment string. Analytical sections present the stored
models as readable tables, with explicit explanations for missing data.
A registry key with no
wired builder raises ``RuntimeError`` at dispatch — the same
loud-failure mode the PDF dispatcher uses.
"""

from __future__ import annotations

import html as html_lib
import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path, PurePosixPath
from typing import Any
from urllib.parse import quote

import pandas as pd

from edf_bill_fetcher.helpers.artifacts import atomic_output
from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.helpers.formatting import fmt_money, fmt_number
from edf_bill_fetcher.helpers.formatting import report_date as fmt_date
from edf_bill_fetcher.helpers.version import get_package_version as _get_package_version
from edf_bill_fetcher.models.config import ConfigDict
from edf_bill_fetcher.models.output_overview import (
    OutputOverview,
    build_output_overview,
    report_account_label,
)
from edf_bill_fetcher.models.output_tables import (
    EVIDENCE_HEADERS,
    EVIDENCE_NOTE,
    evidence_cells,
    timeline_table,
)
from edf_bill_fetcher.models.report_sections import REPORT_SECTIONS, RenderContext
from edf_bill_fetcher.models.report_view import ReportView, capture_report_view
from edf_bill_fetcher.processors.compensation import DISCLAIMER, estimate_compensation

# =============================================================================
# CONSTANTS
# =============================================================================

# Inline stylesheet — deliberately self-contained so the report renders
# offline from a file share without fetching any external asset.
_CSS = """
:root { color-scheme: light; --ink:#18354a; --accent:#007f82; --line:#d6e1e8; }
* { box-sizing:border-box; }
html { scroll-behavior:smooth; scroll-padding-top:24px; }
body { font-family:"Segoe UI",Calibri,Arial,sans-serif; color:#253849; margin:0;
       background:#f2f5f7; font-size:16px; line-height:1.6; }
.container { max-width:1280px; margin:auto; padding:32px; }
.report-layout { display:grid; grid-template-columns:240px minmax(0,1fr); gap:32px; align-items:start; }
.report-layout.no-nav { grid-template-columns:minmax(0,1fr); }
main { min-width:0; background:white; border:1px solid var(--line); border-radius:12px; padding:32px; }
main > div[id^="sec-"] { scroll-margin-top:24px; }
h1 { color:var(--ink); font-size:32px; line-height:1.2; margin:0 0 8px; }
h2.section { color:var(--ink); font-size:23px; line-height:1.3; border-bottom:2px solid var(--line);
             padding-bottom:12px; margin:40px 0 18px; }
main > div:first-child h2 { margin-top:0; }
h3 { color:var(--ink); font-size:18px; line-height:1.4; margin:24px 0 10px; }
p { margin:0 0 16px; overflow-wrap:anywhere; }
p.note { background:#edf6f7; border-left:4px solid var(--accent); padding:14px 18px; border-radius:0 6px 6px 0; }
p.muted { color:#526576; font-size:14px; }
.table-scroll { width:100%; overflow-x:auto; margin:0 0 20px; border:1px solid var(--line); border-radius:6px; }
table { border-collapse:collapse; width:100%; }
th,td { padding:10px 12px; font-size:14px; text-align:left; vertical-align:top;
        border-bottom:1px solid var(--line); overflow-wrap:anywhere; }
thead th { background:var(--ink); color:white; font-weight:600; }
tbody th { font-weight:600; background:transparent; }
tr.alt td,tr.alt th { background:#f4f7fa; }
.numeric { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
.cover { background:white; border:1px solid var(--line); border-top:6px solid var(--accent);
         border-radius:12px; padding:32px; margin-bottom:28px; }
.cover .subtitle { color:var(--accent); font-size:18px; margin:10px 0 24px; }
.cover .confidential { color:#8c3532; font-weight:600; font-size:13px; margin:24px 0 8px; }
.cover .table-scroll { max-width:720px; }
.toc { position:sticky; top:24px; max-height:calc(100vh - 48px); overflow-y:auto; font-size:14px; }
.toc h2.section { margin:0 0 14px; font-size:18px; }
.toc ol { list-style:none; padding:0; margin:0; }
.toc li { margin:0 0 6px; }
.toc a { display:block; padding:7px 10px; border-radius:5px; text-decoration:none; color:var(--ink); }
.toc a:hover,.toc a:focus-visible { background:#dfecef; text-decoration:underline; }
a { color:#006b80; text-underline-offset:3px; }
:focus-visible { outline:3px solid var(--accent); outline-offset:3px; }
.skip-link { position:absolute; left:16px; top:-100px; background:white; padding:12px; z-index:3; }
.skip-link:focus { top:16px; }
.priorities { padding-left:24px; }
.priorities li { margin-bottom:14px; }
details.audit { margin:12px 0 24px; }
details.audit > summary { cursor:pointer; color:var(--ink); font-weight:600; padding:12px;
                         background:#edf3f6; border:1px solid var(--line); border-radius:6px; }
.failed { color:#a12727; }
ul { padding-left:24px; }
@media (max-width:850px) {
 .container { padding:16px; } .report-layout { display:block; }
 .toc { position:static; max-height:none; margin:0 0 24px; }
 .toc ol { columns:2; } .toc li { break-inside:avoid; }
 main,.cover { padding:22px; } h1 { font-size:26px; }
}
@media (max-width:480px) { .toc ol { columns:1; } main,.cover { padding:16px; } }
@media (prefers-reduced-motion:reduce) { html { scroll-behavior:auto; } }
@media print {
 @page { size:A4; margin:18mm; }
 body { background:white; font-size:11pt; } .container { max-width:none; padding:0; }
 .toc,.skip-link { display:none; } .report-layout { display:block; }
 main,.cover { padding:0; border:0; border-radius:0; }
 .cover { margin-bottom:24pt; } h2.section,h3 { break-after:avoid; }
 .table-scroll { overflow:visible; border:0; } th,td { font-size:9pt; padding:5pt; }
 thead { display:table-header-group; } thead th { color:black; background:#e7edf0; }
 tr { break-inside:avoid; } p { orphans:3; widows:3; }
 a { color:inherit; } p.note { border:1px solid #a9bcc6; }
 details::details-content { content-visibility:visible; display:block; }
}
"""


def _esc(value: object) -> str:
    """HTML-escape a value so user-derived text renders as visible text."""
    return html_lib.escape(str(value), quote=True)


@dataclass(frozen=True)
class _SourceLink:
    """A local manifest path, never an arbitrary HTML string from evidence."""

    href: str
    label: str


def _source_links(bundle_dir: str | None) -> dict[str, _SourceLink]:
    """Resolve only saved, contained manifest paths beside this report."""
    if not bundle_dir:
        return {}
    root = Path(bundle_dir).resolve()
    try:
        files = json.loads((root / "evidence_manifest.json").read_text(encoding="utf-8"))["files"]
    except (OSError, ValueError, KeyError, TypeError):
        return {}
    links: dict[str, _SourceLink] = {}
    ambiguous: set[str] = set()
    if not isinstance(files, list):
        return links
    for entry in files:
        if not isinstance(entry, dict) or entry.get("status") != "saved":
            continue
        path = entry.get("saved_path")
        if not isinstance(path, str) or "\\" in path or ":" in path:
            continue
        relative = PurePosixPath(path)
        if relative.is_absolute() or ".." in relative.parts:
            continue
        target = (root / path).resolve()
        if root not in target.parents or not target.is_file():
            continue
        link = _SourceLink(quote(path, safe="/"), "Open source")
        for key in (entry.get("source_id"), entry.get("attachment_name")):
            if isinstance(key, str) and key:
                if key in links and links[key] != link:
                    ambiguous.add(key)
                links[key] = link
    return {key: link for key, link in links.items() if key not in ambiguous}


def _table(rows: list[list[Any]], *, header: bool = True) -> str:
    """Render a list-of-rows as a styled HTML table.

    ``header=True`` treats ``rows[0]`` as a header row (navy background).
    Every cell is HTML-escaped; even-indexed body rows get the
    alternating ``tr.alt`` tint.
    """
    if not rows:
        return ""
    parts = [
        '<div class="table-scroll" tabindex="0" role="region" aria-label="Report data table"><table>'
    ]
    body_start = 0
    if header:
        head = rows[0]
        parts.append(
            "<thead><tr>"
            + "".join(f'<th scope="col">{_esc(c)}</th>' for c in head)
            + "</tr></thead>"
        )
        body_start = 1
    parts.append("<tbody>")
    for i, row in enumerate(rows[body_start:], start=body_start):
        cls = ' class="alt"' if i % 2 == 0 else ""
        cells = "".join(
            f'<td><a href="{_esc(c.href)}">{_esc(c.label)}</a></td>'
            if isinstance(c, _SourceLink)
            else f'<td class="numeric">{_esc(c)}</td>'
            if isinstance(c, (int, float)) or str(c).startswith("£")
            else f"<td>{_esc(c)}</td>"
            for c in row
        )
        parts.append(f"<tr{cls}>{cells}</tr>")
    parts.append("</tbody></table></div>")
    return "\n".join(parts)


def _note(text: str) -> str:
    """Render an italicised note paragraph (placeholder / info box)."""
    return f'<p class="note">{_esc(text)}</p>'


def _not_implemented_in_html(title: str) -> str:
    """Render a "not implemented in HTML" placeholder for a section.

    The HTML surface intentionally renders a note instead of the
    chart/analysis the PDF and DOCX surfaces carry, so the wording
    "not implemented in HTML" is part of the contract pinned by
    ``tests/test_html_report.py``.
    """
    return _note(
        f"{title} is not implemented in HTML. "
        "Please refer to the PDF or DOCX report (or the Excel evidence workbook) "
        "for this analysis."
    )


# =============================================================================
# SECTION CREATORS
# =============================================================================


def create_cover_page(acc_ref: str, period_start: str, period_end: str, report_date: str) -> str:
    """Create the cover header (version carried from pyproject.toml)."""
    version = _get_package_version()
    rows = [
        ["Account Reference", acc_ref],
        ["Record dates", f"{period_start} to {period_end}"],
        ["Report Generated", report_date],
        ["Prepared by", "EDF Evidence Collector"],
    ]
    disclaimer = (
        "This tool was created for personal use in an EDF billing dispute. "
        "It is provided as-is without warranty. "
        "Always verify extracted data against original documents "
        "before using in any formal dispute."
    )
    return (
        '<div class="cover">'
        "<h1>EDF Energy Billing Dispute</h1>"
        "<h1>Ombudsman Evidence Report</h1>"
        '<p class="subtitle">Prepared for Energy Ombudsman Review</p>'
        f"{_table(rows, header=False)}"
        '<p class="confidential">CONFIDENTIAL — FOR OMBUDSMAN REVIEW ONLY</p>'
        f'<p class="muted">Generated by EDF Evidence Collector v{_esc(version)}<br>'
        "All data extracted from original source documents "
        "(EDF bills, HTM exports, email archives).<br>"
        "Methodology detailed in the Methodology &amp; Data Sources appendix.</p>"
        f'<p class="muted">{_esc(disclaimer)}</p>'
        "</div>"
    )


def create_table_of_contents(ctx: RenderContext) -> str:
    """Create the table of contents, driven by ``ctx``."""
    parts = [
        '<nav class="toc" aria-label="Report contents"><h2 class="section">Table of Contents</h2>'
    ]
    if not ctx.sections_in_order:
        parts.append("<p><i>No report sections selected.</i></p>")
        parts.append("</nav>")
        return "\n".join(parts)
    parts.append("<ol>")
    for spec in ctx.sections_in_order:
        anchor = f"sec-{spec.section.key}"
        parts.append(
            f'<li><a href="#{anchor}">{_esc(spec.label)} {_esc(spec.section.title)}</a></li>'
        )
    parts.append("</ol></nav>")
    return "\n".join(parts)


def create_executive_summary(
    df: pd.DataFrame,
    config: dict,
    account_ref: str,
    flag_count: dict,
    total_records: int,
    total_charges: float,
    total_payments: float,
    period_start: str,
    period_end: str,
    opening_balance: float | None = None,
    closing_balance: float | None = None,
    ctx: RenderContext | None = None,
    overview: OutputOverview | None = None,
) -> str:
    """Create the executive summary section."""
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("exec_summary")
    net_change = total_charges - total_payments

    parts = [f'<div id="sec-exec_summary"><h2 class="section">{_esc(heading)}</h2>']
    parts.append(
        "<p>This report reviews the supplied EDF Energy billing evidence "
        f"for account <b>{_esc(account_ref)}</b>, with record dates from "
        f"<b>{_esc(period_start)}</b> to <b>{_esc(period_end)}</b>. "
        f"The analysis encompasses <b>{total_records}</b> evidence record{'s' if total_records != 1 else ''} sourced from "
        "the sources listed in the evidence index.</p>"
    )
    fin_rows = [
        ["Metric", "Amount"],
        ["Total Charges (Debits)", fmt_money(total_charges)],
        ["Net Payments/Credits (After Refunds)", fmt_money(total_payments)],
        ["Recorded Net Movement", fmt_money(net_change)],
        [
            "First observed balance",
            fmt_money(opening_balance) if opening_balance is not None else "—",
        ],
        [
            "Last observed balance",
            fmt_money(closing_balance) if closing_balance is not None else "—",
        ],
    ]
    parts.append("<h3>Financial Summary</h3>")
    parts.append(_table(fin_rows))
    if overview:
        parts.append(_note(f"Analysis as of {overview.analysis_date}. {overview.balance_note}"))
        parts.append('<h3>Review priorities</h3><ol class="priorities">')
        for priority in overview.priorities:
            title = _esc(priority.title)
            if priority.section in ctx.selected:
                title = f'<a href="#sec-{priority.section}">{title}</a>'
            parts.append(f"<li><strong>{title}</strong><br>{_esc(priority.action)}</li>")
        parts.append("</ol>")

    findings: list[str] = []
    if flag_count.get("HIGH", 0) > 0:
        findings.append(
            f"{flag_count['HIGH']} HIGH-severity findings require review. "
            "Read the dated findings for the specific checks triggered and their supporting evidence."
        )
    if flag_count.get("MEDIUM", 0) > 0:
        findings.append(
            f"{flag_count['MEDIUM']} MEDIUM-severity findings require review. "
            "The findings table explains each observation and its account context."
        )
    if flag_count.get("INFO", 0) > 0:
        findings.append(
            f"{flag_count['INFO']} informational findings provide context for manual review."
        )
    if not findings:
        findings.append(
            "No automated flags were generated. This does not establish that all bills and payments are present or correct."
        )

    parts.append("<h3>Key Findings</h3><ul>")
    for finding in findings:
        parts.append(f"<li>{_esc(finding)}</li>")
    parts.append("</ul>")

    parts.append(
        "<h3>Conclusion</h3>"
        "<p>This report summarizes the supplied evidence and highlights observations for review. "
        "Balance changes and missing records alone do not establish overcharging or a regulatory breach. "
        "Check the original documents, account context and applicable rules before making a claim.</p>"
    )
    parts.append("</div>")
    return "\n".join(parts)


def create_key_findings_table(flags: list, ctx: RenderContext | None = None) -> str:
    """Create the key findings summary section."""
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("key_findings")
    parts = [f'<div id="sec-key_findings"><h2 class="section">{_esc(heading)}</h2>']

    if not flags:
        parts.append(
            "<p>No automated flags were generated. Manual review of the original evidence "
            "and its completeness is still required.</p></div>"
        )
        return "\n".join(parts)

    high = [f for f in flags if f[4] == "HIGH"]
    medium = [f for f in flags if f[4] == "MEDIUM"]
    info = [f for f in flags if f[4] == "INFO"]

    summary_rows = [
        ["Severity", "Count", "Description"],
        ["HIGH", str(len(high)), "Review original evidence and account context"],
        ["MEDIUM", str(len(medium)), "Significant deviation — investigation warranted"],
        ["INFO", str(len(info)), "Informational — payments/credits noted"],
        ["TOTAL", str(len(flags)), "All automated findings"],
    ]
    parts.append(_table(summary_rows))

    for severity, group in (("HIGH", high), ("MEDIUM", medium)):
        if not group:
            continue
        parts.append(f"<h3>{severity} Severity Findings</h3><ul>")
        for i, (ftype, date, amt, detail, _sev) in enumerate(group, 1):
            date_str = fmt_date(date)
            amt_str = fmt_money(amt) if amt else ""
            parts.append(
                f"<li><b>{i}. {_esc(ftype)}</b> ({_esc(date_str)}, {_esc(amt_str)}) "
                f"— {_esc(detail)}</li>"
            )
        parts.append("</ul>")

    parts.append("</div>")
    return "\n".join(parts)


def create_evidence_index(
    df: pd.DataFrame,
    engine: Any,
    ctx: RenderContext | None = None,
    *,
    bundle_dir: str | None = None,
) -> str:
    """Create the evidence index with source cross-references."""
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("evidence_index")
    parts = [f'<div id="sec-evidence_index"><h2 class="section">{_esc(heading)}</h2>']

    source_counts = df["Source"].value_counts()
    total = len(df)
    source_rows = [["Source", "Records", "Percentage"]]
    for src, cnt in source_counts.items():
        source_rows.append([src, str(cnt), f"{cnt / total:.1%}"])
    source_rows.append(["TOTAL", str(total), "100.0%"])
    parts.append(_table(source_rows))

    if engine is not None and hasattr(engine, "pdf_count"):
        parts.append(
            f'<p class="muted">PST/OST emails scanned: {getattr(engine, "email_count", 0)}'
            f" · PDF attachments extracted: {getattr(engine, 'pdf_count', 0)}</p>"
        )

    parts.append("<h3>Source Document Inventory</h3>")
    links = _source_links(bundle_dir)
    for src in source_counts.index:
        src_df = df[df["Source"] == src].copy()
        src_df["_dt"] = src_df["Date"].apply(parse_to_sort_date)
        src_df = src_df.sort_values("_dt")

        detail_rows: list[list[Any]] = [
            ["Date", "Invoice #", "Amount", "Period", "Entry Type", "Reading", "Source"]
        ]
        for _, row in src_df.iterrows():
            source_id = row.get("Source ID")
            link = links.get(source_id) if source_id else links.get(row.get("Attachment Name"))
            if (
                link
                and isinstance(row.get("Page From"), (int, float))
                and pd.notna(row["Page From"])
            ):
                link = _SourceLink(f"{link.href}#page={int(row['Page From'])}", link.label)
            detail_rows.append(
                [
                    fmt_date(row.get("Date")),
                    row.get("Invoice #", "N/A"),
                    fmt_money(row.get("Amount (£)")),
                    f"{fmt_date(row.get('Period From'))}–{fmt_date(row.get('Period To'))}",
                    row.get("Entry Type", ""),
                    row.get("Reading", ""),
                    link or "Unavailable",
                ]
            )
        parts.append(f"<h3>{_esc(src)} ({len(src_df)} records)</h3>")
        parts.append(
            f'<details class="audit"><summary>{_esc(src)}: show {len(src_df)} source observations</summary>'
        )
        parts.append(_table(detail_rows))
        parts.append("</details>")

    parts.append("</div>")
    return "\n".join(parts)


def create_anomaly_detail_section(
    flags: list, df: pd.DataFrame, ctx: RenderContext | None = None
) -> str:
    """Create the detailed findings section."""
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("detailed_findings")
    parts = [f'<div id="sec-detailed_findings"><h2 class="section">{_esc(heading)}</h2>']

    if not flags:
        parts.append(
            "<p>No specific anomalies were automatically detected. The timeline and "
            "statistical analysis sections may reveal patterns warranting manual "
            "investigation.</p></div>"
        )
        return "\n".join(parts)

    categories: dict[str, list[tuple]] = {
        "LARGE JUMP": [],
        "BILLING GAP": [],
        "ESTIMATED RUN": [],
        "HIGH DAILY RATE": [],
        "RECONCILIATION MISMATCH": [],
        "BALANCE REDUCTION": [],
    }
    for f in flags:
        if f[0] in categories:
            categories[f[0]].append(f)

    for cat, cat_flags in categories.items():
        if not cat_flags:
            continue
        parts.append(f"<h3>{_esc(cat.replace('_', ' ').title())}</h3>")
        detail_rows = [["#", "Date", "Amount", "Severity", "Detail"]]
        for i, (_ftype, date, amt, detail, sev) in enumerate(cat_flags, 1):
            truncated = detail[:200] + ("..." if len(detail) > 200 else "")
            detail_rows.append(
                [
                    str(i),
                    fmt_date(date),
                    fmt_money(amt) if amt else "",
                    sev,
                    truncated,
                ]
            )
        parts.append(_table(detail_rows))

    parts.append("</div>")
    return "\n".join(parts)


def create_revision_section(model: dict | None) -> str:
    """Append stored revision comparisons to the detailed findings output."""
    if model is None:
        return ""
    from edf_bill_fetcher.models.output_tables import revision_note, revision_table

    return (
        "<h3>Bill revision review</h3>"
        + f'<p class="note">{_esc(revision_note(model))}</p>'
        + (
            _table(revision_table(model))
            if model["relationships"]
            else "<p>No revision relationships were found. You can add a missed relationship in the app.</p>"
        )
    )


def create_timeline_section(df: pd.DataFrame, flags: list, ctx: RenderContext | None = None) -> str:
    """Create the chronological timeline of events."""
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("timeline")
    parts = [f'<div id="sec-timeline"><h2 class="section">{_esc(heading)}</h2>']

    rows = timeline_table(df, flags)
    parts.append(
        f'<details class="audit"><summary>Show the full timeline ({max(0, len(rows) - 1)} events)</summary>'
    )
    parts.append(_table(rows))
    parts.append("</details>")
    parts.append("</div>")
    return "\n".join(parts)


def create_ofgem_comparison(
    df: pd.DataFrame,
    config: dict | None = None,
    ctx: RenderContext | None = None,
    *,
    payload: Any = None,
) -> str:
    """Create the OFGEM price cap comparison section."""
    from edf_bill_fetcher.models.report_models import compute_ofgem_comparison

    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("ofgem")
    parts = [f'<div id="sec-ofgem"><h2 class="section">{_esc(heading)}</h2>']

    parts.append(
        "<p>The following analysis compares the effective unit rates charged on EDF "
        "bills against the OFGEM Default Tariff Cap (Price Cap) for the "
        "corresponding consumption periods. This GB-average direct-debit electricity "
        "benchmark requires known variable/default tariff context; an exceedance is "
        "a review signal, rather than a determination of the account’s legal cap.</p>"
    )

    comparison = payload if payload is not None else compute_ofgem_comparison(df, config)

    if not comparison.rows:
        parts.append(
            "<p>No billing records with a usable extracted energy unit rate and consumption period "
            "are available for comparison. A total charge divided by usage is not a substitute.</p></div>"
        )
        return "\n".join(parts)

    MISSING = "—"
    cap_rows: list[list[Any]] = []
    for row in comparison.rows:
        cap_rows.append(
            [
                row.quarter,
                fmt_number(row.bill_rate, 2),
                fmt_number(row.cap_rate, 2) if row.cap_rate is not None else MISSING,
                fmt_number(row.diff, 2)
                if row.diff is not None and row.diff != 0
                else ("0.00" if row.diff is not None else MISSING),
                row.status,
            ]
        )

    cap_rows.append(["OVERALL", "—", "—", comparison.overall_diff, comparison.overall_verdict])

    header = ["Period", "Bill Unit Rate (p/kWh)", "OFGEM Cap (p/kWh)", "Difference", "Status"]
    parts.append(_table([header, *cap_rows]))
    parts.append(
        '<p class="muted"><b>Methodology:</b> Extracted energy unit rates are allocated '
        "to consumption quarters and weighted by usage, or days when usage is unavailable. "
        "The model does not divide total bill charges by kWh or assess standing charges here. "
        "Missing tariff/payment/fuel context remains unassessed. OFGEM data is a "
        "GB-average direct-debit electricity benchmark, not an account-specific legal cap.</p>"
    )
    parts.append("</div>")
    return "\n".join(parts)


def create_statistical_analysis(
    df: pd.DataFrame, ctx: RenderContext | None = None, *, payload: Any = None
) -> str:
    """Render the shared statistical model."""
    if ctx is None:
        ctx = RenderContext()
    if payload is None:
        from edf_bill_fetcher.models.report_models import compute_statistical_analysis

        payload = compute_statistical_analysis(df)
    heading = (
        f'<div id="sec-statistical"><h2 class="section">{_esc(ctx.heading("statistical"))}</h2>'
    )
    if payload.count < 3:
        return (
            heading
            + _note(
                "At least three usable balance observations in one ledger are needed for descriptive statistics. No statistical conclusion is available for this selection."
            )
            + "</div>"
        )
    rows = [
        ["Balance statistic", "Value"],
        ["Observations", payload.count],
        ["Mean observed balance", fmt_money(payload.mean)],
        ["Median observed balance", fmt_money(payload.median)],
        ["Lowest observed balance", fmt_money(payload.minimum)],
        ["Highest observed balance", fmt_money(payload.maximum)],
        ["Standard deviation", fmt_money(payload.std)],
        ["Range", fmt_money(payload.range)],
    ]
    anomaly_rows = [
        ["Screening method", "Flagged observations", "Dates"],
        ["Z-score", payload.z_count, ", ".join(payload.z_dates) or "None"],
        ["Interquartile range", payload.iqr_count, ", ".join(payload.iqr_dates) or "None"],
    ]
    return (
        heading
        + _note(
            "These statistics describe recorded balances, not the cost of individual bills. Outlier screens identify observations to inspect; they do not establish a billing error."
        )
        + _table(rows)
        + "<h3>Observations to inspect</h3>"
        + _table(anomaly_rows)
        + "</div>"
    )


def create_payment_analysis(
    df: pd.DataFrame, ctx: RenderContext | None = None, *, payload: Any = None
) -> str:
    """Render the shared payment model."""
    if ctx is None:
        ctx = RenderContext()
    if payload is None:
        from edf_bill_fetcher.models.report_models import compute_payment_analysis

        payload = compute_payment_analysis(df)
    heading = f'<div id="sec-payment"><h2 class="section">{_esc(ctx.heading("payment"))}</h2>'
    if not payload.count:
        return (
            heading
            + _note(
                "No payment/credit records were supplied. This is not evidence that no payments were made; check statements or payment confirmations."
            )
            + "</div>"
        )
    rows = [
        ["Payment / credit summary", "Value"],
        ["Records supplied", payload.count],
        ["Recorded total", fmt_money(payload.total_paid)],
        ["Average", fmt_money(payload.avg_payment)],
        ["Median", fmt_money(payload.median_payment)],
        ["Largest", fmt_money(payload.largest_payment)],
        ["Smallest", fmt_money(payload.smallest_payment)],
        ["Last dated payment", fmt_date(payload.last_payment_date)],
        ["Last dated amount", fmt_money(payload.last_payment_amount)],
    ]
    if payload.avg_interval_days is not None:
        rows.append(["Average interval within a ledger", f"{payload.avg_interval_days:.1f} days"])
    chronology = [["Date", "Type", "Recorded payment / credit", "Balance after", "Details"]]
    for row in payload.chronology.to_dict("records"):
        chronology.append(
            [
                fmt_date(row.get("Date")),
                row.get("Entry Type"),
                fmt_money(row.get("_amount")),
                fmt_money(row.get("Balance (£)")),
                row.get("Details", ""),
            ]
        )
    return heading + _table(rows) + "<h3>Payment chronology</h3>" + _table(chronology) + "</div>"


def create_forecast_section(
    df: pd.DataFrame, ctx: RenderContext | None = None, *, payload: Any = None
) -> str:
    """Render the shared forecast model."""
    if ctx is None:
        ctx = RenderContext()
    if payload is None:
        from edf_bill_fetcher.models.report_models import compute_forecast

        payload = compute_forecast(df)
    heading = f'<div id="sec-forecast"><h2 class="section">{_esc(ctx.heading("forecast"))}</h2>'
    if not payload.linear_forecast and not payload.ema_forecast and not payload.hw_forecast:
        return (
            heading
            + _note(
                "No forecast is available. At least three usable monthly balance observations in a single ledger are needed. A forecast is context, not evidence of future charges or compensation."
            )
            + "</div>"
        )
    dates = [
        parse_to_sort_date(value)
        for value in df.get("_dt", df.get("Date", pd.Series(dtype=object)))
    ]
    dated = [value for value in dates if not pd.isna(value)]
    rows = [["Forecast month", "Linear balance", "EMA balance", "Holt–Winters balance"]]
    projections = (payload.linear_forecast, payload.ema_forecast, payload.hw_forecast or [])
    for index in range(max(map(len, projections))):
        label = (
            (max(dated) + pd.DateOffset(months=index + 1)).strftime("%b %Y")
            if dated
            else f"Month {index + 1}"
        )
        rows.append(
            [
                label,
                *(
                    fmt_money(values[index]) if index < len(values) else "Unavailable"
                    for values in projections
                ),
            ]
        )
    information = "".join(f"<li>{_esc(text)}</li>" for text in payload.model_info)
    return (
        heading
        + _note(
            f"Projections use {payload.n} monthly balance observations. They estimate a balance trajectory, not future bill charges. Historical fit does not guarantee future accuracy."
        )
        + _table(rows)
        + f"<details><summary>Model notes</summary><ul>{information}</ul></details></div>"
    )


def create_data_quality_section(
    df: pd.DataFrame, ctx: RenderContext | None = None, *, payload: Any = None
) -> str:
    """Create the data quality assessment section."""
    from edf_bill_fetcher.models.report_models import compute_data_quality_report

    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("data_quality")
    parts = [f'<div id="sec-data_quality"><h2 class="section">{_esc(heading)}</h2>']

    dq = payload if payload is not None else compute_data_quality_report(df)
    total = dq.total_records
    date_parsed = dq.date_parsed
    amt_complete = dq.amt_complete
    period_complete = dq.period_complete
    reading_classified = dq.reading_classified
    dup_count = dq.duplicate_count

    quality_rows = [["Check", "Passed", "Total", "Rate", "Status"]]
    quality_rows.append(
        [
            "Date Parsing",
            str(int(date_parsed)),
            str(total),
            f"{date_parsed / total:.1%}",
            "PASS"
            if date_parsed / total > 0.9
            else "WARN"
            if date_parsed / total > 0.7
            else "FAIL",
        ]
    )
    quality_rows.append(
        [
            "Amount Complete",
            str(int(amt_complete)),
            str(total),
            f"{amt_complete / total:.1%}",
            "PASS" if amt_complete == total else "WARN",
        ]
    )
    quality_rows.append(
        [
            "Period Info Complete",
            str(int(period_complete)),
            str(total),
            f"{period_complete / total:.1%}",
            "PASS"
            if period_complete / total > 0.7
            else "WARN"
            if period_complete / total > 0.5
            else "FAIL",
        ]
    )
    quality_rows.append(
        [
            "Reading Classified",
            str(int(reading_classified)),
            str(total),
            f"{reading_classified / total:.1%}",
            "PASS" if reading_classified / total > 0.5 else "WARN",
        ]
    )
    quality_rows.append(
        [
            "Duplicates (Date+Amount)",
            str(int(dup_count)),
            str(total),
            f"{dup_count / total:.2%}",
            "PASS" if dup_count / total < 0.05 else "WARN" if dup_count / total < 0.15 else "FAIL",
        ]
    )
    parts.append(_table(quality_rows))

    parts.append("<h3>Source Distribution</h3>")
    src_rows = [["Source", "Records", "Percentage"]]
    for src, cnt in df["Source"].value_counts().items():
        src_rows.append([src, str(cnt), f"{cnt / total:.1%}"])
    src_rows.append(["TOTAL", str(total), "100.0%"])
    parts.append(_table(src_rows))
    parts.append("</div>")
    return "\n".join(parts)


def create_tariff_impact_section(
    df: pd.DataFrame, ctx: RenderContext | None = None, *, payload: Any = None
) -> str:
    """Create the tariff impact analysis section."""
    from edf_bill_fetcher.models.report_models import compute_tariff_analysis

    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("tariff")
    parts = [f'<div id="sec-tariff"><h2 class="section">{_esc(heading)}</h2>']

    if "Tariff" not in df.columns or df["Tariff"].isna().all() or (df["Tariff"] == "N/A").all():
        parts.append(
            "<p>No tariff data available in the extracted records. Tariff information "
            "is typically found on new-format (KI/KCR) invoices.</p></div>"
        )
        return "\n".join(parts)

    analysis = payload if payload is not None else compute_tariff_analysis(df)
    if analysis.empty:
        parts.append("<p>No valid tariff records found.</p></div>")
        return "\n".join(parts)

    tariff_stats = analysis.stats

    parts.append("<h3>Unit Rate by Tariff</h3>")
    tariff_rows = [["Tariff", "Records", "Avg Rate (p/kWh)", "Median", "Min", "Max"]]
    for _, row in tariff_stats.iterrows():
        tariff_rows.append(
            [
                row["Tariff"],
                str(int(row["count"])),
                fmt_number(row["avg_unit_rate"], 2),
                fmt_number(row["median_unit_rate"], 2),
                fmt_number(row["min_unit_rate"], 2),
                fmt_number(row["max_unit_rate"], 2),
            ]
        )
    parts.append(_table(tariff_rows))

    parts.append(f"<h3>Tariff Segments Observed: {analysis.tariff_changes}</h3>")
    parts.append("</div>")
    return "\n".join(parts)


def create_compensation_section(
    df: pd.DataFrame,
    config: dict[str, Any] | None = None,
    ctx: RenderContext | None = None,
    *,
    payload: Any = None,
) -> str:
    """Create the Compensation Analysis section."""
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("compensation")
    parts = [f'<div id="sec-compensation"><h2 class="section">{_esc(heading)}</h2>']

    rows = payload if payload is not None else estimate_compensation(df, config)

    if not rows:
        parts.append("<p>No compensation claims identified in the extracted records.</p></div>")
        return "\n".join(parts)

    categories: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        cat = str(row.get("category", ""))
        categories.setdefault(cat, []).append(row)

    for cat, cat_rows in categories.items():
        parts.append(f"<h3>{_esc(cat)}</h3>")
        table_rows: list[list[Any]] = [
            [
                "Category",
                "Invoice #",
                "Date",
                "Base Amount",
                "Days",
                "Rate",
                "Indicative Amount",
                "Legal Basis",
            ]
        ]
        for r in cat_rows:
            table_rows.append(
                [
                    str(r.get("category", "")),
                    str(r.get("invoice_ref", "")),
                    str(r.get("date", "")),
                    fmt_money(r.get("base_amount")),
                    str(r.get("days", "")),
                    fmt_number(r.get("rate"), 4) if r.get("rate") is not None else "N/A",
                    fmt_money(r.get("indicative_amount")),
                    str(r.get("legal_basis", "")),
                ]
            )
        parts.append(_table(table_rows))

    total = sum(float(r.get("indicative_amount") or 0.0) for r in rows)
    parts.append(f"<h3>Total Indicative Compensation: {_esc(fmt_money(total))}</h3>")
    parts.append(f'<p class="note">{_esc(DISCLAIMER)}</p>')
    parts.append("</div>")
    return "\n".join(parts)


def create_appendix_methodology(config: ConfigDict, ctx: RenderContext | None = None) -> str:
    """Create the methodology appendix."""
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("appendix_methodology")
    parts = [f'<div id="sec-appendix_methodology"><h2 class="section">{_esc(heading)}</h2>']

    sections: list[tuple[str, list[str]]] = [
        (
            "A.1 Data Sources",
            [
                "Supported sources include the following. See the evidence index for the sources supplied in this run:",
                "• PDF Bills: EDF Energy invoices (both legacy and new KI/KCR formats) "
                "processed using text extraction or OCR, then format-specific field extraction.",
                "• HTM Export: EDF MyAccount 'Payments and Invoices' export parsed via "
                "BeautifulSoup with pattern matching for charge/payment/reversal entries.",
                "• PST/OST Email Archives: Outlook data files processed via libpff-python, "
                "extracting email bodies (HTML/plain text/RTF) and PDF attachments.",
            ],
        ),
        (
            "A.2 Amount Extraction Logic",
            [
                "Extraction strategies depend on the selected settings and available document text:",
                "1. Smart Context Search: prioritized patterns targeting specific "
                "EDF billing language. Patterns execute in priority order; first match wins.",
                "2. Large Amount Fallback: scans all £ amounts ≥ minimum threshold, "
                "selecting the largest. Used when context patterns fail.",
            ],
        ),
        (
            "A.3 Deduplication",
            [
                "Duplicates require compatible account, fuel, meter and register identities.",
                "Bills merge only with supporting invoice/period identity and consistent amounts; repeated payments remain separate unless transaction identity supports a duplicate.",
                "Original observations, duplicate decisions and source identities are preserved in the accompanying saved result.",
            ],
        ),
        (
            "A.4 Configuration Used",
            [
                f"Valuation date: {config.get('as_of', 'Unavailable')}",
                f"Minimum Amount Threshold: {fmt_money(config.get('min_amount', 500))}",
                f"Analysis Threshold: {fmt_money(config.get('analysis_min', 0))}",
                "Account Filter: "
                f"{'Enabled' if config.get('use_acc_filter') else 'Disabled'} "
                f"({config.get('acc_num', 'N/A')})",
                "Domain Filter: "
                f"{'Enabled' if config.get('use_domain_filter') else 'Disabled'} "
                f"({config.get('domain_filter', 'edfenergy.com')})",
                f"Deduplication: {'Enabled' if config.get('use_dedup') else 'Disabled'}",
                f"Smart Context Search: {'Enabled' if config.get('use_anchors') else 'Disabled'}",
                f"Large Amount Fallback: {'Enabled' if config.get('use_large') else 'Disabled'}",
            ],
        ),
    ]

    for title, bullets in sections:
        parts.append(f"<h3>{_esc(title)}</h3><ul>")
        for bullet in bullets:
            parts.append(f"<li>{_esc(bullet)}</li>")
        parts.append("</ul>")

    parts.append("</div>")
    return "\n".join(parts)


def create_appendix_glossary(ctx: RenderContext | None = None) -> str:
    """Create the glossary appendix."""
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("appendix_glossary")
    parts = [f'<div id="sec-appendix_glossary"><h2 class="section">{_esc(heading)}</h2>']

    terms = {
        "Period Charge (£)": "The charge for the specific billing period (not cumulative "
        "balance). Equivalent to 'Total charges for this period' on new EDF invoices.",
        "Amount (£)": "The primary balance figure — typically the current cumulative "
        "account balance on new invoices, or the running balance on HTM exports.",
        "Energy Unit Rate (p/kWh)": "The energy price extracted from the bill, excluding "
        "standing charges. Total bill cost divided by usage is not an equivalent unit rate.",
        "Standing Charge (p/day)": "Daily fixed charge regardless of consumption, as "
        "published on EDF bills.",
        "OFGEM Price Cap": "Maximum price per unit (p/kWh) and daily standing charge "
        "(p/day) that suppliers can charge customers on default/standard variable "
        "tariffs, set quarterly by OFGEM.",
        "Billing Gap": "Period exceeding 60 days (MEDIUM) or 120 days (HIGH) between "
        "consecutive bills where balance accumulates without a new statement.",
        "Z-Score Anomaly": "Data point exceeding 2.5 standard deviations from the mean, "
        "indicating a statistical outlier (≈99% confidence under normality).",
        "IQR Anomaly": "Data point outside 1.5× the interquartile range (Q3−Q1), robust "
        "to non-normal distributions.",
        "Holt-Winters Forecast": "Exponential smoothing with trend and optional "
        "seasonality, suitable for time series with patterns.",
        "MAPE": "Mean Absolute Percentage Error — average of |forecast − actual|/actual "
        "× 100%. Lower is better; <10% considered good for energy billing.",
    }

    glossary_rows = [["Term", "Definition"]]
    for term, definition in terms.items():
        glossary_rows.append([term, definition])
    parts.append(_table(glossary_rows))
    parts.append("</div>")
    return "\n".join(parts)


def create_appendix_full_evidence(
    df: pd.DataFrame,
    filtered: list | None = None,
    config: dict | None = None,
    ctx: RenderContext | None = None,
) -> str:
    """Create the full evidence table appendix.

    The HTML surface mirrors the DOCX row cap so a large dataset does
    not produce an unbounded document.
    """
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("appendix_full_evidence")
    parts = [f'<div id="sec-appendix_full_evidence"><h2 class="section">{_esc(heading)}</h2>']
    parts.append(f"<p>{_esc(EVIDENCE_NOTE)}</p>")

    table_row_cap = 150

    def _evidence_rows(frame: pd.DataFrame) -> list[list[Any]]:
        return [EVIDENCE_HEADERS, *[evidence_cells(row) for row in frame.to_dict("records")]]

    if df.empty:
        parts.append("<p>No records available.</p></div>")
        return "\n".join(parts)

    if "_dt" not in df.columns:
        df["_dt"] = df["Date"].apply(parse_to_sort_date)
    df_sorted = df.sort_values("_dt").reset_index(drop=True)

    total_rows = int(len(df_sorted))
    if total_rows > table_row_cap:
        parts.append(
            _note(
                f"The full evidence table exceeds the document render cap "
                f"({total_rows} rows > {table_row_cap} row limit), so only the first "
                f"{table_row_cap} records are shown below in chronological order. "
                "Please refer to the accompanying Excel workbook for the complete dataset."
            )
        )
        parts.append(_table(_evidence_rows(df_sorted.head(table_row_cap))))
    else:
        parts.append(_table(_evidence_rows(df_sorted)))

    if filtered:
        cont_label = ctx.short_label("appendix_full_evidence").rstrip(".")
        min_amt = fmt_money(config.get("min_amount", 500)) if config else "£500"
        cont_heading = (
            f"{cont_label}. (cont.) Filtered Records (Below {min_amt} Threshold)"
            if cont_label
            else f"Filtered Records (Below {min_amt} Threshold)"
        )
        parts.append(f"<h3>{_esc(cont_heading)}</h3>")
        filt_rows = _evidence_rows(pd.DataFrame(filtered))
        if len(filt_rows) > 1:
            parts.append(_table(filt_rows))

    parts.append("</div>")
    return "\n".join(parts)


# =============================================================================
# MAIN REPORT GENERATOR
# =============================================================================


def generate_html_report(
    records: list[dict],
    output_path: str,
    config: ConfigDict,
    engine: Any = None,
    filtered: list | None = None,
    report_sections: list[str] | set[str] | None = None,
) -> str:
    """Capture legacy inputs once, then render their accepted result."""
    view = capture_report_view(
        records,
        dict(config),
        engine,
        filtered,
        formats=("html",),
        sections=tuple(report_sections) if report_sections is not None else None,
    )
    render_html(view, output_path)
    return output_path


@atomic_output()
def render_html(view: ReportView, output_path: str) -> str:
    """Render the accepted result using the existing HTML layout."""
    section_failures = []
    records, config, engine = view.records, view.config, view.engine
    filtered = engine.filtered_records
    df = view.frame()
    if df.empty:
        if not view.review_notice:
            raise ValueError("No records to report on")
        from edf_bill_fetcher.io.reporters.review_empty import render_empty_review

        return render_empty_review(view, output_path, "html")

    if engine is None:

        class MinimalEngine:
            pdf_count: int = 0
            email_count: int = 0
            filtered_records: list[Any] = []

        engine = MinimalEngine()

    required_cols = {
        "Date": "N/A",
        "Source": "Unknown",
        "Amount (£)": None,
        "Period From": "N/A",
        "Period To": "N/A",
        "Invoice #": "N/A",
        "Period Charge (£)": "N/A",
        "Entry Type": "Unknown",
        "Reading": "N/A",
        "Units (kWh)": "N/A",
        "Standing Chg (p/day)": "N/A",
        "Tariff": "N/A",
        "Attachment Name": "N/A",
        "Details": "",
        "Logic Used": "",
    }
    for col, default in required_cols.items():
        if col not in df.columns:
            df[col] = default

    df["_sort"] = df["Date"].apply(parse_to_sort_date)
    df = df.sort_values("_sort").reset_index(drop=True)

    acc_ref = report_account_label(view)

    dates_parsed = df["Date"].apply(parse_to_sort_date)
    valid_dates = dates_parsed.dropna()
    period_start = fmt_date(valid_dates.min()) if not valid_dates.empty else "Unknown"
    period_end = fmt_date(valid_dates.max()) if not valid_dates.empty else "Unknown"

    charges, payments = view.snapshot.summary.charges, view.snapshot.summary.net_credits
    opening_balance, closing_balance = (
        view.snapshot.summary.opening_balance,
        view.snapshot.summary.closing_balance,
    )

    dispute = view.model("dispute")
    flags, flag_counts = dispute.flags, dispute.counts

    enabled_sections = view.context.selected

    def section_enabled(key: str) -> bool:
        return key in enabled_sections

    render_ctx = RenderContext(enabled_sections)

    body_parts: list[str] = []
    cover = ""
    navigation = ""

    if section_enabled("cover"):
        cover = create_cover_page(
            acc_ref, period_start, period_end, datetime.now().strftime("%d %B %Y")
        )

    if view.review_notice:
        body_parts.append(_note(view.review_notice))
    if view.ocr_notice:
        body_parts.append(_note(view.ocr_notice))

    if section_enabled("toc"):
        try:
            navigation = create_table_of_contents(render_ctx)
        except Exception as e:
            section_failures.append(str(e))  # pragma: no cover - defensive mirror of PDF
            body_parts.append(_note(f"Table of Contents failed: {_esc(e)}"))

    # === SECTION DISPATCH (data-driven — keys/ordering live in REPORT_SECTIONS) ===
    section_builders: dict[str, tuple] = {
        "exec_summary": (
            lambda: {
                "df": df,
                "config": config,
                "account_ref": acc_ref,
                "flag_count": flag_counts,
                "total_records": len(records),
                "total_charges": charges,
                "total_payments": payments,
                "period_start": period_start,
                "period_end": period_end,
                "opening_balance": opening_balance,
                "closing_balance": closing_balance,
                "overview": build_output_overview(view),
            },
            lambda kwargs: create_executive_summary(**kwargs),
        ),
        "key_findings": (
            lambda: {"flags": flags},
            lambda kwargs: create_key_findings_table(**kwargs),
        ),
        "evidence_index": (
            lambda: {"df": df, "engine": engine, "bundle_dir": str(Path(output_path).parent)},
            lambda kwargs: create_evidence_index(**kwargs),
        ),
        "detailed_findings": (
            lambda: {"flags": flags, "df": df},
            lambda kwargs: (
                create_anomaly_detail_section(**kwargs) + create_revision_section(view.revisions)
            ),
        ),
        "timeline": (
            lambda: {"df": df, "flags": flags},
            lambda kwargs: create_timeline_section(**kwargs),
        ),
        "ofgem": (
            lambda: {"df": df, "config": config},
            lambda kwargs: create_ofgem_comparison(**kwargs),
        ),
        "statistical": (
            lambda: {"df": df},
            lambda kwargs: create_statistical_analysis(**kwargs),
        ),
        "payment": (
            lambda: {"df": df},
            lambda kwargs: create_payment_analysis(**kwargs),
        ),
        "forecast": (
            lambda: {"df": view.frame("monthly")},
            lambda kwargs: create_forecast_section(**kwargs),
        ),
        "data_quality": (
            lambda: {"df": df},
            lambda kwargs: create_data_quality_section(**kwargs),
        ),
        "tariff": (
            lambda: {"df": df},
            lambda kwargs: create_tariff_impact_section(**kwargs),
        ),
        "compensation": (
            lambda: {"df": df, "config": config},
            lambda kwargs: create_compensation_section(**kwargs),
        ),
        "appendix_methodology": (
            lambda: {"config": config},
            lambda kwargs: create_appendix_methodology(**kwargs),
        ),
        "appendix_glossary": (
            lambda: {},
            lambda kwargs: create_appendix_glossary(**kwargs),
        ),
        "appendix_full_evidence": (
            lambda: {"df": df, "filtered": filtered, "config": config},
            lambda kwargs: create_appendix_full_evidence(**kwargs),
        ),
    }

    for section in REPORT_SECTIONS:
        if not section_enabled(section.key):
            continue
        entry = section_builders.get(section.key)
        if entry is None:
            raise RuntimeError(
                f"REPORT_SECTIONS lists '{section.key}' but no builder is wired "
                f"in generate_html_report. Add it to section_builders."
            )
        arg_factory, invoke = entry
        try:
            kwargs = arg_factory()
            kwargs["ctx"] = render_ctx
            if section.key in (
                "ofgem",
                "statistical",
                "payment",
                "forecast",
                "data_quality",
                "tariff",
                "compensation",
            ):
                kwargs["payload"] = view.model(section.key)
            body_parts.append(invoke(kwargs))
        except Exception as e:
            section_failures.append(str(e))
            body_parts.append(f'<p class="failed">{_esc(section.title)} failed: {_esc(e)}</p>')

    html_doc = (
        "<!DOCTYPE html>\n"
        '<html lang="en">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        f"<title>{_esc('EDF Energy Ombudsman Evidence Report')}</title>\n"
        f"<style>{_CSS}</style>\n"
        "</head>\n<body>\n"
        '<a class="skip-link" href="#report-content">Skip to report</a><div class="container">\n'
        + cover
        + f'<div class="report-layout{chr(32) + "no-nav" if not navigation else ""}">'
        + navigation
        + '<main id="report-content">'
        + "\n".join(body_parts)
        + "\n</main></div></div>\n</body>\n</html>\n"
    )

    if section_failures:
        raise RuntimeError(
            "Report incomplete; output was not published: " + "; ".join(section_failures)
        )
    with open(output_path, "w", encoding="utf-8") as fh:
        fh.write(html_doc)
    return output_path


def generate_html_from_gui(
    records: list[dict[str, Any]],
    output_path: str,
    config: ConfigDict,
    engine: Any = None,
    filtered: list[dict[str, Any]] | None = None,
) -> tuple[bool, str]:
    """Generate an HTML report for GUI integration (success tuple)."""
    try:
        path = generate_html_report(records, output_path, config, engine, filtered)
        return True, f"Professional HTML report generated:\n{path}"
    except Exception as e:
        return False, f"Failed to generate HTML:\n{e}"
