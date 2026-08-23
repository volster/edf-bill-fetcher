"""HTML Report Generator for EDF Energy Ombudsman Submissions.

Builds a self-contained HTML document from the same ``REPORT_SECTIONS``
registry the PDF and DOCX renderers use (see
``edf_bill_fetcher.io.reporters.pdf_report``).  Section titles and
numbering are derived from that registry via ``RenderContext``, so all
three surfaces agree on headings and the table of contents.

The document carries inline CSS only — no external stylesheets, scripts
or remote assets — so it opens correctly from an offline file share.

Each registry section is rendered by a ``create_<name>(...)`` builder
returning an HTML fragment string.  Sections whose analysis is
chart-heavy in the PDF/DOCX surfaces (statistical, payment, forecast)
render a "not implemented in HTML" placeholder note instead, directing
the reader to the PDF/DOCX/Excel outputs.  A registry key with no
wired builder raises ``RuntimeError`` at dispatch — the same
loud-failure mode the PDF dispatcher uses.
"""

from __future__ import annotations

import html as html_lib
from datetime import datetime
from typing import Any

import pandas as pd

from edf_bill_fetcher.helpers.date_utils import parse_to_sort_date
from edf_bill_fetcher.helpers.formatting import fmt_money, fmt_number
from edf_bill_fetcher.io.reporters.pdf_report import (
    REPORT_SECTIONS,
    RenderContext,
    _compute_balance_extremes,
    _compute_financial_totals,
    _compute_mean_daily,
    _get_package_version,
    _load_ofgem_caps,
    _period_to_ofgem_quarter,
    fmt_date,
)
from edf_bill_fetcher.models.config import ConfigDict

# =============================================================================
# CONSTANTS
# =============================================================================

# Inline stylesheet — deliberately self-contained so the report renders
# offline from a file share without fetching any external asset.
_CSS = """
body { font-family: "Segoe UI", Helvetica, Arial, sans-serif; color: #333333;
       margin: 0; padding: 0; background: #ffffff; }
.container { max-width: 900px; margin: 0 auto; padding: 24px 32px 64px 32px; }
h1 { color: #10367A; font-size: 28px; margin: 0 0 4px 0; }
h2.section { color: #10367A; font-size: 18px; border-bottom: 2px solid #10367A;
             padding-bottom: 6px; margin: 40px 0 12px 0; }
h3 { color: #1B4F9E; font-size: 14px; margin: 20px 0 8px 0; }
p { font-size: 11px; line-height: 1.5; margin: 0 0 10px 0; }
p.note { background: #EBF3FA; border-left: 4px solid #2E75B6; padding: 10px 12px;
         font-style: italic; color: #333333; }
p.muted { color: #666666; font-size: 9px; }
table { border-collapse: collapse; width: 100%; margin: 0 0 16px 0; }
th, td { border: 1px solid #B4C6E7; padding: 5px 7px; font-size: 10px;
         text-align: left; vertical-align: top; }
th { background: #10367A; color: #ffffff; font-weight: bold; }
tr.alt td { background: #EBF3FA; }
.cover { text-align: center; padding: 48px 0 24px 0; }
.cover .subtitle { color: #2E75B6; font-size: 14px; margin: 8px 0 20px 0; }
.cover .confidential { color: #C00000; font-weight: bold; font-size: 11px;
                       margin: 24px 0 6px 0; }
.cover table { width: 60%; margin: 8px auto; }
.toc ol { font-size: 11px; }
.toc a { color: #10367A; text-decoration: none; }
.toc a:hover { text-decoration: underline; }
.failed { color: #C00000; font-style: italic; }
ul { font-size: 11px; line-height: 1.5; }
"""


def _esc(value: object) -> str:
    """HTML-escape a value so user-derived text renders as visible text."""
    return html_lib.escape(str(value), quote=True)


def _table(rows: list[list[Any]], *, header: bool = True) -> str:
    """Render a list-of-rows as a styled HTML table.

    ``header=True`` treats ``rows[0]`` as a header row (navy background).
    Every cell is HTML-escaped; even-indexed body rows get the
    alternating ``tr.alt`` tint.
    """
    if not rows:
        return ""
    parts = ["<table>"]
    body_start = 0
    if header:
        head = rows[0]
        parts.append("<tr>" + "".join(f"<th>{_esc(c)}</th>" for c in head) + "</tr>")
        body_start = 1
    for i, row in enumerate(rows[body_start:], start=body_start):
        cls = ' class="alt"' if i % 2 == 0 else ""
        cells = "".join(f"<td>{_esc(c)}</td>" for c in row)
        parts.append(f"<tr{cls}>{cells}</tr>")
    parts.append("</table>")
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
        ["Period Covered", f"{period_start} to {period_end}"],
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
    parts = ['<div class="toc"><h2 class="section">Table of Contents</h2>']
    if not ctx.sections_in_order:
        parts.append("<p><i>No report sections selected.</i></p>")
        parts.append("</div>")
        return "\n".join(parts)
    parts.append("<ol>")
    for spec in ctx.sections_in_order:
        anchor = f"sec-{spec.section.key}"
        parts.append(
            f'<li><a href="#{anchor}">{_esc(spec.label)} {_esc(spec.section.title)}</a></li>'
        )
    parts.append("</ol></div>")
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
) -> str:
    """Create the executive summary section."""
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("exec_summary")
    net_change = total_charges - total_payments

    parts = [f'<div id="sec-exec_summary"><h2 class="section">{_esc(heading)}</h2>']
    parts.append(
        "<p>This report presents the findings of a comprehensive analysis of EDF Energy "
        f"billing data for account <b>{_esc(account_ref)}</b>, covering the period "
        f"<b>{_esc(period_start)}</b> to <b>{_esc(period_end)}</b>. "
        f"The analysis encompasses <b>{total_records}</b> billing records sourced from "
        "EDF bills (PDF), HTM account exports, and email archives (PST/OST).</p>"
    )
    fin_rows = [
        ["Metric", "Amount"],
        ["Total Charges (Debits)", fmt_money(total_charges)],
        ["Total Payments/Credits", fmt_money(total_payments)],
        ["Net Balance Increase", fmt_money(net_change)],
        [
            "Opening Balance (First Record)",
            fmt_money(opening_balance) if opening_balance is not None else "—",
        ],
        [
            "Closing Balance (Latest Record)",
            fmt_money(closing_balance) if closing_balance is not None else "—",
        ],
    ]
    parts.append("<h3>Financial Summary</h3>")
    parts.append(_table(fin_rows))

    findings: list[str] = []
    if flag_count.get("HIGH", 0) > 0:
        findings.append(
            f"{flag_count['HIGH']} HIGH-severity anomalies detected, including billing "
            "spikes exceeding 50% period-over-period, gaps over 120 days without "
            "billing, and reconciliation mismatches suggesting unrecorded payments or "
            "billing errors."
        )
    if flag_count.get("MEDIUM", 0) > 0:
        findings.append(
            f"{flag_count['MEDIUM']} MEDIUM-severity issues identified, including "
            "billing gaps of 60-120 days, daily rate anomalies 2.5-4x average, and "
            "estimated reading runs."
        )
    if flag_count.get("INFO", 0) > 0:
        findings.append(
            f"{flag_count['INFO']} informational items noted, primarily balance "
            "reductions from payments/credits over £500."
        )
    if not findings:
        findings.append("No significant anomalies detected in the billing data.")

    parts.append("<h3>Key Findings</h3><ul>")
    for finding in findings:
        parts.append(f"<li>{_esc(finding)}</li>")
    parts.append("</ul>")

    parts.append(
        "<h3>Conclusion</h3>"
        "<p>Based on the systematic analysis of all available billing records, this "
        "report identifies multiple instances where EDF Energy's billing practices "
        "deviate from expected norms and regulatory requirements. The documented "
        "anomalies—particularly the high-severity billing spikes, extended billing "
        "gaps, and reconciliation failures—warrant formal investigation by the "
        "Energy Ombudsman. The complainant requests a full billing audit for the "
        "identified periods and appropriate redress for any overcharging.</p>"
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
            "<p>No automated flags were generated. The billing data appears "
            "consistent within established thresholds.</p></div>"
        )
        return "\n".join(parts)

    high = [f for f in flags if f[4] == "HIGH"]
    medium = [f for f in flags if f[4] == "MEDIUM"]
    info = [f for f in flags if f[4] == "INFO"]

    summary_rows = [
        ["Severity", "Count", "Description"],
        ["HIGH", str(len(high)), "Immediate concern — regulatory breach likely"],
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


def create_evidence_index(df: pd.DataFrame, engine: Any, ctx: RenderContext | None = None) -> str:
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
    for src in source_counts.index:
        src_df = df[df["Source"] == src].copy()
        src_df["_dt"] = src_df["Date"].apply(parse_to_sort_date)
        src_df = src_df.sort_values("_dt")

        detail_rows = [["Date", "Invoice #", "Amount", "Period", "Entry Type", "Reading"]]
        for _, row in src_df.iterrows():
            detail_rows.append(
                [
                    fmt_date(row.get("Date")),
                    row.get("Invoice #", "N/A"),
                    fmt_money(row.get("Amount (£)")),
                    f"{fmt_date(row.get('Period From'))}–{fmt_date(row.get('Period To'))}",
                    row.get("Entry Type", ""),
                    row.get("Reading", ""),
                ]
            )
        parts.append(f"<h3>{_esc(src)} ({len(src_df)} records)</h3>")
        parts.append(_table(detail_rows))

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


def create_timeline_section(df: pd.DataFrame, flags: list, ctx: RenderContext | None = None) -> str:
    """Create the chronological timeline of events."""
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("timeline")
    parts = [f'<div id="sec-timeline"><h2 class="section">{_esc(heading)}</h2>']

    events: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        events.append(
            {
                "date": row["Date"],
                "type": row.get("Entry Type", "Record"),
                "amount": row.get("Amount (£)"),
                "detail": f"{row.get('Source', '')} — {str(row.get('Details', ''))[:100]}",
            }
        )
    for ftype, date, amt, detail, _sev in flags:
        events.append(
            {
                "date": date,
                "type": f"⚠ {ftype}",
                "amount": amt,
                "detail": detail,
            }
        )
    events.sort(key=lambda e: parse_to_sort_date(e["date"]) or pd.Timestamp.min)

    timeline_rows = [["Date", "Event Type", "Amount", "Detail"]]
    for ev in events:
        timeline_rows.append(
            [
                fmt_date(ev["date"]),
                ev["type"],
                fmt_money(ev["amount"]) if ev["amount"] else "",
                str(ev["detail"])[:150],
            ]
        )
    parts.append(_table(timeline_rows))
    parts.append("</div>")
    return "\n".join(parts)


def create_ofgem_comparison(
    df: pd.DataFrame, config: dict | None = None, ctx: RenderContext | None = None
) -> str:
    """Create the OFGEM price cap comparison section.

    Mirrors ``pdf_report.create_ofgem_comparison``: effective bill unit
    rate is computed from ``Period Charge (£)`` ÷ ``Units (kWh)`` × 100,
    quarters beyond the published cap table fall back to the
    carry-forward cap, and the summary row verdicts (REVIEW REQUIRED /
    INCOMPLETE / COMPLIANT) match the PDF surface.
    """
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("ofgem")
    parts = [f'<div id="sec-ofgem"><h2 class="section">{_esc(heading)}</h2>']

    parts.append(
        "<p>The following analysis compares the effective unit rates charged on EDF "
        "bills against the OFGEM Default Tariff Cap (Price Cap) for the "
        "corresponding periods. Any charges exceeding the cap may indicate "
        "regulatory non-compliance.</p>"
    )

    ofgem_caps, latest_known_cap = _load_ofgem_caps(auto_carry=True)

    work = df.copy()
    work["_dt"] = work["Date"].apply(parse_to_sort_date)
    work = work.sort_values("_dt").reset_index(drop=True)

    valid_pc = work["Period Charge (£)"].notna() & (work["Period Charge (£)"] != "N/A")
    valid_units = (
        work["Units (kWh)"].notna() & (work["Units (kWh)"] != "N/A") & (work["Units (kWh)"] != "")
    )
    bills = work[valid_pc & valid_units].copy()

    if bills.empty:
        parts.append(
            "<p>No billing records with both Period Charge and Units (kWh) available "
            "for comparison.</p></div>"
        )
        return "\n".join(parts)

    bills["_unit_rate"] = (
        bills["Period Charge (£)"].astype(float) / bills["Units (kWh)"].astype(float) * 100
    )
    bills["_quarter"] = bills["_dt"].apply(_period_to_ofgem_quarter)
    bills = bills[bills["_quarter"].notna()].copy()

    if bills.empty:
        parts.append("<p>No billing records fall within an OFGEM-published cap window.</p></div>")
        return "\n".join(parts)

    MISSING = "—"
    UNAVAILABLE = "CAP DATA UNAVAILABLE"
    CARRIED = "CAP CARRIED FORWARD"
    cap_rows: list[list[Any]] = []
    exceed_count = 0
    unavailable_count = 0
    carried_count = 0
    for quarter in sorted(bills["_quarter"].dropna().unique()):
        avg_rate = bills[bills["_quarter"] == quarter]["_unit_rate"].mean()
        if pd.isna(avg_rate):
            continue
        if quarter not in ofgem_caps:
            if latest_known_cap:
                carried_count += 1
                cap_rate = latest_known_cap["unit_rate"]
                diff = avg_rate - cap_rate
                status = (
                    f"EXCEEDS CAP ({CARRIED})"
                    if diff > 0
                    else f"AT CAP ({CARRIED})"
                    if abs(diff) < 0.01
                    else f"BELOW CAP ({CARRIED})"
                )
                if diff > 0:
                    exceed_count += 1
            else:
                unavailable_count += 1
                cap_rate = MISSING
                diff = MISSING
                status = UNAVAILABLE
        else:
            cap_rate = ofgem_caps[quarter]["unit_rate"]
            diff = avg_rate - cap_rate
            status = "EXCEEDS CAP" if diff > 0 else "AT CAP" if abs(diff) < 0.01 else "BELOW CAP"
            if diff > 0:
                exceed_count += 1
        cap_rows.append(
            [
                quarter,
                fmt_number(avg_rate, 2),
                fmt_number(cap_rate, 2) if isinstance(cap_rate, float) else cap_rate,
                fmt_number(diff, 2) if isinstance(diff, float) else diff,
                status,
            ]
        )

    if exceed_count > 0:
        summary_diff = f"{exceed_count} periods exceed cap"
        summary_status = "REVIEW REQUIRED"
    elif unavailable_count > 0:
        summary_diff = f"{unavailable_count} period(s) not benchmarked"
        summary_status = "INCOMPLETE"
    elif carried_count > 0:
        summary_diff = f"{carried_count} period(s) used carried-forward cap"
        summary_status = "COMPLIANT (CARRIED)"
    else:
        summary_diff = "No exceedances"
        summary_status = "COMPLIANT"
    cap_rows.append(["OVERALL", "—", "—", summary_diff, summary_status])

    header = ["Period", "Bill Unit Rate (p/kWh)", "OFGEM Cap (p/kWh)", "Difference", "Status"]
    parts.append(_table([header, *cap_rows]))
    parts.append(
        '<p class="muted"><b>Methodology:</b> Unit rates calculated as '
        "Period Charge (£) ÷ Units (kWh) × 100. Only records with both "
        "Period Charge and Units (kWh) are included. OFGEM cap data sourced "
        "from official Default Tariff Cap publications.</p>"
    )
    parts.append("</div>")
    return "\n".join(parts)


def create_statistical_analysis(df: pd.DataFrame, ctx: RenderContext | None = None) -> str:
    """Render the statistical analysis placeholder section."""
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("statistical")
    return (
        f'<div id="sec-statistical"><h2 class="section">{_esc(heading)}</h2>'
        f"{_not_implemented_in_html(heading)}</div>"
    )


def create_payment_analysis(df: pd.DataFrame, ctx: RenderContext | None = None) -> str:
    """Render the payment & credit analysis placeholder section."""
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("payment")
    return (
        f'<div id="sec-payment"><h2 class="section">{_esc(heading)}</h2>'
        f"{_not_implemented_in_html(heading)}</div>"
    )


def create_forecast_section(df: pd.DataFrame, ctx: RenderContext | None = None) -> str:
    """Render the forecast & projection placeholder section."""
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("forecast")
    return (
        f'<div id="sec-forecast"><h2 class="section">{_esc(heading)}</h2>'
        f"{_not_implemented_in_html(heading)}</div>"
    )


def create_data_quality_section(df: pd.DataFrame, ctx: RenderContext | None = None) -> str:
    """Create the data quality assessment section."""
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("data_quality")
    parts = [f'<div id="sec-data_quality"><h2 class="section">{_esc(heading)}</h2>']

    total = len(df)
    date_parsed = df["Date"].apply(lambda x: parse_to_sort_date(x) is not pd.NaT).sum()
    amt_complete = df["Amount (£)"].notna().sum()
    period_complete = (df["Period From"] != "N/A").sum()
    reading_classified = (df["Reading"] != "N/A").sum() if "Reading" in df.columns else 0
    dup_count = df.duplicated(subset=["Date", "Amount (£)"]).sum()

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


def create_tariff_impact_section(df: pd.DataFrame, ctx: RenderContext | None = None) -> str:
    """Create the tariff impact analysis section."""
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

    tariff_data = df.dropna(subset=["Tariff"])
    tariff_data = tariff_data[tariff_data["Tariff"] != "N/A"].copy()
    if tariff_data.empty:
        parts.append("<p>No valid tariff records found.</p></div>")
        return "\n".join(parts)

    tariff_data["unit_rate_num"] = pd.to_numeric(tariff_data["Unit Rate (p/kWh)"], errors="coerce")
    tariff_data = tariff_data.dropna(subset=["unit_rate_num"])
    if tariff_data.empty:
        parts.append("<p>No computable unit rates found.</p></div>")
        return "\n".join(parts)

    tariff_stats = (
        tariff_data.groupby("Tariff")
        .agg(
            count=("unit_rate_num", "count"),
            avg_rate=("unit_rate_num", "mean"),
            median_rate=("unit_rate_num", "median"),
            min_rate=("unit_rate_num", "min"),
            max_rate=("unit_rate_num", "max"),
        )
        .reset_index()
    )

    parts.append("<h3>Unit Rate by Tariff</h3>")
    tariff_rows = [["Tariff", "Records", "Avg Rate (p/kWh)", "Median", "Min", "Max"]]
    for _, row in tariff_stats.iterrows():
        tariff_rows.append(
            [
                row["Tariff"],
                str(int(row["count"])),
                fmt_number(row["avg_rate"], 2),
                fmt_number(row["median_rate"], 2),
                fmt_number(row["min_rate"], 2),
                fmt_number(row["max_rate"], 2),
            ]
        )
    parts.append(_table(tariff_rows))

    tariff_data["_dt"] = tariff_data["Date"].apply(parse_to_sort_date)
    tariff_data = tariff_data.sort_values("_dt")
    changes = tariff_data["Tariff"].ne(tariff_data["Tariff"].shift()).cumsum()
    n_changes = int(changes.max()) if not changes.empty else 0
    parts.append(f"<h3>Tariff Changes Detected: {n_changes}</h3>")
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
                "All billing records were extracted from three primary source types:",
                "• PDF Bills: EDF Energy invoices (both legacy and new KI/KCR formats) "
                "processed via pdfplumber with format-specific regex extraction.",
                "• HTM Export: EDF MyAccount 'Payments and Invoices' export parsed via "
                "BeautifulSoup with pattern matching for charge/payment/reversal entries.",
                "• PST/OST Email Archives: Outlook data files processed via libpff-python, "
                "extracting email bodies (HTML/plain text/RTF) and PDF attachments.",
            ],
        ),
        (
            "A.2 Amount Extraction Logic",
            [
                "Two complementary strategies ensure comprehensive amount detection:",
                "1. Smart Context Search: 10 prioritized regex patterns targeting specific "
                "EDF billing language. Patterns execute in priority order; first match wins.",
                "2. Large Amount Fallback: scans all £ amounts ≥ minimum threshold, "
                "selecting the largest. Used when context patterns fail.",
            ],
        ),
        (
            "A.3 Deduplication",
            [
                "Multi-pass deduplication matches the same bill across sources:",
                "• Pass 1: Exact match on Period To date + Amount.",
                "• Pass 2: For records without period info, match by Amount within a "
                "60-day window of any kept record.",
            ],
        ),
        (
            "A.4 Configuration Used",
            [
                f"Minimum Amount Threshold: {fmt_money(config.get('min_amount', 500))}",
                f"Analysis Threshold: {fmt_money(config.get('analysis_min', 500))}",
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
        "Unit Rate (p/kWh)": "Effective price per kWh = Period Charge ÷ Units (kWh) × 100. "
        "Includes both energy and standing charge components unless separated.",
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
    parts.append(
        "<p>This appendix contains the complete set of billing records used in this "
        "analysis. Records are sorted chronologically by date.</p>"
    )

    table_row_cap = 150
    evidence_header = [
        "Date",
        "Source",
        "Entry Type",
        "Amount (£)",
        "Period Charge (£)",
        "Period From",
        "Period To",
        "Invoice #",
        "Reading",
        "Units (kWh)",
        "Standing Chg (p/day)",
        "Attachment",
        "Details",
    ]

    def _evidence_rows(frame: pd.DataFrame) -> list[list[Any]]:
        rows: list[list[Any]] = [evidence_header]
        for _, row in frame.iterrows():
            rows.append(
                [
                    fmt_date(row.get("Date", "")),
                    str(row.get("Source", ""))[:30],
                    str(row.get("Entry Type", "")),
                    fmt_money(row.get("Amount (£)", 0)),
                    fmt_money(row.get("Period Charge (£)", 0))
                    if row.get("Period Charge (£)") not in ("", "N/A", None)
                    else "N/A",
                    fmt_date(row.get("Period From", "")),
                    fmt_date(row.get("Period To", "")),
                    str(row.get("Invoice #", ""))[:15],
                    str(row.get("Reading", ""))[:15],
                    str(row.get("Units (kWh)", ""))[:10],
                    str(row.get("Standing Chg (p/day)", ""))[:10],
                    str(row.get("Attachment Name", ""))[:20],
                    str(row.get("Details", ""))[:50],
                ]
            )
        return rows

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
    """Generate an HTML report from the supplied records.

    The output sections are derived from ``REPORT_SECTIONS`` via the
    dispatcher wired up below.  Section selection follows the PDF/DOCX
    convention (``config["report_sections"]``); the explicit
    ``report_sections`` argument overrides config when provided.

    Args:
        records: List of extracted billing records.
        output_path: Path to save the HTML document.
        config: Configuration dictionary.
        engine: EvidenceEngine instance (for metadata).
        filtered: Filtered-out records (below threshold).
        report_sections: Optional explicit section selection; when
            ``None`` the selection is read from ``config``.

    Returns:
        Path to the generated HTML file.

    """
    if not records:
        raise ValueError("No records to report on")

    df = pd.DataFrame(records)
    if df.empty:
        raise ValueError("Records DataFrame is empty")

    if engine is None:

        class MinimalEngine:
            pdf_count: int = 0
            email_count: int = 0
            filtered_records: list[Any] = []

        engine = MinimalEngine()

    required_cols = {
        "Date": "01/01/1970",
        "Source": "Unknown",
        "Amount (£)": 0.0,
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

    acc_ref = config.get("report_account_ref") or config.get("acc_num") or "Unknown"

    dates_parsed = df["Date"].apply(parse_to_sort_date)
    valid_dates = dates_parsed.dropna()
    period_start = fmt_date(valid_dates.min()) if not valid_dates.empty else "Unknown"
    period_end = fmt_date(valid_dates.max()) if not valid_dates.empty else "Unknown"

    charges, payments = _compute_financial_totals(df)
    opening_balance, closing_balance = _compute_balance_extremes(df)

    from edf_bill_fetcher.processors.analysis import compute_dispute_flags

    if "_dt" not in df.columns:
        df["_dt"] = df["Date"].apply(parse_to_sort_date)
    df_sorted = df.sort_values("_dt").reset_index(drop=True)
    mean_daily = _compute_mean_daily(df_sorted)
    flags, flag_counts = compute_dispute_flags(df_sorted, mean_daily)

    # Section selection: explicit argument wins, then config, then all
    # implementable sections (backward compatibility, as in the PDF).
    all_sections = {
        "cover",
        "toc",
        "exec_summary",
        "key_findings",
        "evidence_index",
        "detailed_findings",
        "timeline",
        "ofgem",
        "statistical",
        "payment",
        "forecast",
        "data_quality",
        "tariff",
        "appendix_methodology",
        "appendix_glossary",
        "appendix_full_evidence",
    }
    if report_sections is not None:
        enabled_sections = set(report_sections)
    else:
        enabled_sections = set(config.get("report_sections", []))
    if not enabled_sections:
        enabled_sections = all_sections

    def section_enabled(key: str) -> bool:
        return key in enabled_sections

    render_ctx = RenderContext(enabled_sections)

    body_parts: list[str] = []

    if section_enabled("cover"):
        body_parts.append(
            create_cover_page(
                acc_ref, period_start, period_end, datetime.now().strftime("%d %B %Y")
            )
        )

    if section_enabled("toc"):
        try:
            body_parts.append(create_table_of_contents(render_ctx))
        except Exception as e:  # pragma: no cover - defensive mirror of PDF
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
            },
            lambda kwargs: create_executive_summary(**kwargs),
        ),
        "key_findings": (
            lambda: {"flags": flags},
            lambda kwargs: create_key_findings_table(**kwargs),
        ),
        "evidence_index": (
            lambda: {"df": df, "engine": engine},
            lambda kwargs: create_evidence_index(**kwargs),
        ),
        "detailed_findings": (
            lambda: {"flags": flags, "df": df},
            lambda kwargs: create_anomaly_detail_section(**kwargs),
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
            lambda: {"df": df},
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
            body_parts.append(invoke(kwargs))
        except Exception as e:
            body_parts.append(f'<p class="failed">{_esc(section.title)} failed: {_esc(e)}</p>')

    html_doc = (
        "<!DOCTYPE html>\n"
        '<html lang="en">\n<head>\n<meta charset="utf-8">\n'
        f"<title>{_esc('EDF Energy Ombudsman Evidence Report')}</title>\n"
        f"<style>{_CSS}</style>\n"
        "</head>\n<body>\n"
        '<div class="container">\n' + "\n".join(body_parts) + "\n</div>\n</body>\n</html>\n"
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
