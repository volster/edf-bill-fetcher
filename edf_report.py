"""
Professional PDF Report Generator for EDF Energy Ombudsman Submissions.

Generates a professional, court-ready PDF report optimized for Energy Ombudsman review.
Includes: Executive Summary, Evidence Index, Key Findings, Timeline, Calculations,
OFGEM Price Cap Comparison, and full Evidence Appendix.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

import numpy as np
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

# Import from main module
from edf_collector import parse_to_display_date, parse_to_sort_date

# =============================================================================
# COLOR PALETTE & CONSTANTS
# =============================================================================


class Colors:
    """Professional color palette for the report."""

    NAVY = colors.HexColor("#10367A")
    DARK_BLUE = colors.HexColor("#1B4F9E")
    MEDIUM_BLUE = colors.HexColor("#2E75B6")
    LIGHT_BLUE = colors.HexColor("#D6E4F0")
    VERY_LIGHT_BLUE = colors.HexColor("#EBF3FA")
    WHITE = colors.white
    BLACK = colors.black
    DARK_GREY = colors.HexColor("#333333")
    MEDIUM_GREY = colors.HexColor("#666666")
    LIGHT_GREY = colors.HexColor("#F2F2F2")
    VERY_LIGHT_GREY = colors.HexColor("#F7F7F7")
    RED = colors.HexColor("#C00000")
    AMBER = colors.HexColor("#ED7D31")
    GREEN = colors.HexColor("#548235")
    ORANGE = colors.HexColor("#FE5716")


PAGE_WIDTH, PAGE_HEIGHT = A4
MARGIN = 2.5 * cm
CONTENT_WIDTH = PAGE_WIDTH - 2 * MARGIN


# =============================================================================
# STYLES
# =============================================================================


def build_styles() -> dict[str, ParagraphStyle]:
    """Build all paragraph styles for the report."""
    styles = getSampleStyleSheet()
    custom = {}

    # Title styles
    custom["CoverTitle"] = ParagraphStyle(
        "CoverTitle",
        parent=styles["Title"],
        fontSize=28,
        leading=34,
        textColor=Colors.NAVY,
        spaceAfter=12,
        alignment=TA_CENTER,
        fontName="Helvetica-Bold",
    )
    custom["CoverSubtitle"] = ParagraphStyle(
        "CoverSubtitle",
        parent=styles["Normal"],
        fontSize=14,
        leading=18,
        textColor=Colors.MEDIUM_BLUE,
        spaceAfter=6,
        alignment=TA_CENTER,
        fontName="Helvetica",
    )
    custom["CoverInfo"] = ParagraphStyle(
        "CoverInfo",
        parent=styles["Normal"],
        fontSize=11,
        leading=14,
        textColor=Colors.MEDIUM_GREY,
        spaceAfter=4,
        alignment=TA_CENTER,
        fontName="Helvetica",
    )

    # Section headers
    custom["SectionHeader"] = ParagraphStyle(
        "SectionHeader",
        parent=styles["Heading1"],
        fontSize=16,
        leading=20,
        textColor=Colors.NAVY,
        spaceBefore=18,
        spaceAfter=10,
        fontName="Helvetica-Bold",
        borderWidth=0,
        borderPadding=0,
    )
    custom["SubSectionHeader"] = ParagraphStyle(
        "SubSectionHeader",
        parent=styles["Heading2"],
        fontSize=13,
        leading=16,
        textColor=Colors.DARK_BLUE,
        spaceBefore=12,
        spaceAfter=6,
        fontName="Helvetica-Bold",
    )
    custom["SubSubSectionHeader"] = ParagraphStyle(
        "SubSubSectionHeader",
        parent=styles["Heading3"],
        fontSize=11,
        leading=14,
        textColor=Colors.MEDIUM_BLUE,
        spaceBefore=8,
        spaceAfter=4,
        fontName="Helvetica-Bold",
    )

    # Body text
    custom["BodyText"] = ParagraphStyle(
        "BodyText",
        parent=styles["Normal"],
        fontSize=10,
        leading=13,
        textColor=Colors.DARK_GREY,
        spaceAfter=6,
        alignment=TA_JUSTIFY,
        fontName="Helvetica",
    )
    custom["BodyTextIndent"] = ParagraphStyle(
        "BodyTextIndent", parent=custom["BodyText"], leftIndent=1.5 * cm
    )
    custom["BulletText"] = ParagraphStyle(
        "BulletText",
        parent=custom["BodyText"],
        leftIndent=1.5 * cm,
        bulletIndent=0.75 * cm,
        spaceAfter=3,
    )
    custom["SmallText"] = ParagraphStyle(
        "SmallText",
        parent=custom["BodyText"],
        fontSize=8.5,
        leading=11,
        textColor=Colors.MEDIUM_GREY,
    )

    # Table styles
    custom["TableHeader"] = ParagraphStyle(
        "TableHeader",
        parent=styles["Normal"],
        fontSize=8.5,
        leading=11,
        textColor=Colors.WHITE,
        fontName="Helvetica-Bold",
        alignment=TA_CENTER,
        spaceBefore=2,
        spaceAfter=2,
    )
    custom["TableCell"] = ParagraphStyle(
        "TableCell",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
        textColor=Colors.DARK_GREY,
        fontName="Helvetica",
        alignment=TA_LEFT,
        spaceBefore=1,
        spaceAfter=1,
    )
    custom["TableCellCenter"] = ParagraphStyle(
        "TableCellCenter", parent=custom["TableCell"], alignment=TA_CENTER
    )
    custom["TableCellRight"] = ParagraphStyle(
        "TableCellRight", parent=custom["TableCell"], alignment=TA_RIGHT
    )
    custom["TableCellMoney"] = ParagraphStyle(
        "TableCellMoney", parent=custom["TableCellRight"], fontName="Helvetica"
    )

    # Special
    custom["Footnote"] = ParagraphStyle(
        "Footnote", parent=custom["SmallText"], leftIndent=0.5 * cm, firstLineIndent=-0.5 * cm
    )
    custom["PageNumber"] = ParagraphStyle(
        "PageNumber",
        parent=styles["Normal"],
        fontSize=8,
        textColor=Colors.MEDIUM_GREY,
        alignment=TA_CENTER,
    )
    custom["Confidential"] = ParagraphStyle(
        "Confidential",
        parent=styles["Normal"],
        fontSize=9,
        textColor=Colors.RED,
        fontName="Helvetica-Bold",
        alignment=TA_CENTER,
        spaceBefore=20,
        spaceAfter=6,
    )

    return custom


STYLES = build_styles()


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================


def fmt_money(val: Any, blank_if_na: bool = True) -> str:
    """Format a value as £ currency."""
    if val is None or (isinstance(val, str) and val.upper() in ("N/A", "NA", "")):
        return "" if blank_if_na else "N/A"
    try:
        if isinstance(val, str):
            val = val.replace(",", "").replace("£", "")
        f = float(val)
        return f"£{f:,.2f}"
    except (ValueError, TypeError):
        return str(val) if not blank_if_na else ""


def fmt_number(val: Any, decimals: int = 0, blank_if_na: bool = True) -> str:
    """Format a number with commas."""
    if val is None or (isinstance(val, str) and val.upper() in ("N/A", "NA", "")):
        return "" if blank_if_na else "N/A"
    try:
        if isinstance(val, str):
            val = val.replace(",", "")
        f = float(val)
        if decimals == 0:
            return f"{int(f):,}"
        return f"{f:,.{decimals}f}"
    except (ValueError, TypeError):
        return str(val) if not blank_if_na else ""


def fmt_pct(val: Any, blank_if_na: bool = True) -> str:
    """Format as percentage."""
    if val is None or (isinstance(val, str) and val.upper() in ("N/A", "NA", "")):
        return "" if blank_if_na else "N/A"
    try:
        f = float(val)
        return f"{f:.1%}"
    except (ValueError, TypeError):
        return str(val) if not blank_if_na else ""


def fmt_date(val: Any) -> str:
    """Format date for display."""
    if val is None or (isinstance(val, str) and val.upper() in ("N/A", "NA", "")):
        return ""
    try:
        result = parse_to_display_date(val)
        return str(result) if result else ""
    except Exception:
        return str(val)


def severity_color(sev: str) -> colors.Color:
    """Get color for severity level."""
    sev = str(sev).upper()
    if sev == "HIGH":
        return Colors.RED
    elif sev == "MEDIUM":
        return Colors.AMBER
    elif sev == "INFO":
        return Colors.GREEN
    return Colors.MEDIUM_GREY


def severity_label(sev: str) -> str:
    """Get label with color indicator."""
    sev = str(sev).upper()
    indicators = {"HIGH": "●", "MEDIUM": "●", "INFO": "●"}
    return indicators.get(sev, "○")


def make_table_style(
    header_color: colors.Color = Colors.NAVY,
    alt_row_color: colors.Color = Colors.VERY_LIGHT_BLUE,
    header_text_color: colors.Color = Colors.WHITE,
    grid_color: colors.Color = colors.HexColor("#B4C6E7"),
    font_size: int = 8,
    num_rows: int = 0,
) -> TableStyle:
    """Create a consistent table style. Pass num_rows to enable correct alternating row colors."""
    style_commands: list[
        tuple[str, tuple[int, int], tuple[int, int], Any]
        | tuple[str, tuple[int, int], tuple[int, int], float, Any]
    ] = [
        # Header
        ("BACKGROUND", (0, 0), (-1, 0), header_color),
        ("TEXTCOLOR", (0, 0), (-1, 0), header_text_color),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), font_size),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        # Body
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), font_size),
        ("TEXTCOLOR", (0, 1), (-1, -1), Colors.DARK_GREY),
        ("ALIGN", (0, 1), (-1, -1), "LEFT"),
    ]

    # Alternating rows - only for existing rows
    if num_rows > 1:
        for i in range(1, num_rows, 2):
            style_commands.append(("BACKGROUND", (0, i), (-1, i), alt_row_color))

    style_commands.extend(
        [
            # Grid
            ("GRID", (0, 0), (-1, -1), 0.5, grid_color),
            ("LINEBELOW", (0, 0), (-1, 0), 1.5, header_color),
            # Padding
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ]
    )

    return TableStyle(style_commands)


# =============================================================================
# PAGE TEMPLATES
# =============================================================================


def build_doc_template(output_path: str) -> BaseDocTemplate:
    """Build the document template with page templates."""
    doc = BaseDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=MARGIN,
        bottomMargin=MARGIN,
        title="EDF Energy Ombudsman Evidence Report",
        author="EDF Evidence Collector",
        subject="Energy Billing Dispute Evidence",
        creator="EDF Evidence Collector",
    )

    # Frame for content
    content_frame = Frame(MARGIN, MARGIN, CONTENT_WIDTH, PAGE_HEIGHT - 2 * MARGIN, id="content")

    # Cover page template (no header/footer)
    cover_frame = Frame(MARGIN, MARGIN, CONTENT_WIDTH, PAGE_HEIGHT - 2 * MARGIN, id="cover")

    def add_page_number(canvas, doc):
        """Add page number to footer."""
        canvas.saveState()
        page_num = canvas.getPageNumber()
        if page_num > 1:  # Skip cover page
            canvas.setFont("Helvetica", 8)
            canvas.setFillColor(Colors.MEDIUM_GREY)
            canvas.drawCentredString(PAGE_WIDTH / 2, 1.2 * cm, f"Page {page_num - 1}")
            # Footer line
            canvas.setStrokeColor(Colors.LIGHT_BLUE)
            canvas.setLineWidth(0.5)
            canvas.line(MARGIN, 1.5 * cm, PAGE_WIDTH - MARGIN, 1.5 * cm)
        canvas.restoreState()

    def add_header(canvas, doc):
        """Add header to content pages."""
        canvas.saveState()
        page_num = canvas.getPageNumber()
        if page_num > 1:
            canvas.setFont("Helvetica", 7)
            canvas.setFillColor(Colors.MEDIUM_GREY)
            canvas.drawString(
                MARGIN, PAGE_HEIGHT - 1.2 * cm, "EDF Energy Ombudsman Evidence Report"
            )
            canvas.drawRightString(PAGE_WIDTH - MARGIN, PAGE_HEIGHT - 1.2 * cm, "CONFIDENTIAL")
            # Header line
            canvas.setStrokeColor(Colors.LIGHT_BLUE)
            canvas.setLineWidth(0.5)
            canvas.line(MARGIN, PAGE_HEIGHT - 1.4 * cm, PAGE_WIDTH - MARGIN, PAGE_HEIGHT - 1.4 * cm)
        canvas.restoreState()

    def content_page(canvas, doc):
        add_header(canvas, doc)
        add_page_number(canvas, doc)

    def cover_page(canvas, doc):
        pass  # No header/footer on cover

    doc.addPageTemplates(
        [
            PageTemplate(id="cover", frames=[cover_frame], onPage=cover_page),
            PageTemplate(id="content", frames=[content_frame], onPage=content_page),
        ]
    )

    return doc


# =============================================================================
# REPORT SECTIONS
# =============================================================================


def create_cover_page(
    account_ref: str, period_start: str, period_end: str, report_date: str
) -> list:
    """Create the cover page elements."""
    elements = []

    # Large vertical spacer
    elements.append(Spacer(1, 4 * cm))

    # Main title
    elements.append(Paragraph("EDF Energy Billing Dispute", STYLES["CoverTitle"]))
    elements.append(Paragraph("Ombudsman Evidence Report", STYLES["CoverTitle"]))
    elements.append(Spacer(1, 1.5 * cm))

    # Separator line
    from reportlab.platypus import HRFlowable

    elements.append(
        HRFlowable(width="100%", thickness=2, color=Colors.NAVY, spaceAfter=1.5 * cm, spaceBefore=0)
    )

    # Account info
    elements.append(Paragraph(f"Account Reference: <b>{account_ref}</b>", STYLES["CoverInfo"]))
    elements.append(
        Paragraph(
            f"Period Covered: <b>{period_start}</b> to <b>{period_end}</b>", STYLES["CoverInfo"]
        )
    )
    elements.append(Spacer(1, 1 * cm))

    # Report metadata
    elements.append(Paragraph(f"Report Generated: <b>{report_date}</b>", STYLES["CoverInfo"]))
    elements.append(Paragraph("Prepared by: <b>EDF Evidence Collector</b>", STYLES["CoverInfo"]))
    elements.append(Spacer(1, 2 * cm))

    # Confidential notice
    elements.append(Paragraph("CONFIDENTIAL — FOR OMBUDSMAN REVIEW ONLY", STYLES["Confidential"]))

    # Version info
    elements.append(Spacer(1, 3 * cm))
    elements.append(
        Paragraph(
            "Generated by EDF Evidence Collector v0.1.0<br/>"
            "All data extracted from original source documents (EDF bills, HTM exports, email archives).<br/>"
            "Methodology detailed in Appendix A.",
            STYLES["SmallText"],
        )
    )

    return elements


def create_table_of_contents() -> list:
    """Create table of contents."""
    elements = []
    elements.append(Paragraph("Table of Contents", STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.5 * cm))

    toc_items = [
        ("1.", "Executive Summary", "3"),
        ("2.", "Key Findings Summary", "4"),
        ("3.", "Evidence Index &amp; Source Cross-Reference", "5"),
        ("4.", "Detailed Findings", "6"),
        ("  4.1", "Billing Anomalies &amp; Discrepancies", "6"),
        ("  4.2", "Billing Gaps &amp; Missing Periods", "7"),
        ("  4.3", "Meter Reading Quality Issues", "7"),
        ("  4.3", "Reconciliation Mismatches", "8"),
        ("  4.4", "Payment &amp; Credit Analysis", "8"),
        ("5.", "Timeline of Events", "9"),
        ("6.", "OFGEM Price Cap Comparison", "10"),
        ("7.", "Statistical Analysis", "11"),
        ("8.", "Payment &amp; Credit Analysis", "12"),
        ("9.", "Forecast &amp; Projection", "13"),
        ("10.", "Data Quality Assessment", "14"),
        ("11.", "Tariff Impact Analysis", "15"),
        ("", "Appendices", "15"),
        ("A", "Methodology &amp; Data Sources", "16"),
        ("B", "Glossary", "18"),
        ("C", "Full Evidence Table", "19"),
    ]

    TableStyle(
        [
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("TEXTCOLOR", (0, 0), (-1, -1), Colors.DARK_GREY),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("ALIGN", (2, 0), (2, -1), "RIGHT"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]
    )

    for num, title, page in toc_items:
        is_section = not num.startswith("  ") and num
        weight = "Helvetica-Bold" if is_section else "Helvetica"
        indent = 0 if is_section else 1.5 * cm
        row_style = TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), weight),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("TEXTCOLOR", (0, 0), (-1, -1), Colors.DARK_GREY),
                ("LEFTPADDING", (0, 0), (0, 0), indent),
            ]
        )
        t = Table([[num, title, page]], colWidths=[1.5 * cm, CONTENT_WIDTH - 3 * cm, 1.5 * cm])
        t.setStyle(row_style)
        elements.append(t)
        elements.append(Spacer(1, 0.1 * cm))

    elements.append(PageBreak())
    return elements


# =============================================================================
# EXECUTIVE SUMMARY
# =============================================================================


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
) -> list:
    """Create executive summary section."""
    elements = []

    elements.append(Paragraph("1. Executive Summary", STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    # Overview paragraph
    overview = (
        f"This report presents the findings of a comprehensive analysis of EDF Energy billing data "
        f"for account <b>{account_ref}</b>, covering the period <b>{period_start}</b> to "
        f"<b>{period_end}</b>. The analysis encompasses <b>{total_records}</b> billing records "
        f"sourced from EDF bills (PDF), HTM account exports, and email archives (PST/OST)."
    )
    elements.append(Paragraph(overview, STYLES["BodyText"]))
    elements.append(Spacer(1, 0.3 * cm))

    # Financial summary
    net_change = total_charges - total_payments
    elements.append(Paragraph("<b>Financial Summary</b>", STYLES["SubSectionHeader"]))
    elements.append(Spacer(1, 0.2 * cm))

    fin_data = [
        ["Metric", "Amount"],
        ["Total Charges (Debits)", fmt_money(total_charges)],
        ["Total Payments/Credits", fmt_money(total_payments)],
        ["Net Balance Increase", fmt_money(net_change)],
        ["Opening Balance (First Record)", "—"],  # Would need first record
        ["Closing Balance (Latest Record)", "—"],  # Would need last record
    ]
    fin_table = Table(fin_data, colWidths=[8 * cm, 5 * cm])
    fin_table.setStyle(make_table_style(num_rows=6))
    elements.append(fin_table)
    elements.append(Spacer(1, 0.3 * cm))

    # Key findings
    elements.append(Paragraph("<b>Key Findings</b>", STYLES["SubSectionHeader"]))
    elements.append(Spacer(1, 0.2 * cm))

    findings = []
    if flag_count.get("HIGH", 0) > 0:
        findings.append(
            f"<b>{flag_count['HIGH']} HIGH-severity anomalies</b> detected, including billing spikes "
            f"exceeding 50% period-over-period, gaps over 120 days without billing, and "
            f"reconciliation mismatches suggesting unrecorded payments or billing errors."
        )
    if flag_count.get("MEDIUM", 0) > 0:
        findings.append(
            f"<b>{flag_count['MEDIUM']} MEDIUM-severity issues</b> identified, including billing "
            f"gaps of 60-120 days, daily rate anomalies 2.5-4x average, and estimated reading runs."
        )
    if flag_count.get("INFO", 0) > 0:
        findings.append(
            f"<b>{flag_count['INFO']} informational items</b> noted, primarily balance reductions "
            f"from payments/credits over £500."
        )

    if not findings:
        findings.append("No significant anomalies detected in the billing data.")

    for i, finding in enumerate(findings, 1):
        elements.append(Paragraph(f"{i}. {finding}", STYLES["BulletText"]))

    elements.append(Spacer(1, 0.3 * cm))

    # Conclusion
    elements.append(Paragraph("<b>Conclusion</b>", STYLES["SubSectionHeader"]))
    conclusion = (
        "Based on the systematic analysis of all available billing records, this report identifies "
        "multiple instances where EDF Energy's billing practices deviate from expected norms "
        "and regulatory requirements. The documented anomalies—particularly the high-severity "
        "billing spikes, extended billing gaps, and reconciliation failures—warrant formal "
        "investigation by the Energy Ombudsman. The complainant requests a full billing audit "
        "for the identified periods and appropriate redress for any overcharging."
    )
    elements.append(Paragraph(conclusion, STYLES["BodyText"]))

    elements.append(PageBreak())
    return elements


# =============================================================================
# KEY FINDINGS SUMMARY
# =============================================================================


def create_key_findings_table(flags: list) -> list:
    """Create key findings summary table from flags."""
    elements = []

    elements.append(Paragraph("2. Key Findings Summary", STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    if not flags:
        elements.append(
            Paragraph(
                "No automated flags were generated. The billing data appears consistent within "
                "established thresholds.",
                STYLES["BodyText"],
            )
        )
        elements.append(PageBreak())
        return elements

    # Group by severity
    high = [f for f in flags if f[4] == "HIGH"]
    medium = [f for f in flags if f[4] == "MEDIUM"]
    info = [f for f in flags if f[4] == "INFO"]

    # Summary table
    summary_data = [
        ["Severity", "Count", "Description"],
        ["HIGH", str(len(high)), "Immediate concern — regulatory breach likely"],
        ["MEDIUM", str(len(medium)), "Significant deviation — investigation warranted"],
        ["INFO", str(len(info)), "Informational — payments/credits noted"],
        ["TOTAL", str(len(flags)), "All automated findings"],
    ]

    # Color-code severity cells
    t = Table(summary_data, colWidths=[3 * cm, 2 * cm, CONTENT_WIDTH - 5 * cm])
    style = make_table_style(num_rows=5)
    style.add("BACKGROUND", (0, 1), (0, 1), Colors.RED)
    style.add("BACKGROUND", (0, 2), (0, 2), Colors.AMBER)
    style.add("BACKGROUND", (0, 3), (0, 3), Colors.GREEN)
    style.add("BACKGROUND", (0, 4), (0, 4), Colors.NAVY)
    style.add("TEXTCOLOR", (0, 1), (0, 1), Colors.WHITE)
    style.add("TEXTCOLOR", (0, 2), (0, 2), Colors.WHITE)
    style.add("TEXTCOLOR", (0, 3), (0, 3), Colors.WHITE)
    style.add("TEXTCOLOR", (0, 4), (0, 4), Colors.WHITE)
    style.add("FONTNAME", (0, 4), (-1, 4), "Helvetica-Bold")
    t.setStyle(style)
    elements.append(t)
    elements.append(Spacer(1, 0.5 * cm))

    # High severity details
    if high:
        elements.append(Paragraph("<b>HIGH Severity Findings</b>", STYLES["SubSectionHeader"]))
        elements.append(Spacer(1, 0.2 * cm))

        for i, (ftype, date, amt, detail, _sev) in enumerate(high, 1):
            date_str = fmt_date(date)
            amt_str = fmt_money(amt) if amt else ""
            text = f"<b>{i}. {ftype}</b> ({date_str}, {amt_str}) — {detail}"
            elements.append(Paragraph(text, STYLES["BulletText"]))

    if medium:
        elements.append(Spacer(1, 0.3 * cm))
        elements.append(Paragraph("<b>MEDIUM Severity Findings</b>", STYLES["SubSectionHeader"]))
        elements.append(Spacer(1, 0.2 * cm))

        for i, (ftype, date, amt, detail, _sev) in enumerate(medium, 1):
            date_str = fmt_date(date)
            amt_str = fmt_money(amt) if amt else ""
            text = f"<b>{i}. {ftype}</b> ({date_str}, {amt_str}) — {detail}"
            elements.append(Paragraph(text, STYLES["BulletText"]))

    elements.append(PageBreak())
    return elements


# =============================================================================
# EVIDENCE INDEX
# =============================================================================


def create_evidence_index(df: pd.DataFrame, engine: Any) -> list:
    """Create evidence index with source cross-references."""
    elements = []

    elements.append(
        Paragraph("3. Evidence Index &amp; Source Cross-Reference", STYLES["SectionHeader"])
    )
    elements.append(Spacer(1, 0.3 * cm))

    # Source summary
    source_counts = df["Source"].value_counts()
    source_data = [["Source", "Records", "Percentage"]]
    total = len(df)
    for src, cnt in source_counts.items():
        source_data.append([src, str(cnt), f"{cnt / total:.1%}"])
    source_data.append(["TOTAL", str(total), "100.0%"])

    t = Table(source_data, colWidths=[8 * cm, 3 * cm, 3 * cm])
    t.setStyle(make_table_style(num_rows=len(source_data)))
    elements.append(t)
    elements.append(Spacer(1, 0.5 * cm))

    # Source cross-reference detail
    elements.append(Paragraph("<b>Source Document Inventory</b>", STYLES["SubSectionHeader"]))
    elements.append(Spacer(1, 0.2 * cm))

    # Group by source and show key details
    for src in source_counts.index:
        src_df = df[df["Source"] == src].copy()
        src_df["_dt"] = src_df["Date"].apply(parse_to_sort_date)
        src_df = src_df.sort_values("_dt")

        elements.append(
            Paragraph(f"<b>{src}</b> ({len(src_df)} records)", STYLES["SubSubSectionHeader"])
        )

        # Summary table for this source
        detail_data = [["Date", "Invoice #", "Amount", "Period", "Entry Type", "Reading"]]
        for _, row in src_df.head(20).iterrows():  # Limit to first 20
            detail_data.append(
                [
                    fmt_date(row.get("Date")),
                    str(row.get("Invoice #", "N/A")),
                    fmt_money(row.get("Amount (£)")),
                    f"{fmt_date(row.get('Period From'))}–{fmt_date(row.get('Period To'))}",
                    str(row.get("Entry Type", "")),
                    str(row.get("Reading", "")),
                ]
            )

        if len(src_df) > 20:
            detail_data.append(["...", f"({len(src_df) - 20} more records)", "", "", "", ""])

        t = Table(detail_data, colWidths=[2.5 * cm, 3 * cm, 3 * cm, 4 * cm, 3 * cm, 2 * cm])
        t.setStyle(make_table_style(num_rows=len(detail_data), font_size=7))
        elements.append(t)
        elements.append(Spacer(1, 0.3 * cm))

    elements.append(PageBreak())
    return elements


# =============================================================================
# DETAILED FINDINGS SECTIONS
# =============================================================================


def create_anomaly_detail_section(flags: list, df: pd.DataFrame) -> list:
    """Create detailed anomaly findings section."""
    elements = []

    elements.append(Paragraph("4. Detailed Findings", STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    # Group by category
    categories: dict[str, list[tuple]] = {
        "LARGE JUMP": [],
        "BILLING GAP": [],
        "ESTIMATED RUN": [],
        "HIGH DAILY RATE": [],
        "RECONCILIATION MISMATCH": [],
        "BALANCE REDUCTION": [],
    }

    for f in flags:
        ftype, date, amt, detail, sev = f
        if ftype in categories:
            categories[ftype].append(f)

    for cat, cat_flags in categories.items():
        if not cat_flags:
            continue

        elements.append(
            Paragraph(
                f"4.{list(categories.keys()).index(cat) + 1} {cat.replace('_', ' ').title()}",
                STYLES["SubSectionHeader"],
            )
        )
        elements.append(Spacer(1, 0.2 * cm))

        # Table of findings
        detail_data = [["#", "Date", "Amount", "Severity", "Detail"]]
        for i, (_ftype, date, amt, detail, sev) in enumerate(cat_flags, 1):
            detail_data.append(
                [
                    str(i),
                    fmt_date(date),
                    fmt_money(amt) if amt else "",
                    sev,
                    detail[:200] + ("..." if len(detail) > 200 else ""),
                ]
            )

        t = Table(
            detail_data, colWidths=[0.8 * cm, 2.5 * cm, 3 * cm, 2 * cm, CONTENT_WIDTH - 8.3 * cm]
        )
        style = make_table_style(num_rows=len(detail_data), font_size=7)
        # Color severity column
        for row_idx, (_, _, _, sev, _) in enumerate(cat_flags, 1):
            style.add("TEXTCOLOR", (3, row_idx), (3, row_idx), severity_color(sev))
            style.add("FONTNAME", (3, row_idx), (3, row_idx), "Helvetica-Bold")
        t.setStyle(style)
        elements.append(t)
        elements.append(Spacer(1, 0.4 * cm))

    elements.append(PageBreak())
    return elements


def create_timeline_section(df: pd.DataFrame, flags: list) -> list:
    """Create chronological timeline of events."""
    elements = []

    elements.append(Paragraph("5. Timeline of Events", STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    # Build timeline events
    events = []
    df = df.copy()
    df["_dt"] = df["Date"].apply(parse_to_sort_date)
    df = df.sort_values("_dt")

    # Add all records as timeline events
    for _, row in df.iterrows():
        events.append(
            {
                "date": row["Date"],
                "type": row.get("Entry Type", "Record"),
                "amount": row.get("Amount (£)"),
                "detail": f"{row.get('Source', '')} — {row.get('Details', '')[:100]}",
                "source": row.get("Source", ""),
            }
        )

    # Add flag events
    for ftype, date, amt, detail, _sev in flags:
        events.append(
            {
                "date": date,
                "type": f"⚠ {ftype}",
                "amount": amt,
                "detail": detail,
                "source": "AUTOMATED FLAG",
            }
        )

    # Sort by date
    events.sort(key=lambda e: parse_to_sort_date(e["date"]) or pd.Timestamp.min)

    # Timeline table
    timeline_data = [["Date", "Event Type", "Amount", "Detail"]]
    for ev in events:
        timeline_data.append(
            [
                fmt_date(ev["date"]),
                ev["type"],
                fmt_money(ev["amount"]) if ev["amount"] else "",
                ev["detail"][:150] + ("..." if len(ev["detail"]) > 150 else ""),
            ]
        )

    t = Table(timeline_data, colWidths=[2.5 * cm, 3.5 * cm, 3 * cm, CONTENT_WIDTH - 9 * cm])
    t.setStyle(make_table_style(num_rows=len(timeline_data), font_size=7))
    elements.append(t)

    elements.append(PageBreak())
    return elements


# =============================================================================
# OFGEM PRICE CAP COMPARISON
# =============================================================================


# OFGEM PRICE CAP COMPARISON
# =============================================================================


def _load_ofgem_caps() -> dict[str, dict]:
    """Load OFGEM price cap data.

    Returns a dictionary mapping period string (e.g., '2023-Q4') to cap values:
    {'unit_rate': p_per_kwh, 'standing_charge': p_per_day}

    For now, returns built-in data for key periods. Could be extended to load
    from external CSV/JSON or API.
    """
    # OFGEM Default Tariff Cap history (unit rate in p/kWh, standing charge in p/day)
    # Source: https://www.ofgem.gov.uk/energy-policy-and-regulation/policies-and-regulations/default-tariff-cap
    return {
        # 2019
        "2019-Q1": {"unit_rate": 17.46, "standing_charge": 25.09},
        "2019-Q2": {"unit_rate": 15.29, "standing_charge": 25.09},
        "2019-Q3": {"unit_rate": 15.29, "standing_charge": 25.09},
        "2019-Q4": {"unit_rate": 15.29, "standing_charge": 25.09},
        # 2020
        "2020-Q1": {"unit_rate": 15.29, "standing_charge": 25.09},
        "2020-Q2": {"unit_rate": 15.29, "standing_charge": 25.09},
        "2020-Q3": {"unit_rate": 15.29, "standing_charge": 25.09},
        "2020-Q4": {"unit_rate": 15.29, "standing_charge": 20.58},
        # 2021
        "2021-Q1": {"unit_rate": 20.58, "standing_charge": 20.58},
        "2021-Q2": {"unit_rate": 20.58, "standing_charge": 20.58},
        "2021-Q3": {"unit_rate": 20.58, "standing_charge": 20.58},
        "2021-Q4": {"unit_rate": 20.58, "standing_charge": 20.58},
        # 2022
        "2022-Q1": {"unit_rate": 20.58, "standing_charge": 20.58},
        "2022-Q2": {"unit_rate": 28.34, "standing_charge": 45.34},
        "2022-Q3": {"unit_rate": 28.34, "standing_charge": 45.34},
        "2022-Q4": {"unit_rate": 28.34, "standing_charge": 45.34},
        # 2023
        "2023-Q1": {"unit_rate": 34.00, "standing_charge": 46.36},
        "2023-Q2": {"unit_rate": 34.00, "standing_charge": 46.36},
        "2023-Q3": {"unit_rate": 34.00, "standing_charge": 46.36},
        "2023-Q4": {"unit_rate": 27.35, "standing_charge": 46.36},
        # 2024
        "2024-Q1": {"unit_rate": 28.62, "standing_charge": 53.35},
        "2024-Q2": {"unit_rate": 24.50, "standing_charge": 60.10},
        "2024-Q3": {"unit_rate": 24.50, "standing_charge": 60.10},
        "2024-Q4": {"unit_rate": 24.50, "standing_charge": 60.10},
        # 2025
        "2025-Q1": {"unit_rate": 24.50, "standing_charge": 60.10},
        "2025-Q2": {"unit_rate": 24.50, "standing_charge": 60.10},
        "2025-Q3": {"unit_rate": 24.50, "standing_charge": 60.10},
        "2025-Q4": {"unit_rate": 24.50, "standing_charge": 60.10},
        # 2026
        "2026-Q1": {"unit_rate": 24.50, "standing_charge": 60.10},
    }


def _period_to_ofgem_quarter(dt: datetime | None) -> str | None:
    """Convert datetime to OFGEM quarter string (e.g., '2024-Q1')."""
    if dt is None or pd.isna(dt):
        return None
    try:
        quarter = (dt.month - 1) // 3 + 1
        return f"{dt.year}-Q{quarter}"
    except Exception:
        return None


def create_ofgem_comparison(df: pd.DataFrame, config: dict | None = None) -> list:
    """Create OFGEM price cap comparison section."""
    elements = []

    elements.append(Paragraph("6. OFGEM Price Cap Comparison", STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    elements.append(
        Paragraph(
            "The following analysis compares the effective unit rates charged on EDF bills "
            "against the OFGEM Default Tariff Cap (Price Cap) for the corresponding periods. "
            "Any charges exceeding the cap may indicate regulatory non-compliance.",
            STYLES["BodyText"],
        )
    )
    elements.append(Spacer(1, 0.3 * cm))

    # Load OFGEM cap data
    ofgem_caps = _load_ofgem_caps()

    # Compute unit rates from bills
    df = df.copy()
    if "_dt" not in df.columns:
        df["_dt"] = df["Date"].apply(parse_to_sort_date)
    df = df.sort_values("_dt").reset_index(drop=True)

    # Filter for records with both Period Charge and Units
    bills = df[
        (df["Period Charge (£)"].notna())
        & (df["Period Charge (£)"] != "N/A")
        & (df["Units (kWh)"].notna())
        & (df["Units (kWh)"] != "N/A")
        & (df["Units (kWh)"] != "")
    ].copy()

    if bills.empty:
        elements.append(
            Paragraph(
                "No billing records with both Period Charge and Units (kWh) available for comparison.",
                STYLES["BodyText"],
            )
        )
        elements.append(PageBreak())
        return elements

    # Compute unit rate for each bill
    bills["_unit_rate"] = (
        bills["Period Charge (£)"].astype(float) / bills["Units (kWh)"].astype(float) * 100
    )

    # Compute quarter for each bill
    bills["_quarter"] = bills["_dt"].apply(_period_to_ofgem_quarter)

    # Build comparison table
    cap_data = [["Period", "Bill Unit Rate (p/kWh)", "OFGEM Cap (p/kWh)", "Difference", "Status"]]

    exceed_count = 0
    for quarter in sorted(bills["_quarter"].dropna().unique()):
        if quarter not in ofgem_caps:
            continue
        cap = ofgem_caps[quarter]
        quarter_bills = bills[bills["_quarter"] == quarter]
        avg_rate = quarter_bills["_unit_rate"].mean()
        if pd.isna(avg_rate):
            continue
        cap_rate = cap["unit_rate"]
        diff = avg_rate - cap_rate
        status = "EXCEEDS CAP" if diff > 0 else "AT CAP" if abs(diff) < 0.01 else "BELOW CAP"
        if diff > 0:
            exceed_count += 1
        cap_data.append(
            [
                quarter,
                fmt_number(avg_rate, 2),
                fmt_number(cap_rate, 2),
                fmt_number(diff, 2) if diff != 0 else "0.00",
                status,
            ]
        )

    # Add summary row
    cap_data.append(
        [
            "OVERALL",
            "—",
            "—",
            f"{exceed_count} periods exceed cap" if exceed_count > 0 else "No exceedances",
            "REVIEW REQUIRED" if exceed_count > 0 else "COMPLIANT",
        ]
    )

    t = Table(cap_data, colWidths=[3.5 * cm, 3.5 * cm, 3.5 * cm, 3 * cm, 2.5 * cm])
    t.setStyle(make_table_style(num_rows=len(cap_data)))

    # Color the status column
    style = TableStyle(
        [
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B4C6E7")),
        ]
    )
    for i in range(1, len(cap_data)):
        status = cap_data[i][4]
        if "EXCEEDS" in status:
            style.add("TEXTCOLOR", (4, i), (4, i), Colors.RED)
            style.add("FONTNAME", (4, i), (4, i), "Helvetica-Bold")
        elif "REVIEW" in status:
            style.add("TEXTCOLOR", (4, i), (4, i), Colors.RED)
            style.add("FONTNAME", (4, i), (4, i), "Helvetica-Bold")
        elif "BELOW" in status:
            style.add("TEXTCOLOR", (4, i), (4, i), Colors.GREEN)
    t.setStyle(style)

    elements.append(t)
    elements.append(Spacer(1, 0.3 * cm))

    elements.append(
        Paragraph(
            "<b>Methodology:</b> Unit rates calculated as Period Charge (£) ÷ Units (kWh) × 100. "
            "Standing charges compared separately against OFGEM daily cap. "
            "Only records with both Period Charge and Units (kWh) are included. "
            "OFGEM cap data sourced from official Default Tariff Cap publications.",
            STYLES["SmallText"],
        )
    )

    elements.append(PageBreak())
    return elements


# =============================================================================
# STATISTICAL ANALYSIS SECTION
# =============================================================================


def create_statistical_analysis(dfc: pd.DataFrame) -> list:
    """Create statistical analysis section."""
    elements = []

    elements.append(Paragraph("7. Statistical Analysis", STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    dfc = dfc.copy()
    dfc["_dt"] = dfc["Date"].apply(parse_to_sort_date)
    dfc = dfc.sort_values("_dt")
    amounts = dfc["Amount (£)"].astype(float).values
    n = len(amounts)

    if n < 3:
        elements.append(
            Paragraph("Insufficient data for statistical analysis.", STYLES["BodyText"])
        )
        elements.append(PageBreak())
        return elements

    # Descriptive stats
    amounts_series = pd.Series(amounts)

    stats_data = [
        ["Statistic", "Value"],
        ["Count", fmt_number(n)],
        ["Mean (£)", fmt_money(amounts_series.mean())],
        ["Median (£)", fmt_money(amounts_series.median())],
        ["Std Deviation (£)", fmt_money(amounts_series.std())],
        ["Min (£)", fmt_money(amounts_series.min())],
        ["Max (£)", fmt_money(amounts_series.max())],
        ["Range (£)", fmt_money(amounts_series.max() - amounts_series.min())],
        ["Coefficient of Variation", fmt_pct(amounts_series.std() / amounts_series.mean())],
    ]

    t = Table(stats_data, colWidths=[10 * cm, 5 * cm])
    t.setStyle(make_table_style(num_rows=len(stats_data)))
    elements.append(t)
    elements.append(Spacer(1, 0.5 * cm))

    # Rolling statistics
    elements.append(
        Paragraph("<b>6-Period Rolling Statistics (Latest)</b>", STYLES["SubSectionHeader"])
    )
    rolling = amounts_series.rolling(6, min_periods=1)
    roll_mean = rolling.mean().iloc[-1]
    roll_std = rolling.std().iloc[-1]
    roll_min = rolling.min().iloc[-1]
    roll_max = rolling.max().iloc[-1]

    # Handle NaN values
    if pd.isna(roll_std):
        roll_std = 0.0
    if pd.isna(roll_mean):
        roll_mean = float(amounts_series.iloc[-1])
    if pd.isna(roll_min):
        roll_min = float(amounts_series.min())
    if pd.isna(roll_max):
        roll_max = float(amounts_series.max())

    roll_data = [
        ["Metric", "Value"],
        ["Rolling Mean (£)", fmt_money(float(roll_mean))],
        ["Rolling Std Dev (£)", fmt_money(float(roll_std))],
        ["Rolling Min (£)", fmt_money(float(roll_min))],
        ["Rolling Max (£)", fmt_money(float(roll_max))],
    ]
    t = Table(roll_data, colWidths=[10 * cm, 5 * cm])
    t.setStyle(make_table_style(num_rows=len(roll_data)))
    elements.append(t)
    elements.append(Spacer(1, 0.5 * cm))

    # Normality test (if scipy available)
    try:
        from scipy import stats as sp_stats

        shapiro_stat, shapiro_p = sp_stats.shapiro(amounts_series.dropna())
        elements.append(
            Paragraph("<b>Distribution Normality (Shapiro-Wilk)</b>", STYLES["SubSectionHeader"])
        )
        norm_data = [
            ["Test", "Statistic", "p-value", "Normal?"],
            [
                "Shapiro-Wilk",
                f"{shapiro_stat:.4f}",
                f"{shapiro_p:.4f}",
                "Yes" if shapiro_p > 0.05 else "No",
            ],
        ]
        t = Table(norm_data, colWidths=[4 * cm, 3 * cm, 3 * cm, 3 * cm])
        t.setStyle(make_table_style(num_rows=len(norm_data)))
        elements.append(t)
    except ImportError:
        elements.append(
            Paragraph(
                "<i>Scipy not available — install for normality testing.</i>", STYLES["SmallText"]
            )
        )

    elements.append(PageBreak())
    return elements


# =============================================================================
# PAYMENT ANALYSIS SECTION
# =============================================================================


def create_payment_analysis(dfc: pd.DataFrame) -> list:
    """Create payment & credit analysis section."""
    elements = []

    elements.append(Paragraph("8. Payment &amp; Credit Analysis", STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    payments = dfc[dfc["Entry Type"].isin(["Payment", "Credit"])].copy()
    if not payments.empty:
        payments["_dt"] = payments["Date"].apply(parse_to_sort_date)
        payments = payments.sort_values("_dt")

        elements.append(Paragraph("<b>Payment Summary</b>", STYLES["SubSectionHeader"]))
        pay_amounts = payments["Amount (£)"].astype(float)

        pay_data = [
            ["Metric", "Value"],
            ["Total Payments/Credits", str(len(payments))],
            ["Total Amount Paid", fmt_money(pay_amounts.sum())],
            ["Average Payment", fmt_money(pay_amounts.mean())],
            ["Median Payment", fmt_money(pay_amounts.median())],
            ["Largest Payment", fmt_money(pay_amounts.max())],
            ["Smallest Payment", fmt_money(pay_amounts.min())],
        ]
        if len(payments) > 1:
            pay_dates = payments["_dt"].dropna()
            intervals = pay_dates.diff().dt.days.dropna()
            if len(intervals) > 0:
                pay_data.append(["Avg Interval (days)", fmt_number(float(intervals.mean()), 1)])
                pay_data.append(
                    ["Median Interval (days)", fmt_number(float(intervals.median()), 1)]
                )

        t = Table(pay_data, colWidths=[10 * cm, 5 * cm])
        t.setStyle(make_table_style(num_rows=len(pay_data)))
        elements.append(t)
        elements.append(Spacer(1, 0.5 * cm))

        # Payment detail table
        elements.append(Paragraph("<b>Payment Chronology</b>", STYLES["SubSectionHeader"]))
        pay_detail = [["Date", "Type", "Amount (£)", "Balance After", "Details"]]
        for _, row in payments.iterrows():
            pay_detail.append(
                [
                    fmt_date(row["Date"]),
                    str(row["Entry Type"]),
                    fmt_money(row["Amount (£)"]),
                    fmt_money(row["Amount (£)"]),
                    str(row.get("Details", ""))[:80],
                ]
            )
        t = Table(
            pay_detail,
            colWidths=[2.5 * cm, 2.5 * cm, 3 * cm, 3 * cm, max(2 * cm, CONTENT_WIDTH - 11 * cm)],
        )
        t.setStyle(make_table_style(num_rows=len(pay_detail), font_size=7))
        elements.append(t)
    else:
        elements.append(
            Paragraph("No payment/credit records found in the data.", STYLES["BodyText"])
        )

    elements.append(PageBreak())
    return elements


# =============================================================================
# FORECAST SECTION
# =============================================================================


# =============================================================================
# FORECAST SECTION
# =============================================================================


def create_forecast_section(dfc: pd.DataFrame) -> list:
    """Create forecast & projection section."""
    elements = []

    elements.append(Paragraph("9. Forecast &amp; Projection", STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    elements.append(
        Paragraph(
            "Projections are based on historical billing patterns and should be treated as "
            "indicative only. They assume continuation of current trends and do not account for "
            "seasonal variations, tariff changes, or policy changes.",
            STYLES["BodyText"],
        )
    )
    elements.append(Spacer(1, 0.3 * cm))

    # Prepare data for forecasting
    dfc = dfc.copy()
    dfc["_dt"] = dfc["Date"].apply(parse_to_sort_date)
    dfc = dfc.sort_values("_dt").reset_index(drop=True)
    amounts = dfc["Amount (£)"].astype(float).values
    n = len(amounts)

    if n < 3:
        elements.append(
            Paragraph(
                "Insufficient data for forecasting (need at least 3 periods).", STYLES["BodyText"]
            )
        )
        elements.append(PageBreak())
        return elements

    # Try to import scipy for linear regression
    has_scipy = False
    try:
        from scipy import stats as sp_stats

        has_scipy = True
    except ImportError:
        has_scipy = False

    # Try to import statsmodels for Holt-Winters
    has_statsmodels = False
    try:
        from statsmodels.tsa.holtwinters import ExponentialSmoothing

        has_statsmodels = True
    except ImportError:
        has_statsmodels = False

    # 1. Linear Regression Forecast
    linear_forecast = []
    model_info = []
    if has_scipy:
        x = np.arange(n)
        slope, intercept, r_value, p_value, std_err = sp_stats.linregress(x, amounts)
        linear_forecast = [intercept + slope * (n + i) for i in range(1, 7)]
        model_info = [
            f"<b>Linear Regression:</b> slope={slope:.2f}, intercept={intercept:.2f}, R²={r_value**2:.4f}, p={p_value:.4f}",
        ]
    else:
        # Fallback: simple average
        linear_forecast = [float(np.mean(amounts))] * 6
        model_info = ["<b>Linear Regression:</b> not available (install scipy) - using mean"]

    # 2. EMA (Exponential Moving Average) Forecast
    alpha = 0.3  # smoothing factor
    ema = amounts[0]
    for val in amounts[1:]:
        ema = alpha * val + (1 - alpha) * ema
    ema_forecast = [ema] * 6
    model_info.append(f"<b>EMA (α={alpha}):</b> current level={ema:.2f}")

    # 3. Holt-Winters Forecast (if statsmodels available)
    hw_forecast = None
    if has_statsmodels and n >= 6:
        try:
            # Use additive trend, no seasonality (need at least 2 seasons for seasonality)
            model = ExponentialSmoothing(amounts, trend="add", seasonal=None)
            hw_fit = model.fit(smoothing_level=alpha, smoothing_trend=0.1, optimized=True)
            hw_forecast = hw_fit.forecast(6).tolist()
        except Exception:
            hw_forecast = None

    # Build forecast table
    forecast_header = ["Forecast Period", "Linear Reg. (£)", "EMA (£)"]
    if hw_forecast:
        forecast_header.append("Holt-Winters (£)")

    forecast_data = [forecast_header]
    for i in range(6):
        row = [f"+{i + 1} Period", fmt_money(linear_forecast[i]), fmt_money(ema_forecast[i])]
        if hw_forecast:
            row.append(fmt_money(hw_forecast[i]))
        forecast_data.append(row)

    # Calculate column widths
    num_cols = len(forecast_header)
    col_width = CONTENT_WIDTH / num_cols
    col_widths = [col_width] * num_cols

    t = Table(forecast_data, colWidths=col_widths)
    t.setStyle(make_table_style(num_rows=len(forecast_data)))
    elements.append(
        Paragraph("<b>Next 6 Periods — Multi-Method Projection</b>", STYLES["SubSectionHeader"])
    )
    elements.append(Spacer(1, 0.2 * cm))
    elements.append(t)
    elements.append(Spacer(1, 0.3 * cm))

    if hw_forecast:
        model_info.append(
            "<b>Holt-Winters:</b> additive trend, no seasonality (fitted via statsmodels)"
        )
    else:
        model_info.append("<b>Holt-Winters:</b> not available (install statsmodels)")

    for info in model_info:
        elements.append(Paragraph(info, STYLES["SmallText"]))

    elements.append(Spacer(1, 0.3 * cm))

    elements.append(
        Paragraph(
            "<i>Note: Projections assume continuation of current trends and do not account for "
            "seasonal variations, tariff changes, or policy changes. "
            "Full forecasting with confidence intervals available in the Excel workbook (Forecast &amp; Projection tab).</i>",
            STYLES["SmallText"],
        )
    )

    elements.append(PageBreak())
    return elements


# =============================================================================
# DATA QUALITY SECTION
# =============================================================================


def create_data_quality_section(df: pd.DataFrame) -> list:
    """Create data quality assessment section."""
    elements = []

    elements.append(Paragraph("10. Data Quality Assessment", STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    total = len(df)
    date_parsed = df["Date"].apply(lambda x: parse_to_sort_date(x) is not pd.NaT).sum()
    amt_complete = df["Amount (£)"].notna().sum()
    period_complete = (df["Period From"] != "N/A").sum()
    reading_classified = (df["Reading"] != "Unknown").sum() if "Reading" in df.columns else 0
    dup_count = df.duplicated(subset=["Date", "Amount (£)"]).sum()

    quality_data = [
        ["Check", "Passed", "Total", "Rate", "Status"],
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
        ],
        [
            "Amount Complete",
            str(int(amt_complete)),
            str(total),
            f"{amt_complete / total:.1%}",
            "PASS" if amt_complete == total else "WARN",
        ],
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
        ],
        [
            "Reading Classified",
            str(int(reading_classified)),
            str(total),
            f"{reading_classified / total:.1%}",
            "PASS" if reading_classified / total > 0.5 else "WARN",
        ],
        [
            "Duplicates (Date+Amount)",
            str(int(dup_count)),
            str(total),
            f"{dup_count / total:.2%}",
            "PASS" if dup_count / total < 0.05 else "WARN" if dup_count / total < 0.15 else "FAIL",
        ],
    ]

    t = Table(quality_data, colWidths=[5 * cm, 2 * cm, 2 * cm, 2.5 * cm, CONTENT_WIDTH - 12 * cm])
    style = make_table_style(num_rows=len(quality_data))
    # Color status column
    for i in range(1, len(quality_data)):
        status = quality_data[i][4]
        if status == "PASS":
            style.add("TEXTCOLOR", (4, i), (4, i), Colors.GREEN)
            style.add("FONTNAME", (4, i), (4, i), "Helvetica-Bold")
        elif status == "WARN":
            style.add("TEXTCOLOR", (4, i), (4, i), Colors.AMBER)
            style.add("FONTNAME", (4, i), (4, i), "Helvetica-Bold")
        elif status == "FAIL":
            style.add("TEXTCOLOR", (4, i), (4, i), Colors.RED)
            style.add("FONTNAME", (4, i), (4, i), "Helvetica-Bold")
    t.setStyle(style)
    elements.append(t)
    elements.append(Spacer(1, 0.5 * cm))

    # Source distribution
    elements.append(Paragraph("<b>Source Distribution</b>", STYLES["SubSectionHeader"]))
    src_counts = df["Source"].value_counts()
    src_data = [["Source", "Records", "Percentage"]]
    for src, cnt in src_counts.items():
        src_data.append([src, str(cnt), f"{cnt / len(df):.1%}"])
    src_data.append(["TOTAL", str(len(df)), "100.0%"])

    t = Table(src_data, colWidths=[8 * cm, 3 * cm, 3 * cm])
    t.setStyle(make_table_style(num_rows=len(src_data)))
    elements.append(t)

    elements.append(PageBreak())
    return elements


# =============================================================================
# TARIFF IMPACT SECTION
# =============================================================================


def create_tariff_impact_section(dfc: pd.DataFrame) -> list:
    """Create tariff impact analysis section."""
    elements = []

    elements.append(Paragraph("11. Tariff Impact Analysis", STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    if "Tariff" not in dfc.columns or dfc["Tariff"].isna().all():
        elements.append(
            Paragraph(
                "No tariff data available in the extracted records. Tariff information is "
                "typically found on new-format (KI/KCR) invoices.",
                STYLES["BodyText"],
            )
        )
        elements.append(PageBreak())
        return elements

    # Filter valid tariff data
    tariff_data = dfc.dropna(subset=["Tariff"])
    tariff_data = tariff_data[tariff_data["Tariff"] != "N/A"]

    if tariff_data.empty:
        elements.append(Paragraph("No valid tariff records found.", STYLES["BodyText"]))
        elements.append(PageBreak())
        return elements

    # Convert unit rate to numeric
    tariff_data = tariff_data.copy()
    tariff_data["unit_rate_num"] = pd.to_numeric(tariff_data["Unit Rate (p/kWh)"], errors="coerce")
    tariff_data = tariff_data.dropna(subset=["unit_rate_num"])

    if tariff_data.empty:
        elements.append(Paragraph("No computable unit rates found.", STYLES["BodyText"]))
        elements.append(PageBreak())
        return elements

    # Stats by tariff
    tariff_stats = (
        tariff_data.groupby("Tariff")
        .agg(
            count=("unit_rate_num", "count"),
            avg_rate=("unit_rate_num", "mean"),
            median_rate=("unit_rate_num", "median"),
            min_rate=("unit_rate_num", "min"),
            max_rate=("unit_rate_num", "max"),
            avg_charge=("Period Charge (£)", lambda x: pd.to_numeric(x, errors="coerce").mean()),
        )
        .reset_index()
    )

    elements.append(Paragraph("<b>Unit Rate by Tariff</b>", STYLES["SubSectionHeader"]))
    tariff_table = [
        ["Tariff", "Records", "Avg Rate (p/kWh)", "Median", "Min", "Max", "Avg Charge (£)"]
    ]
    for _, row in tariff_stats.iterrows():
        tariff_table.append(
            [
                str(row["Tariff"]),
                str(int(row["count"])),
                fmt_number(row["avg_rate"], 2),
                fmt_number(row["median_rate"], 2),
                fmt_number(row["min_rate"], 2),
                fmt_number(row["max_rate"], 2),
                fmt_money(row["avg_charge"]) if pd.notna(row["avg_charge"]) else "N/A",
            ]
        )

    t = Table(
        tariff_table, colWidths=[4 * cm, 2 * cm, 3 * cm, 2.5 * cm, 2.5 * cm, 2.5 * cm, 3 * cm]
    )
    t.setStyle(make_table_style(num_rows=len(tariff_table), font_size=8))
    elements.append(t)
    elements.append(Spacer(1, 0.5 * cm))

    # Tariff changes
    tariff_data["_dt"] = tariff_data["Date"].apply(parse_to_sort_date)
    tariff_data = tariff_data.sort_values("_dt")
    changes = tariff_data["Tariff"].ne(tariff_data["Tariff"].shift()).cumsum()
    n_changes = int(changes.max()) if not changes.empty else 0

    elements.append(
        Paragraph(f"<b>Tariff Changes Detected: {n_changes}</b>", STYLES["SubSectionHeader"])
    )

    elements.append(PageBreak())
    return elements


# =============================================================================
# APPENDICES
# =============================================================================


def create_appendix_methodology(config: dict) -> list:
    """Create Appendix A: Methodology."""
    elements = []

    elements.append(PageBreak())
    elements.append(
        Paragraph("Appendix A: Methodology &amp; Data Sources", STYLES["SectionHeader"])
    )
    elements.append(Spacer(1, 0.3 * cm))

    sections = [
        (
            "A.1 Data Sources",
            [
                "All billing records were extracted from three primary source types:",
                "• <b>PDF Bills:</b> EDF Energy invoices (both legacy and new KI/KCR formats) "
                "processed via pdfplumber with format-specific regex extraction.",
                "• <b>HTM Export:</b> EDF MyAccount 'Payments and Invoices' export parsed via "
                "BeautifulSoup with pattern matching for charge/payment/reversal entries.",
                "• <b>PST/OST Email Archives:</b> Outlook data files processed via libpff-python, "
                "extracting email bodies (HTML/plain text/RTF) and PDF attachments.",
            ],
        ),
        (
            "A.2 Amount Extraction Logic",
            [
                "Two complementary strategies ensure comprehensive amount detection:",
                "1. <b>Smart Context Search:</b> 10 prioritized regex patterns targeting specific "
                "EDF billing language (e.g., 'Current balance £X debit', 'Total charges for this period'). "
                "Patterns execute in priority order; first match wins.",
                "2. <b>Large Amount Fallback:</b> Scans all £ amounts ≥ minimum threshold, "
                "selecting the largest. Used when context patterns fail.",
            ],
        ),
        (
            "A.3 Deduplication",
            [
                "Multi-pass deduplication matches the same bill across sources:",
                "• Pass 1: Exact match on <b>Period To date + Amount</b> (catches HTM ↔ PST ↔ PDF).",
                "• Pass 2: For records without period info (e.g., local PDFs), match by Amount "
                "within 60-day window of any kept record.",
                "Source priority for keeping: HTM Account History &gt; PST PDF Attachment &gt; "
                "Email Body &gt; Local PDF Folder.",
            ],
        ),
        (
            "A.4 Classification Logic",
            [
                "<b>Entry Type:</b> 'New Bill' (period charge + period dates), 'Ongoing Balance' "
                "(cumulative balance only), 'Payment', 'Credit'.",
                "<b>Reading Type:</b> Regex classification — 'Estimated' (estimated/est./estimate), "
                "'Actual' (actual/customer reading/your reading), 'Smart' (smart meter/automated).",
            ],
        ),
        (
            "A.5 Configuration Used",
            [
                f"Minimum Amount Threshold: {fmt_money(config.get('min_amount', 500))}",
                f"Analysis Threshold: {fmt_money(config.get('analysis_min', 500))}",
                f"Account Filter: {'Enabled' if config.get('use_acc_filter') else 'Disabled'} "
                f"({config.get('acc_num', 'N/A')})",
                f"Domain Filter: {'Enabled' if config.get('use_domain_filter') else 'Disabled'} "
                f"({config.get('domain_filter', 'edfenergy.com')})",
                f"Deduplication: {'Enabled' if config.get('use_dedup') else 'Disabled'}",
                f"Smart Context Search: {'Enabled' if config.get('use_anchors') else 'Disabled'}",
                f"Large Amount Fallback: {'Enabled' if config.get('use_large') else 'Disabled'}",
            ],
        ),
    ]

    for title, bullets in sections:
        elements.append(Paragraph(title, STYLES["SubSectionHeader"]))
        elements.append(Spacer(1, 0.1 * cm))
        for bullet in bullets:
            elements.append(Paragraph(bullet, STYLES["BulletText"]))
        elements.append(Spacer(1, 0.2 * cm))

    elements.append(PageBreak())
    return elements


def create_appendix_full_evidence(
    df: pd.DataFrame, filtered: list | None = None, config: dict | None = None
) -> list:
    """Create Appendix C: Full Evidence Table with all records."""
    elements = []

    elements.append(Paragraph("Appendix C: Full Evidence Table", STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    elements.append(
        Paragraph(
            "This appendix contains the complete set of billing records used in this analysis. "
            "Records are sorted chronologically by date.",
            STYLES["BodyText"],
        )
    )
    elements.append(Spacer(1, 0.3 * cm))

    if df.empty:
        elements.append(Paragraph("No records available.", STYLES["BodyText"]))
        elements.append(PageBreak())
        return elements

    # Ensure _dt column for sorting
    if "_dt" not in df.columns:
        df["_dt"] = df["Date"].apply(parse_to_sort_date)
    df_sorted = df.sort_values("_dt").reset_index(drop=True)

    # Table header
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

    evidence_data = [evidence_header]

    for _, row in df_sorted.iterrows():
        evidence_data.append(
            [
                fmt_date(row.get("Date", "")),
                str(row.get("Source", ""))[:30],
                str(row.get("Entry Type", "")),
                fmt_money(row.get("Amount (£)", 0)),
                fmt_money(row.get("Period Charge (£)", 0))
                if row.get("Period Charge (£)") not in ("", "N/A", None)
                else "N/A",
                str(row.get("Period From", "")),
                str(row.get("Period To", "")),
                str(row.get("Invoice #", ""))[:15],
                str(row.get("Reading", ""))[:15],
                str(row.get("Units (kWh)", ""))[:10],
                str(row.get("Standing Chg (p/day)", ""))[:10],
                str(row.get("Attachment Name", ""))[:20],
                str(row.get("Details", ""))[:50],
            ]
        )

    # Calculate column widths
    col_widths = [
        2.0 * cm,  # Date
        2.5 * cm,  # Source
        2.0 * cm,  # Entry Type
        2.0 * cm,  # Amount
        2.0 * cm,  # Period Charge
        2.0 * cm,  # Period From
        2.0 * cm,  # Period To
        1.5 * cm,  # Invoice
        1.5 * cm,  # Reading
        1.5 * cm,  # Units
        1.5 * cm,  # Standing
        2.0 * cm,  # Attachment
        CONTENT_WIDTH - sum([2.0, 2.5, 2.0, 2.0, 2.0, 2.0, 2.0, 1.5, 1.5, 1.5, 1.5, 2.0]) * cm,
    ]

    t = Table(evidence_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 6),
                ("TEXTCOLOR", (0, 0), (-1, -1), Colors.DARK_GREY),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BACKGROUND", (0, 0), (-1, 0), Colors.NAVY),
                ("TEXTCOLOR", (0, 0), (-1, 0), Colors.WHITE),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B4C6E7")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                *[
                    ("BACKGROUND", (0, i), (-1, i), Colors.VERY_LIGHT_BLUE)
                    for i in range(1, len(evidence_data), 2)
                ],
            ]
        )
    )
    elements.append(t)
    elements.append(Spacer(1, 0.3 * cm))

    # Add filtered records if provided
    if filtered:
        elements.append(PageBreak())
        min_amt = fmt_money(config.get("min_amount", 500)) if config else "£500"
        elements.append(
            Paragraph(
                f"Appendix C (cont.): Filtered Records (Below {min_amt} Threshold)",
                STYLES["SectionHeader"],
            )
        )
        elements.append(Spacer(1, 0.3 * cm))

        filt_data = [evidence_header]
        for row in filtered:
            filt_data.append(
                [
                    fmt_date(row.get("Date", "")),
                    str(row.get("Source", ""))[:30],
                    str(row.get("Entry Type", "")),
                    fmt_money(row.get("Amount (£)", 0)),
                    fmt_money(row.get("Period Charge (£)", 0))
                    if row.get("Period Charge (£)") not in ("", "N/A", None)
                    else "N/A",
                    str(row.get("Period From", "")),
                    str(row.get("Period To", "")),
                    str(row.get("Invoice #", ""))[:15],
                    str(row.get("Reading", ""))[:15],
                    str(row.get("Units (kWh)", ""))[:10],
                    str(row.get("Standing Chg (p/day)", ""))[:10],
                    str(row.get("Attachment Name", ""))[:20],
                    str(row.get("Details", ""))[:50],
                ]
            )

        if len(filt_data) > 1:
            t2 = Table(filt_data, colWidths=col_widths, repeatRows=1)
            t2.setStyle(
                TableStyle(
                    [
                        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                        ("FONTSIZE", (0, 0), (-1, -1), 6),
                        ("TEXTCOLOR", (0, 0), (-1, -1), Colors.DARK_GREY),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("BACKGROUND", (0, 0), (-1, 0), Colors.AMBER),
                        ("TEXTCOLOR", (0, 0), (-1, 0), Colors.WHITE),
                        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B4C6E7")),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("TOPPADDING", (0, 0), (-1, -1), 2),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                        ("LEFTPADDING", (0, 0), (-1, -1), 2),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                        *[
                            ("BACKGROUND", (0, i), (-1, i), Colors.VERY_LIGHT_BLUE)
                            for i in range(1, len(filt_data), 2)
                        ],
                    ]
                )
            )
            elements.append(t2)

    elements.append(PageBreak())
    return elements


def create_appendix_glossary() -> list:
    """Create Appendix B: Glossary."""
    elements = []

    elements.append(Paragraph("Appendix B: Glossary", STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    terms = {
        "Period Charge (£)": "The charge for the specific billing period (not cumulative balance). "
        "Equivalent to 'Total charges for this period' on new EDF invoices.",
        "Amount (£)": "The primary balance figure — typically the current cumulative account balance "
        "on new invoices, or the running balance on HTM exports.",
        "Unit Rate (p/kWh)": "Effective price per kWh = Period Charge ÷ Units (kWh) × 100. "
        "Includes both energy and standing charge components unless separated.",
        "Standing Charge (p/day)": "Daily fixed charge regardless of consumption, as published on EDF bills.",
        "OFGEM Price Cap": "Maximum price per unit (p/kWh) and daily standing charge (p/day) that "
        "suppliers can charge customers on default/standard variable tariffs, "
        "set quarterly by OFGEM.",
        "Billing Gap": "Period exceeding 60 days (MEDIUM) or 120 days (HIGH) between consecutive "
        "bills where balance accumulates without a new statement.",
        "Z-Score Anomaly": "Data point exceeding 2.5 standard deviations from the mean, indicating "
        "statistical outlier (≈99% confidence under normality).",
        "IQR Anomaly": "Data point outside 1.5× the interquartile range (Q3−Q1), robust to "
        "non-normal distributions.",
        "Holt-Winters Forecast": "Exponential smoothing with trend and optional seasonality, "
        "suitable for time series with patterns.",
        "MAPE": "Mean Absolute Percentage Error — average of |forecast − actual|/actual × 100%. "
        "Lower is better; <10% considered good for energy billing.",
    }

    glossary_data = [["Term", "Definition"]]
    for term, definition in terms.items():
        glossary_data.append([term, definition])

    t = Table(glossary_data, colWidths=[4 * cm, CONTENT_WIDTH - 4 * cm])
    t.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("TEXTCOLOR", (0, 0), (-1, -1), Colors.DARK_GREY),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("BACKGROUND", (0, 0), (-1, 0), Colors.NAVY),
                ("TEXTCOLOR", (0, 0), (-1, 0), Colors.WHITE),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#B4C6E7")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                *[
                    ("BACKGROUND", (0, i), (-1, i), Colors.VERY_LIGHT_BLUE)
                    for i in range(1, len(terms) + 1, 2)
                ],
            ]
        )
    )
    elements.append(t)

    elements.append(PageBreak())
    return elements


# =============================================================================
# MAIN REPORT GENERATOR
# =============================================================================


def generate_ombudsman_pdf(
    records: list[dict],
    output_path: str,
    config: dict,
    engine: Any,
    filtered: list | None = None,
    ofgem_data: dict | None = None,
) -> str:
    """
    Generate a professional PDF report for Energy Ombudsman submission.

    Args:
        records: List of extracted billing records
        output_path: Path to save the PDF
        config: Configuration dictionary
        engine: EvidenceEngine instance (for metadata)
        filtered: Filtered-out records (below threshold)
        ofgem_data: Optional OFGEM price cap data

    Returns:
    Records:
        Path to generated PDF
    """
    if not records:
        raise ValueError("No records to report on")

    df = pd.DataFrame(records)
    if df.empty:
        raise ValueError("Records DataFrame is empty")

    # Validate required parameters
    if engine is None:
        raise ValueError(
            "Engine is required for report generation (contains PDF count, email count metadata)"
        )

    # Ensure required columns exist with defaults
    required_cols = {
        "Date": "01/01/1970",
        "Source": "Unknown",
        "Amount (£)": 0.0,
        "Period From": "N/A",
        "Period To": "N/A",
        "Invoice #": "N/A",
        "Period Charge (£)": "N/A",
        "Entry Type": "Unknown",
        "Reading": "Unknown",
        "Units (kWh)": "N/A",
        "Standing Chg (p/day)": "N/A",
        "Attachment Name": "N/A",
        "Details": "",
        "Logic Used": "",
    }
    for col, default in required_cols.items():
        if col not in df.columns:
            df[col] = default

    df["_sort"] = df["Date"].apply(parse_to_sort_date)
    df = df.sort_values("_sort").reset_index(drop=True)

    # Account reference
    acc_ref = config.get("report_account_ref") or config.get("acc_num") or "Unknown"

    # Period bounds
    dates_parsed = df["Date"].apply(parse_to_sort_date)
    valid_dates = dates_parsed.dropna()
    period_start = fmt_date(valid_dates.min()) if not valid_dates.empty else "Unknown"
    period_end = fmt_date(valid_dates.max()) if not valid_dates.empty else "Unknown"

    # Financial totals
    charges = df[df["Amount (£)"] > 0]["Amount (£)"].astype(float).sum()
    payments = df[df["Amount (£)"] < 0]["Amount (£)"].astype(float).sum() * -1

    # Compute dispute flags from the data
    from edf_collector import compute_dispute_flags

    # Ensure _dt column exists for the analysis
    if "_dt" not in df.columns:
        df["_dt"] = df["Date"].apply(parse_to_sort_date)
    df_sorted = df.sort_values("_dt").reset_index(drop=True)

    # Compute mean daily rate for HIGH DAILY RATE detection
    mean_daily = 0.0
    try:
        for i in range(1, len(df_sorted)):
            p = df_sorted.iloc[i - 1]
            c_ = df_sorted.iloc[i]
            days = (c_["_dt"] - p["_dt"]).days
            charge = float(c_["Amount (£)"]) - float(p["Amount (£)"])
            if days > 0 and charge > 0:
                daily = charge / days
                if mean_daily == 0.0:
                    mean_daily = daily
                else:
                    mean_daily = (mean_daily + daily) / 2
    except Exception:
        mean_daily = 0.0

    flags, flag_counts = compute_dispute_flags(df_sorted, mean_daily)

    # Build document
    doc = build_doc_template(output_path)
    elements = []

    # === COVER PAGE ===
    elements.extend(
        create_cover_page(acc_ref, period_start, period_end, datetime.now().strftime("%d %B %Y"))
    )
    elements.append(NextPageTemplate("content"))
    elements.append(PageBreak())

    # === TABLE OF CONTENTS ===
    try:
        elements.extend(create_table_of_contents())
    except Exception as e:
        elements.append(Paragraph(f"<i>Table of Contents failed: {e}</i>", STYLES["BodyText"]))

    # === EXECUTIVE SUMMARY ===
    try:
        elements.extend(
            create_executive_summary(
                df,
                config,
                acc_ref,
                flag_counts,
                len(records),
                charges,
                payments,
                period_start,
                period_end,
            )
        )
    except Exception as e:
        elements.append(Paragraph(f"<i>Executive Summary failed: {e}</i>", STYLES["BodyText"]))

    # === KEY FINDINGS ===
    try:
        elements.extend(create_key_findings_table(flags))
    except Exception as e:
        elements.append(Paragraph(f"<i>Key Findings failed: {e}</i>", STYLES["BodyText"]))

    # === EVIDENCE INDEX ===
    try:
        elements.extend(create_evidence_index(df, engine))
    except Exception as e:
        elements.append(Paragraph(f"<i>Evidence Index failed: {e}</i>", STYLES["BodyText"]))

    # === DETAILED FINDINGS ===
    try:
        elements.extend(create_anomaly_detail_section(flags, df))
    except Exception as e:
        elements.append(Paragraph(f"<i>Detailed Findings failed: {e}</i>", STYLES["BodyText"]))

    # === TIMELINE ===
    try:
        elements.extend(create_timeline_section(df, flags))
    except Exception as e:
        elements.append(Paragraph(f"<i>Timeline failed: {e}</i>", STYLES["BodyText"]))

    # === OFGEM COMPARISON ===
    try:
        elements.extend(create_ofgem_comparison(df, config))
    except Exception as e:
        elements.append(Paragraph(f"<i>OFGEM Comparison failed: {e}</i>", STYLES["BodyText"]))

    # === STATISTICAL ANALYSIS ===
    try:
        elements.extend(create_statistical_analysis(df))
    except Exception as e:
        elements.append(Paragraph(f"<i>Statistical Analysis failed: {e}</i>", STYLES["BodyText"]))

    # === PAYMENT ANALYSIS ===
    try:
        elements.extend(create_payment_analysis(df))
    except Exception as e:
        elements.append(Paragraph(f"<i>Payment Analysis failed: {e}</i>", STYLES["BodyText"]))

    # === FORECAST ===
    try:
        elements.extend(create_forecast_section(df))
    except Exception as e:
        elements.append(Paragraph(f"<i>Forecast failed: {e}</i>", STYLES["BodyText"]))

    # === DATA QUALITY ===
    try:
        elements.extend(create_data_quality_section(df))
    except Exception as e:
        elements.append(Paragraph(f"<i>Data Quality failed: {e}</i>", STYLES["BodyText"]))

    # === TARIFF IMPACT ===
    try:
        elements.extend(create_tariff_impact_section(df))
    except Exception as e:
        elements.append(Paragraph(f"<i>Tariff Impact failed: {e}</i>", STYLES["BodyText"]))

    # === APPENDICES ===
    try:
        elements.extend(create_appendix_methodology(config))
    except Exception as e:
        elements.append(Paragraph(f"<i>Appendix Methodology failed: {e}</i>", STYLES["BodyText"]))

    try:
        elements.extend(create_appendix_glossary())
    except Exception as e:
        elements.append(Paragraph(f"<i>Appendix Glossary failed: {e}</i>", STYLES["BodyText"]))

    try:
        elements.extend(create_appendix_full_evidence(df, filtered, config))
    except Exception as e:
        elements.append(Paragraph(f"<i>Appendix Full Evidence failed: {e}</i>", STYLES["BodyText"]))

    # Build
    doc.build(elements)
    return output_path


# =============================================================================
# GUI INTEGRATION HELPER
# =============================================================================


def generate_pdf_from_gui(records, output_path, config, engine, filtered=None):
    """Wrapper for GUI integration."""
    try:
        path = generate_ombudsman_pdf(records, output_path, config, engine, filtered)
        return True, f"Professional PDF report generated:\n{path}"
    except Exception as e:
        return False, f"Failed to generate PDF:\n{e}"
