"""EDF Energy PDF Billing-Dispute Report Generator.

Generates a PDF report from extracted billing records.  Output
sections are pulled from ``REPORT_SECTIONS`` so titles and
numbering line up with the DOCX renderer defined in
``edf_report_docx``.  Each section is built by an entry in the
``section_builders`` dispatch dict inside
``generate_ombudsman_pdf``; the
``tests/test_dispatch_parity.py`` test pins the registry ↔
dispatcher ↔ DOCX dispatcher key set parity.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any
from xml.sax.saxutils import escape as xml_escape

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
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

from edf_bill_fetcher.helpers.artifacts import atomic_output
from edf_bill_fetcher.helpers.date_utils import (
    parse_to_sort_date,
)
from edf_bill_fetcher.helpers.formatting import (
    fmt_money,
    fmt_number,
)
from edf_bill_fetcher.helpers.formatting import report_date as fmt_date
from edf_bill_fetcher.models.config import ConfigDict
from edf_bill_fetcher.models.output_overview import (
    OutputOverview,
    build_output_overview,
    report_account_label,
)
from edf_bill_fetcher.models.output_tables import (
    EVIDENCE_HEADERS,
    EVIDENCE_NOTE,
    EVIDENCE_WIDTHS,
    evidence_cells,
    timeline_table,
)
from edf_bill_fetcher.models.report_view import ReportView, capture_report_view
from edf_bill_fetcher.processors.compensation import DISCLAIMER, estimate_compensation

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
        keepWithNext=True,
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
        keepWithNext=True,
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
        alignment=TA_LEFT,
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


def fmt_pct(val: Any, blank_if_na: bool = True) -> str:
    """Format as percentage."""
    if val is None or (isinstance(val, str) and val.upper() in ("N/A", "NA", "")):
        return "" if blank_if_na else "N/A"
    try:
        f = float(val)
        return f"{f:.1%}"
    except (ValueError, TypeError):
        return str(val) if not blank_if_na else ""


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


def _wrapped_table(
    data: list[list],
    widths: list[float],
    *,
    font_size: int = 7,
    content_width: float = CONTENT_WIDTH,
) -> Table:
    """Fit every column inside the frame and display source text literally.

    Plain Table strings do not wrap or interpret XML entities. Paragraph
    cells consume one escaping pass, safely showing the original characters
    while wrapping long invoice and file references within positive widths.
    """
    available = content_width - 12  # Frame's default left/right inner padding.
    total = sum(widths)
    column_widths = [available * width / total for width in widths]
    cells = []
    for row_number, row in enumerate(data):
        style = ParagraphStyle(
            "EvidenceHeader" if row_number == 0 else "EvidenceCell",
            fontName="Helvetica-Bold" if row_number == 0 else "Helvetica",
            fontSize=font_size,
            leading=font_size + 2,
            textColor=Colors.WHITE if row_number == 0 else Colors.DARK_GREY,
            splitLongWords=True,
        )
        cells.append([Paragraph(xml_escape(str(value)), style) for value in row])
    table = Table(cells, colWidths=column_widths, repeatRows=1)
    table.setStyle(make_table_style(num_rows=len(data), font_size=font_size))
    return table


# =============================================================================
# PAGE TEMPLATES
# =============================================================================


def build_doc_template(
    output_path: str, *, include_cover: bool = True, first_landscape: bool = False
) -> BaseDocTemplate:
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
        width, _ = canvas._pagesize
        if page_num > 1 or not include_cover:
            canvas.setFont("Helvetica", 8)
            canvas.setFillColor(Colors.MEDIUM_GREY)
            canvas.drawCentredString(width / 2, 1.2 * cm, f"Page {page_num - int(include_cover)}")
            # Footer line
            canvas.setStrokeColor(Colors.LIGHT_BLUE)
            canvas.setLineWidth(0.5)
            canvas.line(MARGIN, 1.5 * cm, width - MARGIN, 1.5 * cm)
        canvas.restoreState()

    def add_header(canvas, doc):
        """Add header to content pages."""
        canvas.saveState()
        page_num = canvas.getPageNumber()
        width, height = canvas._pagesize
        if page_num > 1 or not include_cover:
            canvas.setFont("Helvetica", 7)
            canvas.setFillColor(Colors.MEDIUM_GREY)
            canvas.drawString(MARGIN, height - 1.2 * cm, "EDF Energy Ombudsman Evidence Report")
            canvas.drawRightString(width - MARGIN, height - 1.2 * cm, "CONFIDENTIAL")
            # Header line
            canvas.setStrokeColor(Colors.LIGHT_BLUE)
            canvas.setLineWidth(0.5)
            canvas.line(MARGIN, height - 1.4 * cm, width - MARGIN, height - 1.4 * cm)
        canvas.restoreState()

    def content_page(canvas, doc):
        add_header(canvas, doc)
        add_page_number(canvas, doc)

    def cover_page(canvas, doc):
        pass  # No header/footer on cover

    templates = [
        PageTemplate(id="content", frames=[content_frame], onPage=content_page, pagesize=A4)
    ]
    if include_cover:
        templates.insert(
            0, PageTemplate(id="cover", frames=[cover_frame], onPage=cover_page, pagesize=A4)
        )
    wide_width, wide_height = landscape(A4)
    templates.append(
        PageTemplate(
            id="landscape",
            frames=[Frame(MARGIN, MARGIN, wide_width - 2 * MARGIN, wide_height - 2 * MARGIN)],
            onPage=content_page,
            pagesize=landscape(A4),
        )
    )
    if first_landscape:
        templates.insert(0, templates.pop())
    doc.addPageTemplates(templates)

    def after_flowable(flowable):
        anchor = getattr(flowable, "_report_anchor", None)
        if anchor:
            doc.canv.bookmarkPage(anchor)
            doc.canv.addOutlineEntry(flowable.getPlainText(), anchor, level=0)

    doc.afterFlowable = after_flowable

    return doc


# =============================================================================
# REPORT SECTIONS
# =============================================================================


# Single source of truth for the report section layout. Both the PDF and DOCX
# generators consume this list to produce the Table of Contents AND to derive
# every body heading at render time. Adding a new section:
#
#   1. Add an entry to ``REPORT_SECTIONS`` below with its key, title, is_appendix.
#   2. Add a matching key to ``ReportOptionsDialog.SECTIONS`` in
#      ui/app.py so the user can toggle it from the GUI.
#   3. Add a ``def create_<name>(...)`` builder function in this module.
#   4. Add the builder to the ``section_builders`` dict in BOTH
#      ``generate_ombudsman_pdf`` (PDF) and ``generate_ombudsman_docx`` (DOCX).
#      Stepping on a missing build entry will raise RuntimeError at dispatch —
#      that's intentional; it's how the registry stays in lockstep with the
#      dispatch.
#
# Sections whose key is in ``config["report_sections"]`` are included. Main
# sections are numbered 1, 2, 3... and appendices are lettered A, B, C..., all
# computed at render time based on the user's selection. So the body's heading
# text and the TOC ALWAYS match, regardless of which sections are toggled.
from edf_bill_fetcher.models.report_sections import (  # noqa: E402,F401
    REPORT_SECTIONS,
    RenderContext,
    ReportSectionMeta,
    _LabelledSection,
)


def _get_package_version() -> str:
    """Return the package version declared in ``pyproject.toml``.

    Delegates to ``edf_bill_fetcher.helpers.version.get_package_version``,
    which walks up from the helpers module to the repo-root
    ``pyproject.toml`` and falls back to ``"0.1.0"`` if the file
    cannot be read or no ``version`` line is present.

    Kept as a thin wrapper (rather than importing the helper name
    directly) because the DOCX report generator and tests import
    this exact name from this module.

    The function is intentionally defensive — a missing or rotated
    pyproject.toml should not break report generation. The cover
    page is what a user will see first, and a missing
    version string is uglier than a stable fallback.
    """
    from edf_bill_fetcher.helpers.version import get_package_version

    return get_package_version()


from edf_bill_fetcher.processors.financial import balance_extremes as _compute_balance_extremes  # noqa: E402,F401
from edf_bill_fetcher.processors.financial import financial_totals as _compute_financial_totals  # noqa: E402,F401
from edf_bill_fetcher.processors.financial import mean_daily_rate as _compute_mean_daily  # noqa: E402,F401


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

    # Account info — escape user-supplied strings to avoid breaking reportlab's
    # internal XML parser if the reference contains <, >, or & characters.
    elements.append(
        Paragraph(f"Account Reference: <b>{xml_escape(account_ref)}</b>", STYLES["CoverInfo"])
    )
    elements.append(
        Paragraph(
            f"Record dates: <b>{xml_escape(str(period_start))}</b> to <b>{xml_escape(str(period_end))}</b>",
            STYLES["CoverInfo"],
        )
    )
    elements.append(Spacer(1, 1 * cm))

    # Report metadata
    elements.append(
        Paragraph(f"Report Generated: <b>{xml_escape(str(report_date))}</b>", STYLES["CoverInfo"])
    )
    elements.append(Paragraph("Prepared by: <b>EDF Evidence Collector</b>", STYLES["CoverInfo"]))
    elements.append(Spacer(1, 2 * cm))

    # Confidential notice
    elements.append(Paragraph("CONFIDENTIAL — FOR OMBUDSMAN REVIEW ONLY", STYLES["Confidential"]))

    # Version info
    elements.append(Spacer(1, 3 * cm))
    version = _get_package_version()
    elements.append(
        Paragraph(
            f"Generated by EDF Evidence Collector v{version}<br/>"
            "All data extracted from original source documents (EDF bills, HTM exports, email archives).<br/>"
            "Methodology detailed in Appendix A.",
            STYLES["SmallText"],
        )
    )

    return elements


def create_table_of_contents(ctx: RenderContext, *, linked: bool = False) -> list:
    """Create table of contents as a single table, driven by ``ctx``.

    Numbers and titles come straight from the RenderContext — same registry
    as the body builders — so the TOC will always line up with the body
    regardless of which sections the user toggled in the GUI.
    """
    elements = []
    elements.append(Paragraph("Table of Contents", STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.5 * cm))

    if not ctx.sections_in_order:
        elements.append(Paragraph("<i>No report sections selected.</i>", STYLES["BodyText"]))
        elements.append(PageBreak())
        return elements

    toc_data = [["No.", "Section"]]
    for spec in ctx.sections_in_order:
        title = (
            Paragraph(
                f'<link href="#sec-{spec.section.key}" color="#10367A">{xml_escape(spec.section.title)}</link>',
                STYLES["BodyText"],
            )
            if linked
            else spec.section.title
        )
        toc_data.append([spec.label, title])

    toc_table = Table(toc_data, colWidths=[1.5 * cm, CONTENT_WIDTH - 1.5 * cm])
    style = TableStyle(
        [
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("TEXTCOLOR", (0, 0), (-1, -1), Colors.DARK_GREY),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),  # Header row bold
            ("BACKGROUND", (0, 0), (-1, 0), Colors.NAVY),
            ("TEXTCOLOR", (0, 0), (-1, 0), Colors.WHITE),
            ("LINEBELOW", (0, 0), (-1, 0), 1.5, Colors.NAVY),
        ]
    )

    # Header row already styled above; nothing per-row needed beyond defaults
    toc_table.setStyle(style)
    elements.append(toc_table)
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
    opening_balance: float | None = None,
    closing_balance: float | None = None,
    ctx: RenderContext | None = None,
    overview: OutputOverview | None = None,
) -> list:
    """Create executive summary section."""
    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("exec_summary")

    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    # Overview paragraph
    intro = (
        f"This report reviews the supplied EDF Energy billing evidence "
        f"for account <b>{xml_escape(account_ref)}</b>, with record dates from "
        f"<b>{xml_escape(str(period_start))}</b> to <b>{xml_escape(str(period_end))}</b>. "
        f"The analysis encompasses <b>{total_records}</b> evidence record{'s' if total_records != 1 else ''} "
        f"from the sources listed in the evidence index."
    )
    elements.append(Paragraph(intro, STYLES["BodyText"]))
    elements.append(Spacer(1, 0.3 * cm))

    # Financial summary
    net_change = total_charges - total_payments
    elements.append(Paragraph("<b>Financial Summary</b>", STYLES["SubSectionHeader"]))
    elements.append(Spacer(1, 0.2 * cm))

    fin_data = [
        ["Metric", "Amount"],
        ["Total Charges (Debits)", fmt_money(total_charges)],
        ["Net Payments/Credits (After Refunds)", fmt_money(total_payments)],
        ["Recorded Net Movement", fmt_money(net_change)],
        # Opening/Closing balance rows were hard-coded "—" placeholders
        # before this commit.  Now the dispatcher pre-computes them via
        # ``_compute_balance_extremes`` so the reader sees the actual
        # earliest and latest Account Balance figures from the dataset.
        # ``"—"`` is still rendered as a fallback when the helper could
        # not extract real numbers (no dated/amounted rows), which is
        # the honest result rather than the steady state.
        [
            "First observed balance",
            fmt_money(opening_balance) if opening_balance is not None else "—",
        ],
        [
            "Last observed balance",
            fmt_money(closing_balance) if closing_balance is not None else "—",
        ],
    ]
    fin_table = Table(fin_data, colWidths=[8 * cm, 5 * cm])
    fin_table.setStyle(make_table_style(num_rows=6))
    elements.append(fin_table)
    elements.append(Spacer(1, 0.3 * cm))
    if overview:
        elements.append(
            Paragraph(
                f"Analysis as of {overview.analysis_date}. {xml_escape(overview.balance_note)}",
                STYLES["SmallText"],
            )
        )
        elements.append(Paragraph("Review priorities", STYLES["SubSectionHeader"]))
        for priority in overview.priorities:
            elements.append(
                Paragraph(
                    f"<b>{xml_escape(priority.title)}</b> — {xml_escape(priority.action)}",
                    STYLES["BodyText"],
                )
            )

    # Key findings
    elements.append(Paragraph("<b>Key Findings</b>", STYLES["SubSectionHeader"]))
    elements.append(Spacer(1, 0.2 * cm))

    findings = []
    if flag_count.get("HIGH", 0) > 0:
        findings.append(
            f"<b>{flag_count['HIGH']} HIGH-severity findings</b> require review. "
            "Read the dated findings for the specific checks triggered and their supporting evidence."
        )
    if flag_count.get("MEDIUM", 0) > 0:
        findings.append(
            f"<b>{flag_count['MEDIUM']} MEDIUM-severity findings</b> require review. "
            "The findings table explains each observation and its account context."
        )
    if flag_count.get("INFO", 0) > 0:
        findings.append(
            f"<b>{flag_count['INFO']} informational findings</b> provide context for manual review."
        )

    if not findings:
        findings.append(
            "No automated flags were generated. This does not establish that all bills and payments are present or correct."
        )

    for i, finding in enumerate(findings, 1):
        elements.append(Paragraph(f"{i}. {finding}", STYLES["BulletText"]))

    elements.append(Spacer(1, 0.3 * cm))

    # Conclusion
    elements.append(Paragraph("<b>Conclusion</b>", STYLES["SubSectionHeader"]))
    conclusion = (
        "This report summarizes the supplied evidence and highlights observations for review. "
        "Balance changes and missing records alone do not establish overcharging or a regulatory breach. "
        "Check the original documents, account context and applicable rules before making a claim."
    )
    elements.append(Paragraph(conclusion, STYLES["BodyText"]))

    elements.append(PageBreak())
    return elements


# =============================================================================
# KEY FINDINGS SUMMARY
# =============================================================================


def create_key_findings_table(flags: list, ctx: RenderContext | None = None) -> list:
    """Create key findings summary table from flags."""
    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("key_findings")

    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    if not flags:
        elements.append(
            Paragraph(
                "No automated flags were generated. Manual review of the original evidence "
                "and its completeness is still required.",
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
        ["HIGH", str(len(high)), "Review original evidence and account context"],
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
            text = (
                f"<b>{i}. {xml_escape(str(ftype))}</b> "
                f"({xml_escape(str(date_str))}, {xml_escape(str(amt_str))}) — "
                f"{xml_escape(str(detail))}"
            )
            elements.append(Paragraph(text, STYLES["BulletText"]))

    if medium:
        elements.append(Spacer(1, 0.3 * cm))
        elements.append(Paragraph("<b>MEDIUM Severity Findings</b>", STYLES["SubSectionHeader"]))
        elements.append(Spacer(1, 0.2 * cm))

        for i, (ftype, date, amt, detail, _sev) in enumerate(medium, 1):
            date_str = fmt_date(date)
            amt_str = fmt_money(amt) if amt else ""
            text = (
                f"<b>{i}. {xml_escape(str(ftype))}</b> "
                f"({xml_escape(str(date_str))}, {xml_escape(str(amt_str))}) — "
                f"{xml_escape(str(detail))}"
            )
            elements.append(Paragraph(text, STYLES["BulletText"]))

    elements.append(PageBreak())
    return elements


# =============================================================================
# EVIDENCE INDEX
# =============================================================================


def create_evidence_index(df: pd.DataFrame, engine: Any, ctx: RenderContext | None = None) -> list:
    """Create evidence index with source cross-references."""
    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("evidence_index")

    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    # Source summary
    source_counts = df["Source"].value_counts()
    source_data = [["Source", "Records", "Percentage"]]
    total = len(df)
    for src, cnt in source_counts.items():
        source_data.append([str(src), str(cnt), f"{cnt / total:.1%}"])
    source_data.append(["TOTAL", str(total), "100.0%"])

    t = _wrapped_table(source_data, [8, 3, 3], font_size=8)
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
            Paragraph(
                f"<b>{xml_escape(str(src))}</b> ({len(src_df)} records)",
                STYLES["SubSubSectionHeader"],
            )
        )

        # Summary table for this source
        detail_data = [["Date", "Invoice #", "Amount", "Period", "Entry Type", "Reading"]]
        for _, row in src_df.iterrows():
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

        t = _wrapped_table(detail_data, [2.5, 3, 3, 4, 3, 2])
        elements.append(t)
        elements.append(Spacer(1, 0.3 * cm))

    elements.append(PageBreak())
    return elements


# =============================================================================
# DETAILED FINDINGS SECTIONS
# =============================================================================


def create_anomaly_detail_section(
    flags: list, df: pd.DataFrame, ctx: RenderContext | None = None
) -> list:
    """Create detailed anomaly findings section."""
    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("detailed_findings")
    # Used for the dynamic 4.x subsection labels under this section.
    parent_label = ctx.short_label("detailed_findings").rstrip(".")  # e.g. "4"

    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    if not flags:
        elements.append(
            Paragraph(
                "No dated automated findings to expand. Check source coverage and review priorities before relying on the absence of flags.",
                STYLES["BodyText"],
            )
        )
        return elements

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

    for cat_idx, (cat, cat_flags) in enumerate(categories.items(), 1):
        if not cat_flags:
            continue

        # Subsections under "Detailed Findings" — number is dynamic: 4.1, 4.2,
        # ... or whatever the parent label resolves to in the live report.
        sub_label = f"{parent_label}.{cat_idx}" if parent_label else str(cat_idx)
        elements.append(
            Paragraph(
                f"{sub_label} {cat.replace('_', ' ').title()}",
                STYLES["SubSectionHeader"],
            )
        )
        elements.append(Spacer(1, 0.2 * cm))

        # Table of findings
        detail_data = [["#", "Date", "Amount", "Severity", "Detail"]]
        for i, (_ftype, date, amt, detail, sev) in enumerate(cat_flags, 1):
            truncated = detail[:200] + ("..." if len(detail) > 200 else "")
            detail_data.append(
                [
                    str(i),
                    fmt_date(date),
                    fmt_money(amt) if amt is not None else "",
                    str(sev),
                    truncated,
                ]
            )

        t = _wrapped_table(detail_data, [0.8, 2.5, 3, 2, 7.7])
        style = make_table_style(num_rows=len(detail_data), font_size=7)
        # Color severity column
        for row_idx, (_, _, _, sev, _) in enumerate(cat_flags, 1):
            style.add("TEXTCOLOR", (3, row_idx), (3, row_idx), severity_color(sev))
            style.add("FONTNAME", (3, row_idx), (3, row_idx), "Helvetica-Bold")
            t._cellvalues[row_idx][3].style.textColor = severity_color(sev)
            t._cellvalues[row_idx][3].style.fontName = "Helvetica-Bold"
        t.setStyle(style)
        elements.append(t)
        elements.append(Spacer(1, 0.4 * cm))

    elements.append(PageBreak())
    return elements


def create_revision_section(model: dict | None) -> list:
    """Append stored revision comparisons with wrapping and repeated table headers."""
    if model is None:
        return []
    from edf_bill_fetcher.models.output_tables import revision_note, revision_table

    elements = [
        Paragraph("Bill revision review", STYLES["SubSectionHeader"]),
        Paragraph(xml_escape(revision_note(model)), STYLES["BodyText"]),
    ]
    if model["relationships"]:
        rows = revision_table(model)
        table = _wrapped_table(rows, [3.3, 4.2, 4.2, 4.3])
        table.splitInRow = 1
        table.setStyle(make_table_style(num_rows=len(rows), font_size=7))
        elements.append(table)
    else:
        elements.append(
            Paragraph(
                "No revision relationships were found. You can add a missed relationship in the app.",
                STYLES["BodyText"],
            )
        )
    return elements


def create_timeline_section(
    df: pd.DataFrame, flags: list, ctx: RenderContext | None = None
) -> list:
    """Create chronological timeline of events."""
    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("timeline")

    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    timeline_data = timeline_table(df, flags)
    t = _wrapped_table(timeline_data, [2, 3, 2.5, 2.5, 7], font_size=8)
    elements.append(t)

    elements.append(PageBreak())
    return elements


# =============================================================================
# OFGEM PRICE CAP COMPARISON
# =============================================================================


# OFGEM PRICE CAP COMPARISON
# =============================================================================


def create_ofgem_comparison(
    df: pd.DataFrame,
    config: dict | None = None,
    ctx: RenderContext | None = None,
    *,
    payload: Any = None,
) -> list:
    """Create OFGEM price cap comparison section."""
    from edf_bill_fetcher.models.report_models import compute_ofgem_comparison

    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("ofgem")

    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    elements.append(
        Paragraph(
            "The following analysis compares the effective unit rates charged on EDF bills "
            "against the OFGEM Default Tariff Cap (Price Cap) for the corresponding periods. "
            "This is a contextual GB-average benchmark; an exceedance requires review of the account’s applicable rate and is not a legal determination.",
            STYLES["BodyText"],
        )
    )
    elements.append(Spacer(1, 0.3 * cm))

    comparison = payload if payload is not None else compute_ofgem_comparison(df, config)

    if not comparison.rows:
        elements.append(
            Paragraph(
                "No billing records with a usable extracted energy unit rate and consumption period "
                "are available for comparison. A total charge divided by usage is not a substitute.",
                STYLES["BodyText"],
            )
        )
        elements.append(PageBreak())
        return elements

    # Build comparison table
    cap_data = [["Period", "Bill Unit Rate (p/kWh)", "OFGEM Cap (p/kWh)", "Difference", "Status"]]
    MISSING = "—"
    for row in comparison.rows:
        cap_data.append(
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

    # Add summary row
    cap_data.append(
        [
            "OVERALL",
            "—",
            "—",
            comparison.overall_diff,
            comparison.overall_verdict,
        ]
    )

    # Phase 1.1 + portability: the Difference and Status columns need
    # to comfortably fit the longest status string this table can
    # ever render, namely ``"1 period(s) not benchmarked"`` wired
    # into the OVERALL summary row in the unbenchmarked-quarter
    # path.  Default column widths were truncating that to
    # "1 period(s) not benchm" with the remaining characters
    # spilling into the next column.  Give the last two columns an
    # extra half-centimetre each and tighten Period to keep the
    # table inside the printable area on a portrait A4.
    t = _wrapped_table(cap_data, [3, 3, 3.5, 4, 3.5], font_size=8)

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
            t._cellvalues[i][4].style.textColor = Colors.RED
            t._cellvalues[i][4].style.fontName = "Helvetica-Bold"
        elif "REVIEW" in status:
            style.add("TEXTCOLOR", (4, i), (4, i), Colors.RED)
            style.add("FONTNAME", (4, i), (4, i), "Helvetica-Bold")
            t._cellvalues[i][4].style.textColor = Colors.RED
            t._cellvalues[i][4].style.fontName = "Helvetica-Bold"
        elif "BELOW" in status:
            style.add("TEXTCOLOR", (4, i), (4, i), Colors.GREEN)
            t._cellvalues[i][4].style.textColor = Colors.GREEN
    t.setStyle(style)

    elements.append(t)
    elements.append(Spacer(1, 0.3 * cm))

    elements.append(
        Paragraph(
            "<b>Methodology:</b> Extracted energy unit rates are allocated to consumption quarters "
            "and weighted by usage, or days when usage is unavailable. The model does not divide "
            "total bill charges by kWh or assess standing charges here. Missing tariff/payment/fuel "
            "context remains unassessed. OFGEM data is a GB-average direct-debit electricity "
            "benchmark, not an account-specific legal cap.",
            STYLES["SmallText"],
        )
    )

    elements.append(PageBreak())
    return elements


# =============================================================================
# STATISTICAL ANALYSIS SECTION
# =============================================================================


def create_statistical_analysis(
    dfc: pd.DataFrame, ctx: RenderContext | None = None, *, payload: Any = None
) -> list:
    """Create statistical analysis section."""
    from edf_bill_fetcher.models.report_models import compute_statistical_analysis

    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("statistical")

    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    sa = payload if payload is not None else compute_statistical_analysis(dfc)
    n = sa.count

    if n < 3:
        elements.append(
            Paragraph(
                "At least three usable balance observations in one ledger are needed. No statistical conclusion is available for this selection.",
                STYLES["BodyText"],
            )
        )
        elements.append(PageBreak())
        return elements

    elements.append(
        Paragraph(
            "These statistics describe recorded balances, not individual bill costs. Outliers identify observations to inspect; they do not establish a billing error.",
            STYLES["BodyText"],
        )
    )
    # Descriptive stats
    mean_amt = sa.mean
    std_amt = sa.std
    cv = fmt_pct(std_amt / mean_amt) if mean_amt and mean_amt > 0 else "N/A"

    stats_data = [
        ["Statistic", "Value"],
        ["Count", fmt_number(n, 0)],
        ["Mean (£)", fmt_money(mean_amt)],
        ["Median (£)", fmt_money(sa.median)],
        ["Std Deviation (£)", fmt_money(std_amt)],
        ["Min (£)", fmt_money(sa.minimum)],
        ["Max (£)", fmt_money(sa.maximum)],
        ["Range (£)", fmt_money(sa.range)],
        ["Coefficient of Variation", cv],
    ]

    t = Table(stats_data, colWidths=[10 * cm, 5 * cm])
    t.setStyle(make_table_style(num_rows=len(stats_data)))
    elements.append(t)
    elements.append(Spacer(1, 0.5 * cm))

    # Rolling statistics
    elements.append(
        Paragraph("<b>6-Period Rolling Statistics (Latest)</b>", STYLES["SubSectionHeader"])
    )
    roll_mean = sa.rolling["mean"]
    roll_std = sa.rolling["std"]
    roll_min = sa.rolling["min"]
    roll_max = sa.rolling["max"]

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
    if sa.shapiro_stat is not None and sa.shapiro_p is not None:
        elements.append(
            Paragraph("<b>Distribution Normality (Shapiro-Wilk)</b>", STYLES["SubSectionHeader"])
        )
        norm_data = [
            ["Test", "Statistic", "p-value", "Normal?"],
            [
                "Shapiro-Wilk",
                f"{sa.shapiro_stat:.4f}",
                f"{sa.shapiro_p:.4f}",
                "Yes" if sa.shapiro_p > 0.05 else "No",
            ],
        ]
        t = Table(norm_data, colWidths=[4 * cm, 3 * cm, 3 * cm, 3 * cm])
        t.setStyle(make_table_style(num_rows=len(norm_data)))
        elements.append(t)
    else:
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


def create_payment_analysis(
    dfc: pd.DataFrame, ctx: RenderContext | None = None, *, payload: Any = None
) -> list:
    """Create payment & credit analysis section."""
    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("payment")

    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    from edf_bill_fetcher.models.report_models import compute_payment_analysis

    pa = payload if payload is not None else compute_payment_analysis(dfc)
    if not pa.chronology.empty:
        elements.append(Paragraph("<b>Payment Summary</b>", STYLES["SubSectionHeader"]))

        pay_data = [
            ["Metric", "Value"],
            ["Total Payments/Credits", str(len(pa.chronology))],
            ["Total Amount Paid", fmt_money(pa.total_paid)],
            ["Average Payment", fmt_money(pa.avg_payment)],
            ["Median Payment", fmt_money(pa.median_payment)],
            ["Largest Payment", fmt_money(pa.largest_payment)],
            ["Smallest Payment", fmt_money(pa.smallest_payment)],
        ]
        if pa.avg_interval_days is not None:
            pay_data.append(["Avg Interval (days)", fmt_number(pa.avg_interval_days, 1)])
            pay_data.append(["Median Interval (days)", fmt_number(pa.median_interval_days, 1)])

        t = Table(pay_data, colWidths=[10 * cm, 5 * cm])
        t.setStyle(make_table_style(num_rows=len(pay_data)))
        elements.append(t)
        elements.append(Spacer(1, 0.5 * cm))

        # Payment detail table
        elements.append(Paragraph("<b>Payment Chronology</b>", STYLES["SubSectionHeader"]))
        pay_detail = [["Date", "Type", "Amount (£)", "Balance After", "Details"]]
        for _, row in pa.chronology.iterrows():
            pay_detail.append(
                [
                    fmt_date(row["Date"]),
                    str(row["Entry Type"]),
                    fmt_money(row["_amount"]),
                    fmt_money(row.get("Balance (£)")),
                    str(row.get("Details", ""))[:80],
                ]
            )
        t = _wrapped_table(pay_detail, [2.5, 2.5, 3, 3, 5])
        elements.append(t)
    else:
        elements.append(
            Paragraph(
                "No payment/credit records were supplied. This does not establish that no payments were made; check statements or payment confirmations.",
                STYLES["BodyText"],
            )
        )

    elements.append(PageBreak())
    return elements


# =============================================================================
# FORECAST SECTION
# =============================================================================


def create_forecast_section(
    dfc: pd.DataFrame, ctx: RenderContext | None = None, *, payload: Any = None
) -> list:
    """Create forecast & projection section."""
    from edf_bill_fetcher.models.report_models import compute_forecast

    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("forecast")

    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    elements.append(
        Paragraph(
            "These are projections of observed balances, not future bill charges or compensation. "
            "They are indicative context only; missing payments, tariff changes and other account "
            "events can make a balance trend misleading.",
            STYLES["BodyText"],
        )
    )
    elements.append(Spacer(1, 0.3 * cm))

    fc = payload if payload is not None else compute_forecast(dfc)
    if fc.n < 3:
        elements.append(
            Paragraph(
                "Insufficient data for forecasting (need at least 3 periods).", STYLES["BodyText"]
            )
        )
        elements.append(PageBreak())
        return elements

    linear_forecast = fc.linear_forecast
    ema_forecast = fc.ema_forecast
    hw_forecast = fc.hw_forecast
    model_info = fc.model_info

    # Build forecast table
    forecast_header = ["Forecast Month", "Linear Reg. (£)", "EMA (£)"]
    if hw_forecast:
        forecast_header.append("Holt-Winters (£)")

    forecast_data = [forecast_header]
    for i in range(6):
        row = [f"+{i + 1} Month", fmt_money(linear_forecast[i]), fmt_money(ema_forecast[i])]
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
        Paragraph("<b>Next 6 Months — Multi-Method Projection</b>", STYLES["SubSectionHeader"])
    )
    elements.append(Spacer(1, 0.2 * cm))
    elements.append(t)
    elements.append(Spacer(1, 0.3 * cm))

    for info in model_info:
        elements.append(Paragraph(info, STYLES["SmallText"]))

    elements.append(Spacer(1, 0.3 * cm))

    elements.append(
        Paragraph(
            "<i>Note: Projections assume continuation of current trends and do not account for "
            "seasonal variations, tariff changes, or policy changes. "
            "Full forecasting with advisory scenario bands available in the Excel workbook (Forecast &amp; Projection tab).</i>",
            STYLES["SmallText"],
        )
    )

    elements.append(PageBreak())
    return elements


# =============================================================================
# DATA QUALITY SECTION
# =============================================================================


def create_data_quality_section(
    df: pd.DataFrame, ctx: RenderContext | None = None, *, payload: Any = None
) -> list:
    """Create data quality assessment section."""
    from edf_bill_fetcher.models.report_models import compute_data_quality_report

    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("data_quality")

    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    dq = payload if payload is not None else compute_data_quality_report(df)
    total = dq.total_records
    date_parsed = dq.date_parsed
    amt_complete = dq.amt_complete
    period_complete = dq.period_complete
    reading_classified = dq.reading_classified
    dup_count = dq.duplicate_count

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


def create_tariff_impact_section(
    dfc: pd.DataFrame, ctx: RenderContext | None = None, *, payload: Any = None
) -> list:
    """Create tariff impact analysis section."""
    from edf_bill_fetcher.models.report_models import compute_tariff_analysis

    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("tariff")

    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    analysis = payload if payload is not None else compute_tariff_analysis(dfc)

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

    if analysis.empty:
        elements.append(Paragraph("No valid tariff records found.", STYLES["BodyText"]))
        elements.append(PageBreak())
        return elements

    tariff_stats = analysis.stats

    elements.append(Paragraph("<b>Unit Rate by Tariff</b>", STYLES["SubSectionHeader"]))
    tariff_table = [
        ["Tariff", "Records", "Avg Rate (p/kWh)", "Median", "Min", "Max", "Avg Charge (£)"]
    ]
    for _, row in tariff_stats.iterrows():
        tariff_table.append(
            [
                str(row["Tariff"]),
                str(int(row["count"])),
                fmt_number(row["avg_unit_rate"], 2),
                fmt_number(row["median_unit_rate"], 2),
                fmt_number(row["min_unit_rate"], 2),
                fmt_number(row["max_unit_rate"], 2),
                fmt_money(row["avg_charge"]) if pd.notna(row["avg_charge"]) else "N/A",
            ]
        )

    t = _wrapped_table(tariff_table, [4, 2, 3, 2.5, 2.5, 2.5, 3], font_size=8)
    elements.append(t)
    elements.append(Spacer(1, 0.5 * cm))

    elements.append(
        Paragraph(
            f"<b>Tariff Segments Observed: {analysis.tariff_changes}</b>",
            STYLES["SubSectionHeader"],
        )
    )

    elements.append(PageBreak())
    return elements


def create_compensation_section(
    dfc: pd.DataFrame,
    config: dict[str, Any] | None = None,
    ctx: RenderContext | None = None,
    *,
    payload: Any = None,
) -> list:
    """Create the Compensation Analysis section.

    Renders the indicative-compensation claims from
    :func:`edf_bill_fetcher.processors.compensation.estimate_compensation`,
    grouped by category, with a trailing total of ``indicative_amount`` and
    the module ``DISCLAIMER``.  Empty rows render a brief no-claims note.
    """
    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("compensation")

    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    rows = payload if payload is not None else estimate_compensation(dfc, config)

    if not rows:
        elements.append(
            Paragraph(
                "No compensation claims identified in the extracted records.",
                STYLES["BodyText"],
            )
        )
        elements.append(PageBreak())
        return elements

    # Group by category, preserving the estimator's stable emission order.
    categories: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        cat = str(row.get("category", ""))
        categories.setdefault(cat, []).append(row)

    for cat, cat_rows in categories.items():
        elements.append(Paragraph(f"<b>{xml_escape(cat)}</b>", STYLES["SubSectionHeader"]))
        table_rows: list[list[str]] = [
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
        t = _wrapped_table(table_rows, [3.2, 2.2, 2.2, 2.2, 1.4, 1.6, 2.6, 4.6])
        elements.append(t)
        elements.append(Spacer(1, 0.3 * cm))

    total = sum(float(r.get("indicative_amount") or 0.0) for r in rows)
    elements.append(
        Paragraph(
            f"<b>Total Indicative Compensation: {fmt_money(total)}</b>", STYLES["SubSectionHeader"]
        )
    )
    elements.append(Spacer(1, 0.3 * cm))
    elements.append(Paragraph(DISCLAIMER, STYLES["BodyText"]))

    elements.append(PageBreak())
    return elements


# =============================================================================
# APPENDICES
# =============================================================================


def create_appendix_methodology(config: ConfigDict, ctx: RenderContext | None = None) -> list:
    """Create Methodology appendix."""
    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("appendix_methodology")

    elements.append(PageBreak())
    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    sections = [
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
                "EDF billing language (e.g., 'Current balance £X debit', 'Total charges for this period'). "
                "Patterns execute in priority order; first match wins.",
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
            "A.4 Classification Logic",
            [
                "Entry Type: 'New Bill' (period charge + period dates), 'Ongoing Balance' "
                "(cumulative balance only), 'Payment', 'Credit'.",
                "Reading Type: regex classification — 'Estimated' (estimated/est./estimate), "
                "'Actual' (actual/customer reading/your reading), 'Smart' (smart meter/automated).",
            ],
        ),
        (
            "A.5 Configuration Used",
            [
                f"Valuation date: {config.get('as_of', 'Unavailable')}",
                f"Minimum Amount Threshold: {fmt_money(config.get('min_amount', 500))}",
                f"Analysis Threshold: {fmt_money(config.get('analysis_min', 0))}",
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
            elements.append(Paragraph(xml_escape(bullet), STYLES["BulletText"]))
        elements.append(Spacer(1, 0.2 * cm))

    elements.append(PageBreak())
    return elements


def create_appendix_full_evidence(
    df: pd.DataFrame,
    filtered: list | None = None,
    config: dict | None = None,
    ctx: RenderContext | None = None,
    content_width: float = CONTENT_WIDTH,
) -> list:
    """Create Full Evidence Table appendix, plus optional Filtered Records sub-table."""
    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("appendix_full_evidence")

    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    elements.append(
        Paragraph(
            EVIDENCE_NOTE,
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

    evidence_header = EVIDENCE_HEADERS
    evidence_data = [
        evidence_header,
        *[evidence_cells(row) for row in df_sorted.to_dict("records")],
    ]

    # Relative widths; the wrapped table scales these to the actual frame.
    col_widths = list(EVIDENCE_WIDTHS)

    t = _wrapped_table(evidence_data, col_widths, font_size=8, content_width=content_width)
    t.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
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
        # Continuation of the same appendix, not a new section. Use the same
        # alphabetic label as the parent (e.g. "A." for methodology).
        cont_label = ctx.short_label("appendix_full_evidence").rstrip(".")
        cont_heading = (
            f"{cont_label}. (cont.) Filtered Records (Below {min_amt} Threshold)"
            if cont_label
            else f"Filtered Records (Below {min_amt} Threshold)"
        )
        elements.append(Paragraph(cont_heading, STYLES["SectionHeader"]))
        elements.append(Spacer(1, 0.3 * cm))

        filt_data = [evidence_header]
        filt_data.extend(evidence_cells(row) for row in filtered)

        if len(filt_data) > 1:
            t2 = _wrapped_table(filt_data, col_widths, font_size=8, content_width=content_width)
            t2.setStyle(
                TableStyle(
                    [
                        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                        ("FONTSIZE", (0, 0), (-1, -1), 8),
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


def create_appendix_glossary(ctx: RenderContext | None = None) -> list:
    """Create Glossary appendix."""
    elements = []
    if ctx is None:
        ctx = RenderContext()
    heading = ctx.heading("appendix_glossary")

    elements.append(Paragraph(heading, STYLES["SectionHeader"]))
    elements.append(Spacer(1, 0.3 * cm))

    terms = {
        "Period Charge (£)": "The charge for the specific billing period (not cumulative balance). "
        "Equivalent to 'Total charges for this period' on new EDF invoices.",
        "Amount (£)": "The primary balance figure — typically the current cumulative account balance "
        "on new invoices, or the running balance on HTM exports.",
        "Energy Unit Rate (p/kWh)": "The energy price extracted from the bill, excluding "
        "standing charges. Total bill cost divided by usage is not an equivalent unit rate.",
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

    t = _wrapped_table(glossary_data, [4, 12], font_size=8)
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
) -> str:
    """Capture legacy inputs once, then render their accepted result."""
    view = capture_report_view(records, dict(config), engine, filtered, formats=("pdf",))
    render_pdf(view, output_path)
    return output_path


@atomic_output()
def render_pdf(view: ReportView, output_path: str) -> str:
    """Render the accepted result using the existing PDF layout."""
    section_failures = []
    records, config, engine = view.records, view.config, view.engine
    filtered = engine.filtered_records
    df = view.frame()
    if df.empty:
        if not view.review_notice:
            raise ValueError("No records to report on")
        from edf_bill_fetcher.io.reporters.review_empty import render_empty_review

        return render_empty_review(view, output_path, "pdf")

    # Validate required parameters — create a minimal engine if not provided
    # (CLI usage without --engine-data should still work)
    if engine is None:

        class MinimalEngine:
            pdf_count: int = 0
            email_count: int = 0
            filtered_records: list[Any] = []

        engine = MinimalEngine()

    # Ensure required columns exist with defaults
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

    # Account reference
    acc_ref = report_account_label(view)

    # Period bounds
    dates_parsed = df["Date"].apply(parse_to_sort_date)
    valid_dates = dates_parsed.dropna()
    period_start = fmt_date(valid_dates.min()) if not valid_dates.empty else "Unknown"
    period_end = fmt_date(valid_dates.max()) if not valid_dates.empty else "Unknown"

    # Financial totals — see ``_compute_financial_totals`` for why we
    # route off Entry Type rather than splitting ``Amount (£)`` on sign
    # (the column is a positive cumulative balance, not a signed delta).
    charges, payments = view.snapshot.summary.charges, view.snapshot.summary.net_credits
    # Opening/closing balance — see ``_compute_balance_extremes`` for why
    # we no longer ship ``"—"`` placeholders to the ombudsman-facing
    # executive summary.  When the dataframe has at least one dated
    # record with a numeric balance we emit real numbers; otherwise the
    # helper returns ``(None, None)`` and the table renders ``"—"``,
    # which is now an honest fallback rather than the steady state.
    opening_balance, closing_balance = (
        view.snapshot.summary.opening_balance,
        view.snapshot.summary.closing_balance,
    )

    # Compute dispute flags from the data
    analysis_df = view.frame("analysis")
    dispute = view.model("dispute")
    flags, flag_counts = dispute.flags, dispute.counts

    enabled_sections = view.context.selected

    def section_enabled(key: str) -> bool:
        return key in enabled_sections

    # RenderContext derives all number/letter labels once, from the registry.
    render_ctx = RenderContext(enabled_sections)

    # Build document
    only_evidence = enabled_sections == {"appendix_full_evidence"}
    doc = build_doc_template(
        output_path,
        include_cover=section_enabled("cover"),
        first_landscape=only_evidence,
    )
    elements = []

    # === COVER PAGE ===
    if section_enabled("cover"):
        elements.extend(
            create_cover_page(
                acc_ref, period_start, period_end, datetime.now().strftime("%d %B %Y")
            )
        )
        elements.append(NextPageTemplate("content"))
        elements.append(PageBreak())

    if view.review_notice:
        elements.append(Paragraph(xml_escape(view.review_notice), STYLES["BodyText"]))
    if view.ocr_notice:
        elements.append(Paragraph(xml_escape(view.ocr_notice), STYLES["BodyText"]))
        elements.append(Spacer(1, 0.3 * cm))

    # === TABLE OF CONTENTS ===
    if section_enabled("toc"):
        try:
            elements.extend(create_table_of_contents(render_ctx, linked=True))
        except Exception as e:
            section_failures.append(str(e))
            elements.append(
                Paragraph(
                    f"<i>Table of Contents failed: {xml_escape(str(e))}</i>",
                    STYLES["BodyText"],
                )
            )

    # === SECTION DISPATCH (data-driven — keys/ordering live in REPORT_SECTIONS) ===
    # Each entry: (key, required_factory(ctx) -> dict of kwargs, builder_callable)
    # ``required_factory`` returns the kwargs the builder needs; ``builder_callable``
    # is invoked with those kwargs to produce the reportlab elements. Adding a new
    # node to REPORT_SECTIONS without an entry here will raise at dispatch time —
    # keep them in sync. Build registry mirrors ``REPORT_SECTIONS``.
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
            lambda: {"df": df, "engine": engine},
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
            lambda: {"dfc": analysis_df},
            lambda kwargs: create_statistical_analysis(**kwargs),
        ),
        "payment": (
            lambda: {"dfc": analysis_df},
            lambda kwargs: create_payment_analysis(**kwargs),
        ),
        "forecast": (
            lambda: {"dfc": analysis_df},
            lambda kwargs: create_forecast_section(**kwargs),
        ),
        "data_quality": (
            lambda: {"df": df},
            lambda kwargs: create_data_quality_section(**kwargs),
        ),
        "tariff": (
            lambda: {"dfc": analysis_df},
            lambda kwargs: create_tariff_impact_section(**kwargs),
        ),
        "compensation": (
            lambda: {"dfc": analysis_df, "config": config},
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

    # Sections missing a builder entry here will fail loudly rather than skip.
    for section in REPORT_SECTIONS:
        if not section_enabled(section.key):
            continue
        entry = section_builders.get(section.key)
        if entry is None:
            raise RuntimeError(
                f"REPORT_SECTIONS lists '{section.key}' but no builder is wired "
                f"in generate_ombudsman_pdf. Add it to section_builders."
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
            if section.key == "appendix_full_evidence":
                kwargs["content_width"] = landscape(A4)[0] - 2 * MARGIN
                if not only_evidence:
                    if elements and isinstance(elements[-1], PageBreak):
                        elements.insert(len(elements) - 1, NextPageTemplate("landscape"))
                    else:
                        elements.extend([NextPageTemplate("landscape"), PageBreak()])
            section_elements = invoke(kwargs)
            # Section builders also serve standalone callers. Compose their
            # content here without the legacy forced page between every section.
            while section_elements and isinstance(section_elements[0], PageBreak):
                section_elements.pop(0)
            while section_elements and isinstance(section_elements[-1], PageBreak):
                section_elements.pop()
            for flowable in section_elements:
                if isinstance(flowable, Paragraph) and flowable.style.name == "SectionHeader":
                    flowable._report_anchor = f"sec-{section.key}"
                    break
            elements.extend(section_elements)
        except Exception as e:
            section_failures.append(str(e))
            elements.append(
                Paragraph(
                    f"<i>{section.title} failed: {xml_escape(str(e))}</i>",
                    STYLES["BodyText"],
                )
            )

    # Build
    if section_failures:
        raise RuntimeError(
            "Report incomplete; output was not published: " + "; ".join(section_failures)
        )
    doc.build(elements)
    return output_path


# =============================================================================
# GUI INTEGRATION HELPER
# =============================================================================


def generate_pdf_from_gui(
    records: list[dict[str, Any]],
    output_path: str,
    config: ConfigDict,
    engine: Any = None,
    filtered: list[dict[str, Any]] | None = None,
) -> tuple[bool, str]:
    """Generate a professional PDF report for GUI integration."""
    try:
        path = generate_ombudsman_pdf(records, output_path, dict(config), engine, filtered)
        return True, f"Professional PDF report generated:\n{path}"
    except Exception as e:
        return False, f"Failed to generate PDF:\n{e}"
