"""Generate a synthetic EDF KI-style bill PDF as an integration-test fixture.

This script exists so the test_pdf_to_records_no_exception /
test_full_pipeline_creates_pdf_and_xlsx tests in
test_integration_pipeline.py have a real, deliberately-constructed
EDF-style bill PDF to walk the pipeline on. The synthetic PDF:

  * is NOT based on any real EDF customer's bill — call numbers /
    addresses / amounts / meters use FAFO (For Any Fictional Output)
    placeholder data so no PII can leak;
  * mirrors the format of an EDF KI ("Key Insights") invoice: header
    with emergency numbers, supply address, "Your VAT invoice"
    section, charges table, "Total charges for this period GBP X debit",
    "Current balance GBP X debit", and a 2nd detail page with tariff and
    meter info;
  * is reproducible — no random elements, so the file's hash and
    expected parsed-record set stay stable across runs.

Usage
-----
    python tests/fixtures/generate_bill_fixture.py [output_path]

Default output_path: ``output/bill_fixture.pdf`` (so it can be reviewed
before being moved into tests/fixtures/ for the integration test).
"""

from __future__ import annotations

import sys
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# Synthetic, FAKE, no-PII data — see module docstring.
ACCOUNT_NUMBER = "A-0000000"
INVOICE_NUMBER = "KI-0000000-0000"
BIN = "2010-FAFE-FAKE-0000"
SUPPLY_NAME = "Synthetic Test Site"
SUPPLY_ADDR_1 = "1 Test Lane"
SUPPLY_ADDR_2 = "Sample Village"
SUPPLY_POSTCODE = "TST 0ST"
BILL_DATE = "1 March 2026"
PERIOD_FROM = "1 February 2026"
PERIOD_TO = "28 February 2026"
PERIOD_DAYS = 28
PAY_BY = "10 March 2026"
TARIFF_NAME = "Freedom Fake"
ELECTRICITY_USED_KWH = 250.000
UNIT_RATE_PKWH = 24.480
STANDING_CHARGE_PDAY = 75.250
ELECTRICITY_NET = "85.68"
CCL = "1.11"
VAT = "4.34"
TOTAL_CHARGES = "240.50"
CURRENT_BALANCE = "240.50"


def build(output_path: Path) -> None:
    """Build the synthetic bill PDF in-place at output_path.

    Reads only module-level constants; deterministic across runs.
    """
    styles = getSampleStyleSheet()
    body = styles["BodyText"]
    body.fontSize = 9
    body.leading = 12
    h1 = ParagraphStyle(
        "H1",
        parent=styles["Heading1"],
        fontSize=14,
        spaceAfter=6,
        alignment=1,
    )

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
        title="Synthetic EDF Bill Fixture",
    )

    story = []

    # ----- Page 1: Smell-gas / power-cut emergency block + invoice body -----
    story.append(Paragraph("Smell Gas - Immediately call 0800 111 999 (24hrs)", body))
    story.append(
        Paragraph(
            "Power cut? Call 105 to get through to your electricity distributor",
            body,
        )
    )
    story.append(Paragraph("National Grid 0800 096 3080", body))
    story.append(Spacer(1, 0.5 * cm))

    story.append(Paragraph(f"{SUPPLY_NAME} Supply address: {SUPPLY_NAME},", body))
    story.append(Paragraph(SUPPLY_ADDR_1, body))
    story.append(Paragraph(SUPPLY_ADDR_2, body))
    story.append(Paragraph(SUPPLY_POSTCODE, body))
    story.append(Spacer(1, 0.5 * cm))

    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph("Your VAT invoice", h1))

    story.append(Paragraph(f"Invoice number: {INVOICE_NUMBER}", body))
    story.append(Paragraph(f"Account number: {ACCOUNT_NUMBER}", body))
    story.append(Paragraph(f"Date issued: {BILL_DATE}", body))
    story.append(Spacer(1, 0.3 * cm))

    # Use the pound sign directly — matches the production parser's
    # ``AMOUNT_PATTERNS`` regex anchored on the GBP glyph.
    gp = "£"
    story.append(
        Paragraph(
            f"Your charges: {PERIOD_FROM} - {PERIOD_TO}",
            body,
        )
    )

    charges_table = Table(
        [
            ["", "Net charges", "CCL", "VAT", "(Inclusive) Total"],
            [
                "Electricity",
                f"-{gp}{ELECTRICITY_NET}",
                f"{gp}{CCL}",
                f"-{gp}{VAT}",
                f"-{gp}{TOTAL_CHARGES}",
            ],
        ],
    )
    charges_table.setStyle(
        TableStyle(
            [
                ("FONT", (0, 0), (-1, -1), "Helvetica", 9),
                ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 9),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ("ALIGN", (0, 0), (0, -1), "LEFT"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ]
        )
    )
    story.append(charges_table)
    story.append(Spacer(1, 0.3 * cm))

    story.append(
        Paragraph(
            f"Total charges for this period {gp}{TOTAL_CHARGES} credit",
            body,
        )
    )
    story.append(
        Paragraph(
            f"Current balance {gp}{CURRENT_BALANCE} debit",
            body,
        )
    )
    story.append(Paragraph(f"Please pay this invoice by {PAY_BY}.", body))

    # ----- Page 2: "Your charges in detail" --------------------------------
    story.append(Spacer(1, 1.5 * cm))
    story.append(Paragraph("Your charges in detail", h1))

    story.append(Paragraph("About your tariff", body))
    story.append(Paragraph("Electricity Supply Code: 00 000 K 00", body))
    story.append(Spacer(1, 0.3 * cm))

    story.append(Paragraph("Electricity", body))
    story.append(
        Paragraph(
            f"Supply address: {SUPPLY_NAME}, {SUPPLY_ADDR_1}, {SUPPLY_ADDR_2}, {SUPPLY_POSTCODE}",
            body,
        )
    )
    story.append(Paragraph(f"Tariff name {TARIFF_NAME}", body))
    story.append(Paragraph("Payment type Non-Direct Debit", body))
    story.append(Paragraph("Rota Disconnections Alpha Identifier: AA", body))
    story.append(
        Paragraph(
            f"Contract end date {BILL_DATE}",
            body,
        )
    )
    story.append(
        Paragraph(
            f"{TARIFF_NAME} ({PERIOD_FROM} - {PERIOD_TO}) Early exit fee No",
            body,
        )
    )
    story.append(
        Paragraph(
            f"Estimated yearly usage {int(ELECTRICITY_USED_KWH * 12)}.000 kWh",
            body,
        )
    )
    story.append(
        Paragraph(
            "Electricity charges for meter M0000000",
            body,
        )
    )
    story.append(Spacer(1, 0.3 * cm))

    story.append(
        Paragraph(
            f"{PERIOD_FROM} 99999.999 Estimated reading Any PP0 Supplier "
            f"Certificate logged with TR-0000 calculations.",
            body,
        )
    )
    story.append(
        Paragraph(
            f"{PERIOD_TO} 99999.999 Estimated reading VAT charge at reduced rate",
            body,
        )
    )
    story.append(
        Paragraph(
            "Any VAT declaration logged with us has been considered.",
            body,
        )
    )
    story.append(
        Paragraph(
            f"{ELECTRICITY_USED_KWH:.3f} kWh",
            body,
        )
    )
    story.append(
        Paragraph(
            f"Electricity used {UNIT_RATE_PKWH:.3f}p/kWh {gp}{ELECTRICITY_NET}",
            body,
        )
    )
    story.append(
        Paragraph(
            f"Standing charge {PERIOD_DAYS} days @ {STANDING_CHARGE_PDAY:.3f}p/day {gp}{CCL}",
            body,
        )
    )

    doc.build(story)


def main() -> int:
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("output/bill_fixture.pdf")
    out.parent.mkdir(parents=True, exist_ok=True)
    build(out)
    print(f"Wrote {out} ({out.stat().st_size} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
