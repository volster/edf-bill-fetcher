"""Tests for the reconciliation cross-source sheet writer."""

from __future__ import annotations

import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from edf_collector import write_reconciliation_sheet

_SAP_CONTRACT = [
    {
        "Contract From": "2024-05-14",
        "Contract To": "2024-06-30",
        "Product Code": "PRD_FXD24",
        "Product Description": "Fixed Online 2 Year",
        "Contract Reason": "New Sales",
        "Set Up By": "agent01",
        "Notes": "",
        "Cancelled Flag": "",
        "Source File": "Contract-and-Product-Change-History.pdf",
    },
    {
        "Contract From": "2024-07-01",
        "Contract To": "2024-07-31",
        "Product Code": "PRD_FREE",
        "Product Description": "Freedom",
        "Contract Reason": "Tariff Switch",
        "Set Up By": "agent02",
        "Notes": "",
        "Cancelled Flag": "",
        "Source File": "Contract-and-Product-Change-History.pdf",
    },
]


_INFERRED_CONTRACT = pd.DataFrame(
    [
        {
            "Contract From": "2024-05-14",
            "Contract To": "2024-06-30",
            "Product Code": "PRD_FXD24",
            "Product Description": "Fixed Online 2 Year",
            "Contract Reason": "Inferred from invoice body",
            "Set Up By": "N/A",
            "Notes": "",
            "Cancelled Flag": "",
            "Source File": "i-T12345.pdf",
        }
    ]
)


_SAP_METER = [
    {
        "Scheduled Read Date": "2024-05-14",
        "Meter Read Date": "2024-05-14",
        "Reading (kWh)": "1234.5000",
        "Read Type": "Periodic scheduled",
        "Read Source": "Metering System",
        "Read Status": "Posted",
        "Meter Read Reason": "Move-In",
        "Register": "01",
        "Source File": "Meter-Read-History.pdf",
    }
]


_INFERRED_METER = pd.DataFrame(
    [
        {
            "Scheduled Read Date": "2024-05-14",
            "Meter Read Date": "2024-05-14",
            "Reading (kWh)": "1234.5000",
            "Read Type": "A",
            "Read Source": "Customer",
            "Read Status": "Posted",
            "Meter Read Reason": "Move-In",
            "Register": "01",
        }
    ]
)


_SAP_FINANCIAL = [
    {
        "Document No.": "9000012345",
        "Item": "001",
        "Document Date": "2024-05-14",
        "Posting Date": "2024-05-14",
        "Net Due Date": "2024-05-21",
        "Main Transaction": "Credit Memo",
        "Sub Transaction": "Reversal",
        "Transaction Text": "Reversal Inv T12345",
        "Amount": "1347.96",
        "Clearing Status": "Not Cleared",
        "Clearing Document": "",
        "Clearing Date": "",
        "Clearing Reason": "",
        "Document Type": "CM",
        "Document Type Description": "Credit Memo",
        "Source File": "Financial-Transactions.pdf",
    }
]


_EVIDENCE_DF = pd.DataFrame(
    [
        {
            "Date": "14/05/2024",
            "Invoice #": "T12345",
            "Period From": "14/05/2024",
            "Period To": "30/06/2024",
            "Amount (£)": 1347.96,
            "Entry Type": "Charge",
            "Logic Used": "New Invoice Format",
        },
    ]
)


def _section_row_indexes(ws: Worksheet) -> dict[str, tuple[int, int]]:
    """Return dict  {section_label: (start_row_excl_banner, end_row_incl)}.

    The section banners are written with a leading/trailing ``■`` visual
    marker (see ``_section_banner``).  Pre-fix the banners used a
    ``"== {text} =="`` literal that openpyxl was silently serialising
    as a worksheet formula (because the value starts with ``=``),
    triggering Excel's "Removed Records: Formula" recover prompt.  The
    helper therefore accepts BOTH the legacy ``"=="``-delimited form
    (for backwards compatibility with any pre-fix snapshots still in
    flight) and the current ``■ {text} ■`` form.
    """
    rows = {}
    for row in ws.iter_rows():
        v0 = row[0].value
        if isinstance(v0, str):
            stripped = v0.strip()
            if stripped.startswith("==") and stripped.endswith("=="):
                label = stripped.strip("=").strip()
                rows[label] = (row[0].row, -1)
            elif stripped.startswith("\u25a0") and stripped.endswith("\u25a0"):
                label = stripped.replace("\u25a0", "").strip()
                rows[label] = (row[0].row, -1)
    out = {}
    keys = list(rows.keys())
    for i, key in enumerate(keys):
        start = rows[key][0] + 1
        end = rows[keys[i + 1]][0] - 1 if i + 1 < len(keys) else ws.max_row
        out[key] = (start, end)
    return out


def test_three_banners_present() -> None:
    wb = Workbook()
    ws = wb.active
    write_reconciliation_sheet(
        ws,
        _SAP_CONTRACT,
        _INFERRED_CONTRACT,
        _SAP_METER,
        _INFERRED_METER,
        _SAP_FINANCIAL,
        _EVIDENCE_DF,
        account="A-31105244",
    )
    coords = _section_row_indexes(ws)
    assert "Contract Reconciliation" in coords
    assert "Meter Read Reconciliation" in coords
    assert "Financial Reconciliation" in coords


def test_matched_contract_row_has_hyperlink() -> None:
    wb = Workbook()
    ws = wb.active
    write_reconciliation_sheet(
        ws,
        _SAP_CONTRACT,
        _INFERRED_CONTRACT,
        _SAP_METER,
        _INFERRED_METER,
        _SAP_FINANCIAL,
        _EVIDENCE_DF,
        account="A-31105244",
    )
    coords = _section_row_indexes(ws)
    start, end = coords["Contract Reconciliation"]
    body_first = start + 1
    for r in range(body_first, end + 1):
        v = ws.cell(row=r, column=1).value
        if v == "Matched":
            cell = ws.cell(row=r, column=8)
            assert cell.hyperlink is not None
            assert cell.hyperlink.location is not None
            assert cell.hyperlink.location.startswith(
                "'Contract History'!A"
            ) or cell.hyperlink.location.startswith("'EDF Evidence Report'!A")
            return
    raise AssertionError("No Matched contract row found")


def test_matched_meter_row_has_hyperlink() -> None:
    wb = Workbook()
    ws = wb.active
    write_reconciliation_sheet(
        ws,
        _SAP_CONTRACT,
        _INFERRED_CONTRACT,
        _SAP_METER,
        _INFERRED_METER,
        _SAP_FINANCIAL,
        _EVIDENCE_DF,
        account="A-31105244",
    )
    coords = _section_row_indexes(ws)
    start, end = coords["Meter Read Reconciliation"]
    body_first = start + 1
    for r in range(body_first, end + 1):
        v = ws.cell(row=r, column=1).value
        if v == "Matched":
            cell = ws.cell(row=r, column=8)
            assert cell.hyperlink is not None
            assert cell.hyperlink.location is not None
            # Meter-read back to SAP meter or inferred meter: must reference the SAP sheet name.
            assert "Meter" in cell.hyperlink.location
            return
    raise AssertionError("No Matched meter row found")


def test_matched_financial_row_has_hyperlink_evidence_or_sap() -> None:
    wb = Workbook()
    ws = wb.active
    write_reconciliation_sheet(
        ws,
        _SAP_CONTRACT,
        _INFERRED_CONTRACT,
        _SAP_METER,
        _INFERRED_METER,
        _SAP_FINANCIAL,
        _EVIDENCE_DF,
        account="A-31105244",
    )
    coords = _section_row_indexes(ws)
    start, end = coords["Financial Reconciliation"]
    body_first = start + 1
    for r in range(body_first, end + 1):
        v = ws.cell(row=r, column=1).value
        if v == "Matched":
            cell = ws.cell(row=r, column=8)
            assert cell.hyperlink is not None
            assert cell.hyperlink.location is not None
            assert (
                "EDF Evidence Report" in cell.hyperlink.location
                or "SAP Financial" in cell.hyperlink.location
            )
            return
    raise AssertionError("No Matched financial row found")


def test_unmatched_input_emits_missing_status() -> None:
    wb = Workbook()
    ws = wb.active
    write_reconciliation_sheet(
        ws,
        _SAP_CONTRACT,
        _INFERRED_CONTRACT,
        _SAP_METER,
        _INFERRED_METER,
        _SAP_FINANCIAL,
        _EVIDENCE_DF,
        account="A-31105244",
    )
    coords = _section_row_indexes(ws)
    start, end = coords["Contract Reconciliation"]
    body_first = start + 1
    statuses = [ws.cell(row=r, column=1).value for r in range(body_first, end + 1)]
    # The second SAP contract row is 0-2024 matching the product code PRD_FREE
    # but no inferred row has that product code, so it should be marked.
    assert "Missing in Inferred" in statuses or "Missing in SAP" in statuses


def test_discrepancy_emitted_when_amount_differs() -> None:
    # SAP amount £1347.96 vs evidence amount £1348.16 — within ±£0.50 but
    # above the strict £0.01 equality gate → Discrepancy status emitted.
    sap_fin = [
        {
            "Document No.": "9000012345",
            "Item": "001",
            "Document Date": "2024-05-14",
            "Posting Date": "2024-05-14",
            "Net Due Date": "2024-05-21",
            "Main Transaction": "Credit Memo",
            "Sub Transaction": "Reversal",
            "Transaction Text": "Reversal Inv T12345",
            "Amount": "1347.96",
            "Clearing Status": "Not Cleared",
            "Clearing Document": "",
            "Clearing Date": "",
            "Clearing Reason": "",
            "Document Type": "CM",
            "Document Type Description": "Credit Memo",
            "Source File": "Financial-Transactions.pdf",
        }
    ]
    ev = pd.DataFrame(
        [
            {
                "Date": "14/05/2024",
                "Invoice #": "T12345",
                "Period From": "14/05/2024",
                "Period To": "30/06/2024",
                "Amount (£)": 1348.16,
                "Entry Type": "Charge",
                "Logic Used": "New Invoice Format",
            },
        ]
    )
    wb = Workbook()
    ws = wb.active
    write_reconciliation_sheet(
        ws,
        _SAP_CONTRACT,
        _INFERRED_CONTRACT,
        _SAP_METER,
        _INFERRED_METER,
        sap_fin,
        ev,
        account="A-31105244",
    )
    coords = _section_row_indexes(ws)
    start, end = coords["Financial Reconciliation"]
    found_disc = False
    for r in range(start + 1, end + 1):
        v = ws.cell(row=r, column=1).value
        if v == "Discrepancy":
            found_disc = True
            break
    assert found_disc, "expected a Discrepancy row given amount differs > £0.01 but ≤ £0.50"


def test_section_banners_are_not_serialised_as_formulas(tmp_path: object) -> None:
    """Regression: ``_section_banner`` used to write ``"== Contract
    Reconciliation =="`` as the cell value.  openpyxl sees any string
    that starts with ``=`` and serialises it as a worksheet formula
    (``<f>...</f>``).  Excel then errors on file-open with:

        "We found a problem with some content ... do you want us to
        try and recover ... " / "Removed Records: Formula from
        /xl/worksheets/sheetN.xml"

    Pre-fix this corrupted every Reconciliation sheet, prompting the
    user each open even though the sheet content itself was intact.

    The fix changed the banner visual marker from ``"== {text} =="``
    to ``"\u25a0 {text} \u25a0"`` so the cell value no longer begins
    with ``=``.  This test unzips the saved workbook and asserts ZERO
    ``<f>`` (formula) tags across ALL worksheets so a regression in
    any writer surfaces fast.
    """
    import zipfile

    wb = Workbook()
    ws = wb.active
    write_reconciliation_sheet(
        ws,
        _SAP_CONTRACT,
        _INFERRED_CONTRACT,
        _SAP_METER,
        _INFERRED_METER,
        _SAP_FINANCIAL,
        _EVIDENCE_DF,
        account="A-31105244",
    )
    out = tmp_path / "recon.xlsx"  # type: ignore[operator]
    wb.save(out)

    with zipfile.ZipFile(out) as z:
        for n in z.namelist():
            if not n.startswith("xl/worksheets/sheet") or not n.endswith(".xml"):
                continue
            xml = z.read(n).decode("utf-8", errors="replace")
            nf = xml.count("<f>") + xml.count("<f ") + xml.count("<f/")
            assert nf == 0, (
                f"{n} contains {nf} <f> formula tag(s) -- banner cells must "
                f"be inline strings, not formulas (Excel 'Removed Records: "
                f"Formula' corruption)."
            )
