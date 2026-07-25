"""Unit tests for the SAP Back-billing analyser (spec §9.1).

Covers ``parse_sap_financial_transactions`` ( widened parser),
``detect_sap_back_billing_events`` (clustering), and
``match_sap_events_to_edf`` (fuzzy match + confidence bands).
"""

from __future__ import annotations

import pandas as pd

from edf_collector import (
    SapBackBillingEvent,
    detect_sap_back_billing_events,
    match_sap_events_to_edf,
    parse_sap_financial_transactions,
)

# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------


def _synthetic_sap_csv(rows: list[dict]) -> str:
    """Build a minimal CSV body whose header row matches the source PDF."""
    cols = [
        "Kraken ID",
        "SAP Account Number",
        "Business Partner",
        "Account Determination ID",
        "Contract",
        "Fuel Type",
        "Document No.",
        "Item",
        "Sub Item",
        "Payment Method",
        "Document Date",
        "Posting Date ",
        "Net Due Date",
        "Clearing Status",
        "Main Transactions",
        "Sub Transactions",
        "Transaction Text",
        "Amount",
        "Down Payment Flag",
        "Statistical Key Flag",
        "Clearing Document",
        "Clearing Date",
        "Clearing Reason",
        "Clearing Posting Date",
        "Clearing Amount",
        "Restriction",
        "Document Type",
        "Document Type Description",
        "Tax Code",
        "Tax Code Description",
        "G/L Account",
        "G/L Description",
        "Deferral Date",
    ]
    lines = ['"' + '","'.join(cols) + '"']
    for r in rows:
        lines.append('"' + '","'.join(str(r.get(c, "")) for c in cols) + '"')
    return "\n".join(lines)


def test_parse_sap_financial_transactions_now_returns_26_columns() -> None:
    """Spec §3.1: parser emits 26-key dicts (16 historical + 10 added)."""
    sample = {
        "Document No.": "551000421040",
        "Item": "1",
        "Sub Item": "0",
        "Document Date": "18-07-2016",
        "Posting Date ": "18-07-2016",
        "Net Due Date": "21-07-2016",
        "Main Transactions": "0100",
        "Sub Transactions": "0020",
        "Transaction Text": "Dr- Consum Billing Receivable",
        "Amount": "436",
        "Clearing Status": "Cleared Item",
        "Clearing Document": "376001212905",
        "Clearing Date": "26-03-2020",
        "Clearing Reason": "Automatic Clearing",
        "Clearing Posting Date": "26-03-2020",
        "Clearing Amount": "436",
        "Statistical Key Flag": "",
        "Down Payment Flag": "No",
        "Document Type": "IN",
        "Document Type Description": "Energy Invoicing",
        "Tax Code": "A4",
        "Tax Code Description": "Donations or payment for equity funds",
        "G/L Account": "0000210251",
        "G/L Description": "Billed Debtor SME Elec",
        "Contract": "2011040650",
        "Deferral Date": "",
        "Restriction": "No restriction",
    }
    rows = parse_sap_financial_transactions(_synthetic_sap_csv([sample]))
    assert len(rows) == 1
    r = rows[0]
    assert len(r) == 26, f"expected 26 keys, got {len(r)}: {list(r.keys())}"
    # Sanity-check the new keys are populated
    assert r["Contract"] == "2011040650"
    assert r["Sub Item"] == "0"
    assert r["Clearing Posting Date"] == "2020-03-26"  # iso-normalised
    assert r["Clearing Amount"] == "436"
    assert r["Statistical Key Flag"] == ""
    assert r["Tax Code"] == "A4"
    assert r["Tax Code Description"] == "Donations or payment for equity funds"
    assert r["G/L Account"] == "0000210251"
    assert r["G/L Description"] == "Billed Debtor SME Elec"
    assert r["Deferral Date"] == ""


# ---------------------------------------------------------------------------
# Detect events
# ---------------------------------------------------------------------------


def _mkrow(
    doc_no: str = "DOC1",
    item: str = "1",
    posting: str = "2020-01-01",
    amount: str = "100",
    text: str = "Dr- Consum Billing Receivable",
    clear_doc: str = "",
    clear_date: str = "",
    clear_reason: str = "",
    doc_type: str = "IN",
    stat_flag: str = "",
) -> dict:
    return {
        "Document No.": doc_no,
        "Item": item,
        "Document Date": "",
        "Posting Date": posting,
        "Net Due Date": "",
        "Main Transaction": "0100",
        "Sub Transaction": "0020",
        "Transaction Text": text,
        "Amount": amount,
        "Clearing Status": "Cleared Item",
        "Clearing Document": clear_doc,
        "Clearing Date": clear_date,
        "Clearing Reason": clear_reason,
        "Document Type": doc_type,
        "Document Type Description": "",
        "Source File": "test.pdf",
        "Contract": "",
        "Sub Item": "0",
        "Clearing Posting Date": "",
        "Clearing Amount": "",
        "Statistical Key Flag": stat_flag,
        "Tax Code": "",
        "Tax Code Description": "",
        "G/L Account": "",
        "G/L Description": "",
        "Deferral Date": "",
    }


def test_detect_sap_back_billing_events_filters_debt_management() -> None:
    """A row flagged ``Installment Plan Item`` must not appear in any event."""
    rows = [
        _mkrow(
            doc_no="A",
            clear_doc="C1",
            clear_date="2020-01-01",
            amount="100",
            stat_flag="Installment Plan Item",
        ),
        _mkrow(
            doc_no="B",
            clear_doc="C1",
            clear_date="2020-01-01",
            amount="100",
            stat_flag="Installment Plan Item",
        ),
        _mkrow(
            doc_no="C",
            clear_doc="C1",
            clear_date="2020-01-01",
            amount="100",
            stat_flag="Installment Plan Item",
        ),
        _mkrow(
            doc_no="D",
            clear_doc="C1",
            clear_date="2020-01-01",
            amount="100",
            stat_flag="Installment Plan Item",
        ),
    ]
    events = detect_sap_back_billing_events(rows)
    assert events == [], "debt-management rows should be filtered before clustering"


def test_detect_sap_back_billing_events_groups_by_clearing_document() -> None:
    """Six rows on the same Clearing Document produce ONE event of size 6."""
    rows = [
        _mkrow(
            doc_no=f"D{i}",
            clear_doc="C1",
            clear_date="2020-03-01",
            amount="50",
            posting=f"2020-03-{i + 1:02d}",
        )
        for i in range(6)
    ]
    events = detect_sap_back_billing_events(rows)
    assert len(events) == 1, f"expected 1 event, got {len(events)}"
    ev = events[0]
    assert ev.clearing_doc == "C1"
    assert len(ev.rows) == 6
    assert ev.net_amount == 300.0


def test_detect_sap_back_billing_events_min_cluster_size() -> None:
    """A 3-row cluster falls below the default min_cluster_size of 4."""
    rows = [
        _mkrow(doc_no=f"D{i}", clear_doc="C1", clear_date="2020-03-01", amount="50")
        for i in range(3)
    ]
    events = detect_sap_back_billing_events(rows)
    assert events == [], "cluster below min_cluster_size should be filtered"


def test_detect_sap_back_billing_events_net_zero_is_back_billing() -> None:
    """A £0-net cluster with a Cr-Credit row signals back-billing."""
    rows = [
        _mkrow(
            doc_no="D1",
            clear_doc="C1",
            clear_date="2020-03-01",
            amount="6108.66",
            text="Dr- Consum Billing Receivable",
            posting="2020-10-01",
        ),
        _mkrow(
            doc_no="D1",
            clear_doc="C1",
            clear_date="2020-03-01",
            amount="-6108.66",
            text="Cr- Credit for Consum Billing",
            posting="2020-10-01",
        ),
        _mkrow(
            doc_no="D2", clear_doc="C1", clear_date="2020-03-01", amount="100", posting="2020-10-01"
        ),
        _mkrow(
            doc_no="D3",
            clear_doc="C1",
            clear_date="2020-03-01",
            amount="-100",
            text="Cr- Credit for Consum Billing",
            posting="2020-10-01",
        ),
    ]
    events = detect_sap_back_billing_events(rows)
    assert len(events) == 1
    ev = events[0]
    assert abs(ev.net_amount) < 0.01
    assert ev.has_credit_for_consum_billing is True


def test_detect_sap_back_billing_events_sorts_by_clearing_date_ascending() -> None:
    """Events return in ascending Clearing Date order."""
    rows = [
        _mkrow(clear_doc="LATE", clear_date="2023-08-09", amount="10", posting="2023-08-09"),
        _mkrow(clear_doc="LATE", clear_date="2023-08-09", amount="10", posting="2023-08-09"),
        _mkrow(clear_doc="LATE", clear_date="2023-08-09", amount="10", posting="2023-08-09"),
        _mkrow(clear_doc="LATE", clear_date="2023-08-09", amount="10", posting="2023-08-09"),
        _mkrow(clear_doc="EARLY", clear_date="2016-11-08", amount="10", posting="2016-11-08"),
        _mkrow(clear_doc="EARLY", clear_date="2016-11-08", amount="10", posting="2016-11-08"),
        _mkrow(clear_doc="EARLY", clear_date="2016-11-08", amount="10", posting="2016-11-08"),
        _mkrow(clear_doc="EARLY", clear_date="2016-11-08", amount="10", posting="2016-11-08"),
    ]
    events = detect_sap_back_billing_events(rows)
    assert len(events) == 2
    assert events[0].clearing_doc == "EARLY"
    assert events[1].clearing_doc == "LATE"


# ---------------------------------------------------------------------------
# Match to EDF
# ---------------------------------------------------------------------------


def _mk_ev(
    clearing_doc: str = "C1",
    clearing_date: str = "2023-08-09",
    net_amount: float = 0.0,
    rows_amounts: list[float] | None = None,
    has_credit: bool = True,
) -> SapBackBillingEvent:
    sub_rows = []
    if rows_amounts:
        for i, amt in enumerate(rows_amounts):
            sub_rows.append(
                _mkrow(
                    doc_no=f"D{i}",
                    item=str(i),
                    clear_doc=clearing_doc,
                    clear_date=clearing_date,
                    posting=clearing_date,
                    amount=str(amt),
                    text="Cr- Credit for Consum Billing"
                    if amt < 0
                    else "Dr- Consum Billing Receivable",
                )
            )
    else:
        sub_rows = [
            _mkrow(
                doc_no="D0",
                clear_doc=clearing_doc,
                clear_date=clearing_date,
                posting=clearing_date,
                amount="0",
                text="Dr- Consum Billing Receivable",
            ),
            _mkrow(
                doc_no="D1",
                clear_doc=clearing_doc,
                clear_date=clearing_date,
                posting=clearing_date,
                amount="0",
                text="Dr- Consum Billing Receivable",
            ),
            _mkrow(
                doc_no="D2",
                clear_doc=clearing_doc,
                clear_date=clearing_date,
                posting=clearing_date,
                amount="0",
                text="Dr- Consum Billing Receivable",
            ),
            _mkrow(
                doc_no="D3",
                clear_doc=clearing_doc,
                clear_date=clearing_date,
                posting=clearing_date,
                amount="0",
                text="Dr- Consum Billing Receivable",
            ),
        ]
    return SapBackBillingEvent(
        clearing_doc=clearing_doc,
        clearing_date=pd.Timestamp(clearing_date),
        clearing_reason="Account Maintenance",
        rows=sub_rows,
        net_amount=net_amount,
        has_credit_for_consum_billing=has_credit,
    )


def _mk_edf(
    invoice: str = "T78",
    period_from: str = "02/10/2020",
    period_to: str = "09/08/2023",
    amount: float = 28192.35,
) -> dict:
    # Use UK DD/MM/YYYY strings — that's the form EDF Evidence Report
    # rows carry their Period From/To fields in when match_sap_events_to_edf
    # is called via export_to_excel (dfc.to_dict(orient="records")).
    return {
        "Invoice #": invoice,
        "Period From": period_from,
        "Period To": period_to,
        "Amount (£)": amount,
        "Date": period_to,
    }


def test_match_sap_events_to_edf_exact_day_high_conf() -> None:
    """Clearing Date inside EDF period + amount within 5% → High."""
    ev = _mk_ev(
        clearing_date="2023-08-09",
        net_amount=28000.0,
    )
    ev.rows = [
        _mkrow(
            amount="28000",
            clear_doc=ev.clearing_doc,
            clear_date="2023-08-09",
            posting="2023-08-09",
            text="Dr- Consum Billing Receivable",
        ),
        _mkrow(
            amount="0", clear_doc=ev.clearing_doc, clear_date="2023-08-09", posting="2023-08-09"
        ),
        _mkrow(
            amount="0", clear_doc=ev.clearing_doc, clear_date="2023-08-09", posting="2023-08-09"
        ),
        _mkrow(
            amount="0", clear_doc=ev.clearing_doc, clear_date="2023-08-09", posting="2023-08-09"
        ),
    ]
    edf = [
        _mk_edf(invoice="T78", period_from="02/10/2020", period_to="09/08/2023", amount=28192.35)
    ]
    matches = match_sap_events_to_edf([ev], edf)
    assert len(matches) == 1
    assert matches[0].confidence_band == "High"
    assert matches[0].edf_record["Invoice #"] == "T78"
    assert matches[0].date_delta_days == 0


def test_match_sap_events_to_edf_three_day_medium_conf() -> None:
    """3-day delta + amount within 25% → Medium."""
    ev = _mk_ev(
        clearing_date="2023-08-09",
        net_amount=25000.0,
    )
    ev.rows = [
        _mkrow(
            amount="25000",
            clear_doc=ev.clearing_doc,
            clear_date="2023-08-09",
            posting="2023-08-09",
            text="Dr- Consum Billing Receivable",
        ),
        _mkrow(
            amount="0", clear_doc=ev.clearing_doc, clear_date="2023-08-09", posting="2023-08-09"
        ),
        _mkrow(
            amount="0", clear_doc=ev.clearing_doc, clear_date="2023-08-09", posting="2023-08-09"
        ),
        _mkrow(
            amount="0", clear_doc=ev.clearing_doc, clear_date="2023-08-09", posting="2023-08-09"
        ),
    ]
    edf = [
        _mk_edf(invoice="T78", period_from="02/10/2020", period_to="06/08/2023", amount=28192.35)
    ]
    matches = match_sap_events_to_edf([ev], edf)
    assert len(matches) == 1, f"got {len(matches)} matches"
    assert matches[0].confidence_band == "Medium"
    assert matches[0].date_delta_days == 3


def test_match_sap_events_to_edf_no_match_omitted() -> None:
    """A SAP event 60 days off any EDF period yields no match."""
    ev = _mk_ev(clearing_date="2023-08-09", net_amount=1.0)
    ev.rows = [
        _mkrow(
            amount="1",
            clear_doc=ev.clearing_doc,
            clear_date="2023-08-09",
            posting="2023-08-09",
            text="Dr- Consum Billing Receivable",
        ),
        _mkrow(
            amount="0", clear_doc=ev.clearing_doc, clear_date="2023-08-09", posting="2023-08-09"
        ),
        _mkrow(
            amount="0", clear_doc=ev.clearing_doc, clear_date="2023-08-09", posting="2023-08-09"
        ),
        _mkrow(
            amount="0", clear_doc=ev.clearing_doc, clear_date="2023-08-09", posting="2023-08-09"
        ),
    ]
    edf = [_mk_edf(invoice="T999", period_from="01/01/2010", period_to="31/01/2010", amount=100.0)]
    matches = match_sap_events_to_edf([ev], edf)
    assert matches == [], f"expected no match, got {matches}"


def test_match_sap_events_to_edf_net_zero_gross_amount_match_high() -> None:
    """Spec §3.3: net-zero cluster with a row gross-amount == EDF amount → High."""
    ev = _mk_ev(
        clearing_date="2023-08-09",
        net_amount=0.0,
        rows_amounts=[23961.35, -23961.35, 0.0, 0.0],
        has_credit=True,
    )
    ev = SapBackBillingEvent(
        clearing_doc=ev.clearing_doc,
        clearing_date=ev.clearing_date,
        clearing_reason="Account Maintenance",
        rows=[
            _mkrow(
                doc_no=f"D{i}",
                amount=str(amt),
                clear_doc=ev.clearing_doc,
                clear_date="2023-08-09",
                posting="2023-08-09",
                text="Cr- Credit for Consum Billing"
                if amt < 0
                else "Dr- Consum Billing Receivable",
            )
            for i, amt in enumerate([23961.35, -23961.35, 0.0, 0.0])
        ],
        net_amount=0.0,
        has_credit_for_consum_billing=True,
    )
    edf = [
        _mk_edf(invoice="T78", period_from="02/10/2020", period_to="09/08/2023", amount=23961.35)
    ]
    matches = match_sap_events_to_edf([ev], edf)
    assert len(matches) == 1
    assert matches[0].confidence_band == "High", (
        f"net-zero cluster with exact gross match should be High, "
        f"got {matches[0].confidence_band} (score={matches[0].confidence_score})"
    )


# ---------------------------------------------------------------------------
# PR #2 — Option C: require amount match for Medium+, gate in-span on amount
# Spec §3.1 (issue 1): matches that scored "Medium" purely because a
# SAP clearing date happened to fall inside a wide EDF invoice period
# (with zero amount correspondence) must be demoted to Low or dropped.
# ---------------------------------------------------------------------------


def test_matcher_demotes_in_span_no_amount_to_low() -> None:
    """Clearing date inside EDF period but amount wildly off → Low (not Medium).

    Spec §3.1 — Option C: previously this row would have scored Medium
    purely from the in-span 50-point bonus with zero amount correspondence;
    the new gate caps it at Low.
    """
    ev = _mk_ev(clearing_date="2019-09-03", net_amount=-831.45)
    edf = [
        _mk_edf(
            invoice="T-001",
            period_from="14/06/2018",
            period_to="04/09/2019",
            amount=20828.82,
        )
    ]
    matches = match_sap_events_to_edf([ev], edf)
    band = matches[0].confidence_band if matches else None
    assert band in ("Low", None), (
        f"in-span no-amount must be Low or dropped, got {band}"
    )


def test_matcher_in_span_with_amount_within_5pct_stays_high() -> None:
    """Date in-span + amount within 5% keeps the High band."""
    ev = _mk_ev(clearing_date="2024-01-15", net_amount=100.00)
    edf = [
        _mk_edf(
            invoice="T-002",
            period_from="01/01/2024",
            period_to="31/01/2024",
            amount=100.00,
        )
    ]
    matches = match_sap_events_to_edf([ev], edf)
    assert matches, "expected a match"
    assert matches[0].confidence_band == "High", matches[0].confidence_band


def test_matcher_in_span_with_amount_within_25pct_caps_at_medium() -> None:
    """Date in-span + amount within 25% but not 5% → Medium (amount_score>0)."""
    ev = _mk_ev(clearing_date="2024-01-15", net_amount=120.00)
    edf = [
        _mk_edf(
            invoice="T-003",
            period_from="01/01/2024",
            period_to="31/01/2024",
            amount=100.00,
        )
    ]
    matches = match_sap_events_to_edf([ev], edf)
    assert matches
    assert matches[0].confidence_band == "Medium", matches[0].confidence_band


def test_matcher_near_boundary_no_amount_caps_at_low() -> None:
    """Clearing within 3d of period end (25 pts) but no amount → Low not Medium."""
    ev = _mk_ev(clearing_date="2024-01-28", net_amount=50.00)
    edf = [
        _mk_edf(
            invoice="T-004",
            period_from="01/01/2024",
            period_to="31/01/2024",
            amount=10000.00,  # wildly off — no amount band hit
        )
    ]
    matches = match_sap_events_to_edf([ev], edf)
    # 28 vs 31 = 3d → date_score=25; amount_score=0 → total=25 → capped Low
    assert matches, "expected a match (Low)"
    assert matches[0].confidence_band == "Low", matches[0].confidence_band


def test_matcher_notes_string_says_coincidental_when_no_amount() -> None:
    """The in-span-no-amount notes string must say 'coincidental' so the
    surviving Low rows are self-explaining (spec §3.1)."""
    ev = _mk_ev(clearing_date="2024-01-15", net_amount=10.00)
    edf = [
        _mk_edf(
            invoice="T-005",
            period_from="01/01/2024",
            period_to="31/01/2024",
            amount=10000.00,
        )
    ]
    matches = match_sap_events_to_edf([ev], edf)
    assert matches, "expected a match (Low)"
    assert "coincidental" in matches[0].notes.lower(), matches[0].notes
