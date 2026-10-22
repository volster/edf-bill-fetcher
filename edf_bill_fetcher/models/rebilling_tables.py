"""Shared presentation of accepted rebilling history and human dispute positions."""

from __future__ import annotations

from edf_bill_fetcher.helpers.formatting import fmt_money, report_date
from edf_bill_fetcher.models.output_tables import ledger_label
from edf_bill_fetcher.processors.preparation import identity


def reference(row: dict) -> str:
    """Retain original filename, source identity and the document's whole-PDF pages."""
    pages = "–".join(
        identity(row.get(key)) for key in ("Page From", "Page To") if identity(row.get(key))
    )
    return "\n".join(
        filter(
            None,
            (
                identity(row.get("Attachment Name")) or "Source unavailable",
                f"Pages: {pages or 'unknown'}",
                identity(row.get("Source ID")),
            ),
        )
    )


def document_label(doc: dict) -> str:
    """Identify a canonical document without merging equal invoice numbers."""
    row = doc["record"]
    return f"{identity(row.get('Invoice #')) or 'Invoice unavailable'} · {report_date(doc['invoice_date'])}"


def history_table(model: dict) -> list[list[str]]:
    """Show every document once, regardless of how many revision edges reference it."""
    indexed = {doc["row_id"]: doc for doc in model["documents"]}
    rows = [["Document / ledger", "Dates / readings", "Recorded amounts", "Relationships / source"]]
    for doc in model["documents"]:
        row = doc["record"]
        relations = []
        for direction, links, endpoint in (
            ("From", doc["incoming"], "older"),
            ("To", doc["outgoing"], "newer"),
        ):
            for link in links:
                other = indexed.get(link[endpoint])
                relations.append(
                    f"{direction} {document_label(other) if other else link[endpoint]}: {link['kind'].replace('_', ' ')} ({link['status']})"
                )
        rows.append(
            [
                document_label(doc)
                + "\n"
                + ledger_label(row)
                + "\n"
                + ("Included" if doc["included"] else "EXCLUDED")
                + " · "
                + identity(row.get("Entry Type")),
                f"Consumption: {report_date(doc['period_from']) or 'unknown'} – {report_date(doc['period_to']) or 'unknown'}\n"
                f"Reading: {report_date(doc['reading_date']) or 'unknown'} · {identity(row.get('Closing Reading Type')) or identity(row.get('Reading')) or 'type unknown'}\n"
                f"Register: {identity(row.get('Register Reading (kWh)')) or 'unknown'}; usage: {identity(row.get('Units (kWh)')) or 'unknown'} kWh",
                f"Period charge: {fmt_money(row.get('Period Charge (£)')) or 'unknown'}\n"
                f"Transaction: {fmt_money(row.get('Transaction Amount (£)')) or 'unknown'}\n"
                f"Gross electricity, printed: {fmt_money(row.get('Gross Electricity Charges (£)')) or 'unknown'}\n"
                f"Observed balance: {fmt_money(row.get('Balance (£)')) or 'unknown'}",
                "\n".join(
                    relations or ["No stored relationship; this document remains in the history."]
                )
                + "\n"
                + reference(row),
            ]
        )
    return rows


def comparison_table(model: dict, suggestion_limit: int | None = None) -> list[list[str]]:
    """Describe matched components and retain partial overlaps as unallocated."""
    indexed = {doc["row_id"]: doc for doc in model["documents"]}
    rows = [["Earlier → later", "Comparison scope", "Energy change, ex VAT", "Allocation / checks"]]
    links = model["comparisons"]
    if suggestion_limit is not None:
        links = [link for link in links if link["status"] != "suggested"] + [
            link for link in links if link["status"] == "suggested"
        ][:suggestion_limit]
    for link in links:
        a, b = indexed[link["older"]], indexed[link["newer"]]
        changes = []
        for part in link["components"]:
            period = f"{report_date(part['period_from'])} – {report_date(part['period_to'])}"
            changes.append(
                period
                + (
                    f": charge {fmt_money(part['charge_difference'])}; units effect {fmt_money(part['units_effect'])}; rate effect {fmt_money(part['rate_effect'])}; residual {fmt_money(part['residual'])}."
                    if part["status"] == "matched"
                    else ": unallocated; no unique matching component across both documents."
                )
            )
        for part in link["other_components"]:
            changes.append(
                f"Observed {part['kind']}: {fmt_money(part['earlier_amount']) or 'unknown'} → {fmt_money(part['later_amount']) or 'unknown'}; difference {fmt_money(part['difference']) or 'unallocated'} (only comparable for the same whole-document period)."
            )
        rows.append(
            [
                document_label(a) + " → " + document_label(b) + "\n" + link["status"].title(),
                "Same document period"
                if link["same_period"]
                else "Different / partial / unknown periods; whole-document amounts are not comparable allocations",
                "\n".join(changes)
                or "No comparable energy components extracted. Compare original pages and full history.",
                "Units effect uses the earlier rate. Rate effect uses the later units. Residual retains rounding or unexplained differences. Standing charges, VAT, discounts and cancellation components are separate in Rebilling components; no day-based allocation or inferred entitlement.",
            ]
        )
    return rows


def position_state(doc: dict) -> str:
    """Make exclusions and changed evidence impossible to mistake for an active claim."""
    position = doc["position"]
    status = position["status"].replace("_", " ").title() if position else "Unreviewed"
    return (
        ("EXCLUDED · " if not doc["included"] else "")
        + ("NEEDS REVIEW · " if doc["needs_review"] else "")
        + status
    )


def schedule_table(model: dict) -> list[list[str]]:
    """Include every saved position without adding overlapping challenges together."""
    rows = [
        [
            "Disputed document / status",
            "Human position (£)",
            "First demand / scope",
            "Basis / EDF evidence requested",
        ]
    ]
    for doc in model["documents"]:
        p = doc["position"]
        if p is None:
            continue
        rows.append(
            [
                document_label(doc)
                + "\n"
                + ledger_label(doc["record"])
                + "\n"
                + position_state(doc)
                + "\n"
                + reference(doc["record"]),
                f"Challenged: {fmt_money(p['challenged_amount']) or 'not specified'}\nProposed reduction: {fmt_money(p['proposed_reduction']) or 'not specified'}\nProposed only; not deducted from account figures.",
                f"Verified first demand: {report_date(doc['first_demand_date']) or 'not verified'}\n12-month reference: {report_date(doc['cutoff_date']) or 'unavailable'}\nInvoice and reading dates remain separate. Reference date alone does not establish protected charges or eligibility.",
                p["basis"] + "\nEDF must substantiate:\n" + (p["edf_request"] or "Not specified"),
            ]
        )
    return rows


def report_tables(model: dict | None) -> list[tuple[str, str, list[list[str]]]]:
    """Use the same accepted figures and wording in HTML, PDF and DOCX."""
    if model is None:
        return []
    note = model["notice"] + (
        " Automatic relationship suggestions were capped; all documents and explicit decisions remain visible."
        if model["suggestions_truncated"]
        else ""
    )
    return [
        ("Rebilling ledger", note, history_table(model)),
        (
            "Rebilling component comparisons",
            "Only uniquely matched consumption periods receive a units/rate comparison. Amounts are not reductions. All explicit comparisons and up to 30 automatic suggestions appear here; Rebilling comparisons in Excel contains every stored comparison. Scalar corrections do not rewrite original extracted components; verify any discrepancy against the source.",
            comparison_table(model, suggestion_limit=30),
        ),
        (
            "Dispute schedule",
            "Human proposals; do not sum overlapping positions. Excluded, withdrawn and stale positions remain labelled for traceability. Edit the schedule in the app or Excel, then regenerate the pack.",
            schedule_table(model),
        ),
    ]
