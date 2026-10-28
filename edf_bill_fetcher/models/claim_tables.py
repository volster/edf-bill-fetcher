"""Shared human claim wording and tables for the workbook, reports and complaint pack."""

from __future__ import annotations

from edf_bill_fetcher.helpers.formatting import fmt_money, report_date
from edf_bill_fetcher.models.rebilling_tables import document_label, reference
from edf_bill_fetcher.processors.claim_reconciliation import has_active_claims


def scope_label(scope: dict) -> str:
    """Keep linked accounts and physical supply identities visible beside proposals."""
    values = [" / ".join(scope["accounts"]) or "Account unassigned"]
    values.extend(
        f"{field}: {', '.join(scope[field])}"
        for field in ("Fuel", "Meter", "Register")
        if scope[field]
    )
    return " · ".join(values)


def total_table(model: dict) -> list[list[str]]:
    """Distinguish a complete proposal from a reviewed subtotal and dated observations."""
    rows = [["Item", "Amount / date", "Meaning", "Scope"]]
    rows.append(
        [
            "Further reduction proposed",
            fmt_money(model["proposed_reduction"])
            if model["ready"]
            else "WITHHELD — unresolved claims"
            if has_active_claims(model)
            else "No active claims",
            "Human proposal after explicitly allocated relief already received; not an entitlement finding",
            "All reviewed claim groups"
            if model["ready"]
            else "Resolve the review issues before asserting one total",
        ]
    )
    rows.append(
        [
            "Reviewed subtotal",
            fmt_money(model["reviewed_subtotal"]),
            "Only current, reconciled groups; unresolved claims may change the requested outcome",
            f"{len(model['unresolved_positions'])} challenged documents still need reconciliation",
        ]
    )
    for ledger in model["ledgers"]:
        label = scope_label(ledger["scope"])
        rows.append(
            [
                "Further reduction in this ledger",
                fmt_money(ledger["reduction"]),
                "Groups: " + ", ".join(ledger["groups"]),
                label,
            ]
        )
        rows.append(
            [
                "Last usable observed balance",
                fmt_money(ledger["observed_balance"])
                if ledger["observed_balance"] is not None
                else "Unavailable",
                report_date(ledger["balance_date"]) or "No unambiguous dated observation",
                label,
            ]
        )
        rows.append(
            [
                "Illustrative balance after the reviewed subtotal",
                fmt_money(ledger["illustrative_balance"])
                if ledger["illustrative_balance"] is not None
                else "Unavailable",
                "Conditional arithmetic against that observation; remaining charges still require substantiation",
                label,
            ]
        )
    return rows


def group_table(model: dict, ledger: dict) -> list[list[str]]:
    """Present each consolidated claim once, with its component credit treatments."""
    docs = {doc["row_id"]: doc for doc in ledger["documents"]}
    rows = [
        ["Claim / status", "Consolidated proposal", "Credit accounting", "Scope / sources / review"]
    ]
    for group in model["groups"]:
        data = group["input"]
        status = (
            "Withdrawn"
            if data["status"] == "withdrawn"
            else "Reviewed for totaling"
            if group["ready"]
            else "NEEDS REVIEW"
            if group["needs_review"]
            else data["status"].title()
        )
        members = "\n".join(
            document_label(docs[key]) + "\n" + reference(docs[key]["record"])
            for key in data["members"]
        )
        credits = "\n".join(
            f"{docs[item['row_id']]['record'].get('Invoice #') or item['row_id']}: {fmt_money(item['amount'])} · {item['treatment'].replace('_', ' ')}. {item['reason']}"
            for item in data["credits"]
        )
        rows.append(
            [
                data["id"]
                + " · "
                + data["title"]
                + "\n"
                + status
                + "\n"
                + scope_label(group["scope"]),
                f"Challenged once: {fmt_money(data['challenged']) or 'Unknown'}\nReduction before received relief: {fmt_money(data['reduction']) or 'Unknown'}\nFurther reduction: {fmt_money(group['further_reduction']) or 'Unknown'}",
                f"Relief already received: {fmt_money(group['relief_received'])}\nCancellation offsets, not deducted again: {fmt_money(group['cancellation_offsets'])}\n"
                + (
                    credits
                    or "No credit allocations entered; review against the full earlier history."
                ),
                data["reason"]
                + (
                    "\nAccount link: " + data["account_link_reason"]
                    if data["account_link_reason"]
                    else ""
                )
                + "\n"
                + members
                + ("\nReview: " + " ".join(group["issues"]) if group["issues"] else ""),
            ]
        )
    return rows


def overlap_table(model: dict) -> list[list[str]]:
    """Disclose unresolved and explicitly separated claims without guessing percentages."""
    rows = [["Groups", "Potential overlap", "Review outcome", "Evidence / reason"]]
    for item in model["overlaps"]:
        rows.append(
            [
                " / ".join(item["groups"]),
                item["kind"].replace("_", " "),
                "Separate charges — human review"
                if item["resolved"]
                else "Needs review; total withheld",
                item["reason"]
                + (" Evidence changed; verify the decision again." if item["needs_review"] else ""),
            ]
        )
    return rows


def reading_table(ledger: dict) -> list[list[str]]:
    """Separate consumption, bill, reading and verified first-demand dates."""
    rows = [
        [
            "Disputed document / source",
            "Consumption and invoice dates",
            "Printed reading context",
            "Verified first demand / reference",
        ]
    ]
    for doc in ledger["documents"]:
        position, row = doc["position"], doc["record"]
        if not position or position["status"] != "challenged" or not doc["included"]:
            continue
        rows.append(
            [
                document_label(doc) + "\n" + reference(row),
                f"Consumption: {report_date(doc['period_from']) or 'Unknown'} – {report_date(doc['period_to']) or 'Unknown'}\nInvoice: {report_date(doc['invoice_date']) or 'Unknown'}",
                f"Reading date: {report_date(doc['reading_date']) or 'Unknown'}\nType: {row.get('Closing Reading Type') or row.get('Reading') or 'Unknown'}\nRegister: {row.get('Register Reading (kWh)') if row.get('Register Reading (kWh)') is not None else 'Unknown'}\n"
                + (
                    "OCR: verify against the original page."
                    if row.get("OCR Review Required")
                    else "A reading date does not establish when EDF received it."
                ),
                f"First demand: {report_date(doc['first_demand_date']) or 'Not verified'}\n12-month reference: {report_date(doc['cutoff_date']) or 'Unavailable'}\nReference only; no protected amount or day-based allocation is inferred.",
            ]
        )
    return rows


def report_tables(
    model: dict | None, ledger: dict | None
) -> list[tuple[str, str, list[list[str]]]]:
    """Render accepted results only; older snapshots do not gain recalculated claims."""
    if model is None or ledger is None or not (model["groups"] or model["unresolved_positions"]):
        return []
    return [
        ("Consolidated claim proposal", model["notice"], total_table(model)),
        (
            "Claim groups and credit allocations",
            "Amounts are entered for each economic challenge once. Earlier cancellation offsets and relief already received have different effects.",
            group_table(model, ledger),
        ),
        (
            "Claim overlap review",
            "A repeated document cannot be treated as an independent second claim. Other potential overlaps require a source-backed human decision.",
            overlap_table(model),
        ),
    ]


def complaint_paragraphs(model: dict) -> list[tuple[str, str]]:
    """Lead with the requested outcome without inventing facts or conceding unsupported shares."""
    context = model["input"]["context"]
    summary = (
        context["complaint_summary"]
        or "I dispute the historical additional and replacement charges identified in the attached schedule and request a complete billing reconciliation."
    )
    outcome = (
        context["requested_outcome"]
        or "Cancel protected historic additional charges, apply the adjustment requested in the schedule, and provide a fully substantiated revised account balance."
    )
    quantification = (
        f"The reconciled proposal is a further reduction of {fmt_money(model['proposed_reduction'])}, after the relief explicitly allocated as already received. The figures include applicable VAT."
        if model["ready"]
        else "Allocations remain unresolved, so a combined reduction is not yet asserted. The individual challenged amounts remain in the attached schedule; they must not be added together while overlap and credit use are unresolved."
    )
    return [
        ("Complaint summary", summary),
        ("Requested outcome", outcome),
        ("Proposed adjustment", quantification),
        (
            "Balance reconciliation",
            "All remaining charges remain subject to a complete reconciliation and supporting evidence. Dated observed balances and conditional arithmetic are provided for transparency; they do not establish current liability.",
        ),
        (
            "Evidence requested from EDF",
            "Please identify the earliest dated bill or statement demanding each additional amount; distinguish it from earlier estimated charges already demanded; and provide the complete cancellation, credit and replacement trail. Explain the meter readings, when they were obtained or received, and how additional units, rates, standing charges and VAT were allocated to the consumption periods. Substantiate any allocation EDF relies on when responding to the requested reduction.",
        ),
    ]
