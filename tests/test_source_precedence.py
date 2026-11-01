"""Tests for the configured source-precedence order used in duplicate selection.

HTM account history has priority 0, local PDFs priority 1, email PDF attachments
priority 2, and email bodies priority 3. Explicit tie buckets cover equivalent
source types. These tests check the mapping independently of workbook export.
"""

from __future__ import annotations

from edf_collector import _SOURCE_PRECEDENCE


class TestSourcePrecedence:
    """Pins the dedup source-precedence order verbatim."""

    def test_htm_account_history_is_highest_precedence(self) -> None:
        # The HTML Account History export carries per-bill
        # metadata we cannot reliably get from any other source
        # (readings, units, invoice numbers), so it sits at
        # precedence 0 == highest priority in pass 2.
        assert _SOURCE_PRECEDENCE["HTM Account History"] == 0

    def test_local_pdf_folder_beats_pst_attachment(self) -> None:
        # Local PDFs have higher priority than email attachments when
        # selecting between otherwise equivalent source observations.
        local_pri = _SOURCE_PRECEDENCE["Local PDF Folder"]
        pst_pri = _SOURCE_PRECEDENCE["PST PDF Attachment"]
        assert local_pri < pst_pri, (
            f"Local PDF Folder precedence {local_pri} must be "
            f"lower (higher priority) than PST PDF Attachment "
            f"precedence {pst_pri} under the configured source order."
        )
        # And the specific values are stable at 1 and 2.
        assert local_pri == 1
        assert pst_pri == 2

    def test_email_body_is_lowest_precedence(self) -> None:
        # Email body extractions are the lowest-quality source;
        # they lose every tie against HTM / Local PDF / PST.
        # Plain and RTF share the same precedence bucket (3)
        # because they are alternative renderings of the same
        # underlying mail.body pipeline.
        for label in ("Email Body", "Email Body (RTF)"):
            assert _SOURCE_PRECEDENCE[label] == 3, (
                f"{label!r} should map to the lowest precedence "
                f"(3); got {_SOURCE_PRECEDENCE[label]}"
            )

    def test_precedence_is_strictly_descending(self) -> None:
        # The configured order is total, so no two distinct
        # sources should share a precedence number -- that would
        # silently undefined which one wins a tie.  Email Body
        # and Email Body (RTF) are the explicit exception: both
        # are alternative representations of the same
        # mail.body pipeline and
        # they'll be tie-broken by the secondary sort key
        # (date, then invoice number).  Local PDF Folder and
        # Statement Reconciliation also share precedence 1:
        # both are direct paper/PDF artifacts produced by EDF
        # for the same account, so when a standalone credit
        # note collides with a "reversal" line on the
        # consolidated statement, the dedup logic collapses
        # them by Amount (£) regardless of which one wins.
        tie_buckets = [
            ("Email Body", "Email Body (RTF)"),
            ("Local PDF Folder", "Statement Reconciliation"),
        ]
        seen: dict[int, list[str]] = {}
        for label, pri in _SOURCE_PRECEDENCE.items():
            this_pair = {label, seen[pri][-1]} if pri in seen else set()
            if pri in seen and not any(this_pair <= set(bucket) for bucket in tie_buckets):
                raise AssertionError(
                    f"Source {label!r} duplicates precedence {pri}; "
                    f"first user: {seen[pri]!r}.  The allowed tie "
                    f"buckets are {tie_buckets!r}; any other shared "
                    f"precedence number is a configuration error."
                )
            seen.setdefault(pri, []).append(label)
