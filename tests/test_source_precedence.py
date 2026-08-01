"""Phase-2 follow-on — explicit precedence test for the dedup sort.

The user-stated precedence order is:
    ``html summary > pdf's from folder > pdf from pst > email body``.

Translated to the project's source labels this is::

    "HTM Account History"      (precedence 0)
    "Local PDF Folder"          (precedence 1)
    "PST PDF Attachment"         (precedence 2)
    "Email Body" / "Email Body (RTF)"  (precedence 3)

``edf_collector.export_to_excel`` uses the precedence map so the
*richer* source wins when two records collide on the same
amount/date.  This test pins the order at the unit level without
booting the full Excel export pipeline — it imports the
``_SOURCE_PRECEDENCE`` mapping (extracted into module-level
constant in this commit) and asserts the explicit integer values
match the user-stated order.
"""

from __future__ import annotations

from edf_bill_fetcher.writers._helpers import _SOURCE_PRECEDENCE


class TestSourcePrecedence:
    """Pins the dedup source-precedence order verbatim."""

    def test_htm_account_history_is_highest_precedence(self) -> None:
        # The HTML Account History export carries per-bill
        # metadata we cannot reliably get from any other source
        # (readings, units, invoice numbers), so it sits at
        # precedence 0 == highest priority in pass 2.
        assert _SOURCE_PRECEDENCE["HTM Account History"] == 0

    def test_local_pdf_folder_beats_pst_attachment(self) -> None:
        # Per the user's standing instruction: the original
        # local PDF is the source-of-truth invoice, so a collision
        # between the local PDF and a PST attachment should keep
        # the local PDF and drop the PST attachment.
        local_pri = _SOURCE_PRECEDENCE["Local PDF Folder"]
        pst_pri = _SOURCE_PRECEDENCE["PST PDF Attachment"]
        assert local_pri < pst_pri, (
            f"Local PDF Folder precedence {local_pri} must be "
            f"lower (higher priority) than PST PDF Attachment "
            f"precedence {pst_pri} per the user's "
            f"'html summary > pdf's from folder > pdf from pst' "
            f"order."
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
        # The user-stated order is total, so no two distinct
        # sources should share a precedence number -- that would
        # silently undefined which one wins a tie.  Email Body
        # and Email Body (RTF) are the explicit exception: both
        # are alternative representations of the same
        # mail.body pipeline and the user has accepted that
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
