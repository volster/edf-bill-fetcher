"""Tests for the ReportSection registry and RenderContext.

LAYMAN'S GUIDE
==============

Both the PDF and DOCX Ombudsman reports are built from a single shared
list called ``REPORT_SECTIONS`` that lives in ``edf_report.py``.
Each entry in that list is a ``ReportSectionMeta``:

    ("exec_summary", "Executive Summary", is_appendix=False)

The first item, the **key** (``"exec_summary"``), is how the rest of the
codebook refers to the section. The second item is the **title** printed
in the report and Table of Contents. The third item marks whether the
section is an Appendix (lettered "A", "B", "C") or a main report section
(numbered "1", "2", "3").

A ``RenderContext`` is built per-render. The user gets to toggle sections
on or off in the GUI dialog (``ReportOptionsDialog.SECTIONS``). The
context is asked: "for every section the user enabled, what number or
letter should it get, and what is its full heading line?". Both the TOC
builder and every body section builder consult the same context so
heading text and TOC numbering always match.

These tests pin down the rules. If the registry is changed in a way
that breaks the contract, the tests should fail loudly so the next
developer notices.
"""

import pytest

from edf_report import REPORT_SECTIONS, RenderContext


def _enabled_keys(*want: str) -> set[str]:
    """Smaller helper: build a set of enabled section keys from a few
    string names. Falls back to every section if no names are given
    (matches ``RenderContext()``'s default behaviour)."""
    if not want:
        return {s.key for s in REPORT_SECTIONS}
    return set(want)


class TestReportSectionsRegistry:
    """The registry is the single source of truth — both PDF and DOCX
    import it. Any change here ripples through both reports and the GUI
    options dialog, so we lock its shape."""

    def test_registry_has_expected_count(self):
        # We expect 14 sections: 11 main numeric + 3 lettered appendices.
        # The exact count matters less than "this is what downstream code
        # assumes" — change this assertion if you add a section.
        assert len(REPORT_SECTIONS) == 14

    def test_main_sections_have_is_appendix_false(self):
        # All numbered sections (1, 2, 3 ...) must have is_appendix=False.
        # Otherwise numbers and letters get tangled.
        for s in REPORT_SECTIONS:
            if s.key.startswith("appendix_"):
                assert s.is_appendix is True, f"{s.key} should be appendix=True"
            else:
                assert s.is_appendix is False, f"{s.key} should be appendix=False"

    def test_keys_are_unique(self):
        # Two sections sharing a key would make ``ctx.heading(key)``
        # ambiguous. The registry is a list, not a dict — uniqueness
        # is by convention.
        keys = [s.key for s in REPORT_SECTIONS]
        assert len(keys) == len(set(keys)), f"Duplicate registry keys: {keys}"

    def test_appendix_keys_come_after_main(self):
        # Appendices are at the tail of the list so the dispatcher can
        # safely use index slicing. Verify by walking the registry.
        seen_appendix = False
        for s in REPORT_SECTIONS:
            if s.is_appendix:
                seen_appendix = True
            else:
                assert not seen_appendix, "main sections must come before appendices"


class TestRenderContextNumbering:
    """RenderContext computes numeric (1, 2, 3...) and alphabetic (A, B, C...)
    labels for main and appendix sections respectively. The full heading
    string is built from ``"label title"``."""

    def test_full_selection_uses_legacy_layout(self):
        # With every section selected, main sections are 1..11 and
        # appendices are A..C. This is the legacy layout the rest of
        # the report pipeline was previously coded against.
        ctx = RenderContext()

        # Walk by registry order; convert to (label, title) pairs.
        labels_titles = [(spec.label, spec.section.title) for spec in ctx.sections_in_order]

        # First 11 entries are numeric (main). The keys match the
        # first 11 entries of REPORT_SECTIONS in registry order.
        for i, (label, _) in enumerate(labels_titles[:11], start=1):
            assert label == f"{i}.", f"slot {i} should be {i}."

        # Last 3 are alphabetic.
        assert labels_titles[11][0] == "A."
        assert labels_titles[12][0] == "B."
        assert labels_titles[13][0] == "C."

    def test_toggling_drop_does_not_break(self):
        """Disabling ``statistical`` should renumber everything after it.

        Layman: imagine the user is preparing a report where statistical
        analysis isn't meaningful (e.g. only 2 bills). They uncheck it.
        The TOC shouldn't still list "8. Statistical Analysis" —
        everything after it must shift up.

        Registry order matters: in ``REPORT_SECTIONS`` the layout is
        exec=1, key_findings=2, evidence_index=3, detailed_findings=4,
        timeline=5, ofgem=6, statistical=7, payment=8, forecast=9,
        data_quality=10, tariff=11, A. appendix_methodology, B.
        appendix_glossary, C. appendix_full_evidence.

        When ``statistical`` is removed, ``payment`` becomes the 7th
        selected main section. The appendices stay lettered A/B/C
        regardless.
        """
        ctx = RenderContext(
            selected={
                "exec_summary",
                "key_findings",
                "evidence_index",
                "detailed_findings",
                "timeline",
                "ofgem",
                "payment",
                "forecast",
                "data_quality",
                "tariff",
                "appendix_methodology",
                "appendix_glossary",
                "appendix_full_evidence",
            }
        )

        labels_by_key = {spec.section.key: spec.label for spec in ctx.sections_in_order}

        # payment moves to slot 7 after statistical is dropped.
        assert labels_by_key["payment"] == "7.", "payment should renumber to 7."
        # forecast moves to slot 8.
        assert labels_by_key["forecast"] == "8.", "forecast should renumber to 8."
        # OFGEM at slot 6 is unchanged.
        assert labels_by_key["ofgem"] == "6."
        # Appendices are still A, B, C in the same order.
        assert labels_by_key["appendix_methodology"] == "A."
        assert labels_by_key["appendix_glossary"] == "B."
        assert labels_by_key["appendix_full_evidence"] == "C."

    def test_only_appendices_selected(self):
        """If the user enables only appendices, the TOC contains only
        letter entries — no numeric ones."""
        ctx = RenderContext({"appendix_methodology", "appendix_glossary", "appendix_full_evidence"})

        labels = [s.label for s in ctx.sections_in_order]
        assert labels == ["A.", "B.", "C."]

    def test_empty_selection_renders_empty(self):
        """No enabled sections → no TOC rows.

        The actual builder converts this to a "<i>No sections
        selected.</i>" paragraph, but the context itself returns an
        empty iteration so the builder can do that."""
        ctx = RenderContext(set())
        assert ctx.sections_in_order == []

    def test_one_section_selected(self):
        """Single-section sanity check: the registry's only entry comes
        back out as label "1." (it's a main section)."""
        ctx = RenderContext({"exec_summary"})
        out = ctx.sections_in_order
        assert len(out) == 1
        assert out[0].label == "1."
        assert out[0].section.title == "Executive Summary"


class TestRenderContextHeading:
    """``ctx.heading(key)`` returns ``"<label> <title>"``. This is what the
    PDF/DOCX body uses for section titles."""

    def test_heading_full_selection_no_number(self):
        # When the section is selected, the heading includes its numeric
        # or alphabetic label followed by the title.
        ctx = RenderContext()
        h = ctx.heading("exec_summary")
        assert h == "1. Executive Summary"

    def test_heading_appendix_letter(self):
        ctx = RenderContext()
        h = ctx.heading("appendix_glossary")
        assert h == "B. Glossary"

    def test_heading_unselected_section_returns_title_only(self):
        # If the user disables a section but the body still asks for
        # the heading, we return the unadorned title — no "0. Title"
        # or stale label. This avoids printing a wrong number when the
        # body is rendered as a fallback.
        ctx = RenderContext({"exec_summary"})  # only one section enabled
        h = ctx.heading("forecast")  # forecast is NOT in enabled
        # The returned heading has no label prefix.
        assert h == "Forecast & Projection"

    def test_short_label_returns_only_number(self):
        # ``short_label`` is used by Detailed Findings to label its
        # 4.1, 4.2, 4.3, ... subsections. For section 4, it returns "4.".
        ctx = RenderContext()
        assert ctx.short_label("detailed_findings") == "4."

    def test_short_label_unselected_returns_empty(self):
        # An unselected section has no label. Subsections would then
        # print a bare number "1", "2", ... instead of "4.1", "4.2".
        ctx = RenderContext({"exec_summary"})
        assert ctx.short_label("detailed_findings") == ""

    def test_heading_unknown_key_raises(self):
        # Defensive: typos in keys must surface immediately so the
        # caller doesn't silently get an empty headline.
        ctx = RenderContext()
        with pytest.raises(KeyError):
            ctx.heading("totally_made_up_section")


class TestRenderContextDispatchGuard:
    """The dispatchers in ``generate_ombudsman_pdf`` and
    ``generate_ombudsman_docx`` raise ``RuntimeError`` when a registry
    entry has no builder wired into their ``section_builders`` dicts.
    This contract stops the registry from drifting away from the
    dispatch table silently.

    We exercise the guard by adding a fake section to ``REPORT_SECTIONS``
    and rendering a report that includes it. Without a matching dispatch
    entry, the dispatch loop must raise.
    """

    def _build_minimal_records(self):
        # A single minimal record is enough to drive the generator
        # up to the dispatch loop.
        return [
            {
                "Date": "01/01/2026",
                "Source": "HTM Account History",
                "Amount (£)": 100.0,
                "Period From": "N/A",
                "Period To": "N/A",
                "Invoice #": "N/A",
                "Period Charge (£)": "N/A",
                "Entry Type": "New Bill",
                "Reading": "Unknown",
                "Units (kWh)": "N/A",
                "Standing Chg (p/day)": "N/A",
                "Attachment Name": "N/A",
                "Details": "test",
                "Logic Used": "test",
            }
        ]

    def _make_minimal_engine(self):
        from types import SimpleNamespace

        return SimpleNamespace(pdf_count=0, email_count=0, records=[], filtered_records=[])

    def test_pdf_dispatcher_raises_when_registry_missing_builder(self, tmp_path, monkeypatch):
        """Adding a section to ``REPORT_SECTIONS`` without wiring it
        into the PDF dispatch table must blow up loud-and-fast.

        We monkeypatch the registry in-place by replacing the module
        attribute ``REPORT_SECTIONS`` with a new list that includes
        a fake orphan section. The dispatcher builds ``section_builders``
        off the live registry each call, so it sees the orphan and
        must raise because we never wired ``__test_orphan_section__``
        into the dispatch dict.
        """
        import edf_report

        original = edf_report.REPORT_SECTIONS
        try:
            monkeypatch.setattr(
                edf_report,
                "REPORT_SECTIONS",
                [
                    *original,
                    edf_report.ReportSectionMeta(
                        key="__test_orphan_section__",
                        title="Orphan Test Section",
                    ),
                ],
            )

            with pytest.raises(RuntimeError) as exc_info:
                edf_report.generate_ombudsman_pdf(
                    records=self._build_minimal_records(),
                    output_path=str(tmp_path / "report.pdf"),
                    config={
                        "report_sections": ["__test_orphan_section__"],
                    },
                    engine=self._make_minimal_engine(),
                    filtered=[],
                )

            # The error message should reference the missing key so the
            # developer immediately knows which builder they forgot to wire.
            assert "__test_orphan_section__" in str(exc_info.value), (
                "RuntimeError message should name the unregistered section key. "
                f"Got: {exc_info.value}"
            )
        finally:
            # ``monkeypatch.setattr`` rolls back automatically on teardown,
            # but in case the test failed mid-way we restore explicitly.
            edf_report.REPORT_SECTIONS = original


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
