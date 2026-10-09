"""Format-independent report section registry and numbering."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ReportSectionMeta:
    """Manifest entry for one report section."""

    key: str
    title: str
    is_appendix: bool = False


REPORT_SECTIONS: list[ReportSectionMeta] = [
    ReportSectionMeta("exec_summary", "Executive Summary"),
    ReportSectionMeta("key_findings", "Key Findings Summary"),
    ReportSectionMeta("evidence_index", "Evidence Index & Source Cross-Reference"),
    ReportSectionMeta("detailed_findings", "Detailed Findings"),
    ReportSectionMeta("timeline", "Timeline of Events"),
    ReportSectionMeta("ofgem", "OFGEM Price Cap Comparison"),
    ReportSectionMeta("statistical", "Statistical Analysis"),
    ReportSectionMeta("payment", "Payment & Credit Analysis"),
    ReportSectionMeta("forecast", "Forecast & Projection"),
    ReportSectionMeta("data_quality", "Data Quality Assessment"),
    ReportSectionMeta("tariff", "Tariff Impact Analysis"),
    ReportSectionMeta("compensation", "Compensation Analysis"),
    ReportSectionMeta("appendix_methodology", "Methodology & Data Sources", is_appendix=True),
    ReportSectionMeta("appendix_glossary", "Glossary", is_appendix=True),
    ReportSectionMeta("appendix_full_evidence", "Full Evidence Table", is_appendix=True),
]


@dataclass(frozen=True)
class _LabelledSection:
    """A section after numbering has been resolved for the selected report."""

    section: ReportSectionMeta
    label: str  # e.g. "1." or "A."
    index: int  # within main (or appendix) list


class RenderContext:
    """Per-render state used to derive headings consistently.

    The body builders call ``ctx.heading(key)`` to get the heading string for
    the section they own. The TOC builder iterates ``ctx.sections_in_order``
    to produce the matching TOC.
    """

    def __init__(self, selected: set[str] | list[str] | None = None) -> None:
        """Build a render context for the selected report sections.

        Args:
            selected: Section keys to include; ``None`` defaults to every
                registered section so legacy context-free callers see the
                full 1..11 + A..C layout.

        """
        # Distinguish three cases:
        #   * ``selected is None`` — no explicit choice; default to every
        #     section so legacy context-free tests/CI still see the
        #     full registry with the 1..11 + A..C layout.
        #   * ``selected`` is any other iterable (including an empty set) —
        #     use it verbatim. An empty set produces an empty context.
        # We previously conflated "no argument" with "empty set",
        # which silently swallowed intent and made it impossible to
        # render an empty report from the GUI dialog. Now empty set
        # is honoured.
        if selected is None:
            self.selected = {s.key for s in REPORT_SECTIONS}
        else:
            self.selected = set(selected)

        # Compute numeric/alphabetic labels only for selected & visible sections,
        # skipping framing sections (cover, toc) which are not listed in
        # REPORT_SECTIONS — they never appear in the body nor the TOC.
        visible = [s for s in REPORT_SECTIONS if s.key in self.selected]
        main = [s for s in visible if not s.is_appendix]
        appendix = [s for s in visible if s.is_appendix]
        self._labelled: dict[str, _LabelledSection] = {}
        for i, section in enumerate(main, start=1):
            self._labelled[section.key] = _LabelledSection(section=section, label=f"{i}.", index=i)
        for i, section in enumerate(appendix):
            label = chr(ord("A") + i) + "."
            self._labelled[section.key] = _LabelledSection(
                section=section, label=label, index=i + 1
            )

    @property
    def sections_in_order(self) -> list[_LabelledSection]:
        """All sections that should appear in the TOC, in order."""
        main = [v for v in self._labelled.values() if not v.section.is_appendix]
        appendix = [v for v in self._labelled.values() if v.section.is_appendix]
        return main + appendix

    def heading(self, key: str) -> str:
        """Return the full heading line for the given section key.

        e.g. ``ctx.heading('timeline')`` returns ``"5. Timeline of Events"`` if
        timeline is the 5th selected main section. Raises KeyError if the key
        is not in REPORT_SECTIONS.
        """
        section = next((s for s in REPORT_SECTIONS if s.key == key), None)
        if section is None:
            raise KeyError(f"Unknown section key: {key!r}")
        labelled = self._labelled.get(key)
        if labelled is None:
            # Section was present but not selected — use its natural title
            # (no number) so we never produce "0. Title" or similar.
            return section.title
        return f"{labelled.label} {section.title}"

    def short_label(self, key: str) -> str:
        """Just the numeric / alphabetic marker, e.g. ``"3."`` or ``"A."``.

        Returns ``""`` if the section is not selected or unrecognised.
        """
        labelled = self._labelled.get(key)
        return labelled.label if labelled else ""
