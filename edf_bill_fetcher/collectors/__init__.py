"""Collectors submodule — extraction orchestrators and source parsers.

Exposes the EvidenceEngine orchestrator plus per-source collector modules
for PDF, PST, and text inputs.  During the modularization refactor window
(Tasks 4-7), the parsing helpers called by these collectors continue to
live in ``edf_collector.py`` and are imported via the compat re-export
block.  Task 7 strips the re-exports.
"""

from edf_bill_fetcher.collectors.engine import EvidenceEngine

__all__ = ["EvidenceEngine"]
