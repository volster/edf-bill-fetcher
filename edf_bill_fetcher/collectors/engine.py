"""EvidenceEngine — orchestrator for PDF/PST/text extraction.

Placeholder for the modularization refactor (Task 4). The full
``EvidenceEngine`` class extraction was deferred because the class
references ~40 module-level names in ``edf_collector.py`` (parsing
helpers, regex patterns, format detectors) that would need to move
with it for standalone operation.

For now, this module re-exports ``EvidenceEngine`` from
``edf_collector`` so callers can use ``from edf_bill_fetcher.collectors
import EvidenceEngine`` without changing their code. The compat
shim is removed by Task 7.
"""

from edf_collector import EvidenceEngine

__all__ = ["EvidenceEngine"]
