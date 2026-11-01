"""Source collection and extraction orchestration.

Exports ``EvidenceEngine``, which coordinates the source adapters and parsers.
"""

from edf_bill_fetcher.collectors.engine import EvidenceEngine

__all__ = ["EvidenceEngine"]
