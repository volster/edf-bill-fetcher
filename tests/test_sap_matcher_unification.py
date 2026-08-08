"""RED/GREEN: the SAP matcher duplicate in ``writers._helpers`` is removed.

The production helper module ``writers/_helpers.py`` must NOT contain a
local ``match_sap_events_to_edf`` definition — it must re-export the
canonical copy from ``processors/matching.py``.
"""

import edf_bill_fetcher.writers._helpers as helpers


def test_sap_matcher_duplicate_removed():
    """The production helper must re-export the canonical matcher.

    Structural delegation probe: a re-exported function keeps its
    defining module in ``__module__`` (and ``inspect.getsource`` reports
    the *defining* module's body, so line-count heuristics cannot
    distinguish a re-export from a local copy).

    RED state: the local duplicate's ``__module__`` is
    ``edf_bill_fetcher.writers._helpers``.
    GREEN state: after re-export it is ``edf_bill_fetcher.processors.matching``.
    """
    assert helpers.match_sap_events_to_edf.__module__ == "edf_bill_fetcher.processors.matching"


def test_sap_matcher_call_site_unchanged():
    """export.py's call site (writers/__init__.py import at :36, __all__
    entry at :65) must still resolve to match_sap_events_to_edf after
    the consolidation."""
    from edf_bill_fetcher.writers._helpers import match_sap_events_to_edf

    assert callable(match_sap_events_to_edf)
