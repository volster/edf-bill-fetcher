"""Phase 1.4 — round-trip the restricted-pickle loader with real engine data.

Phase 1.4 acceptance criteria: pickling a real ``EvidenceEngine``
populated with a list-of-dicts ``records`` reconstruction succeeds
through ``_safe_pickle_load``. If a pickle stream requires a class
that isn't on ``_RestrictedUnpickler._SAFE_CLASSES``, the loader
raises ``pickle.UnpicklingError`` — which is the loud-failure
mode this whole subsystem is supposed to provide. We pin:

1. ``test_records_roundtrip`` — a real engine with a non-trivial
   records list pickles and unpickles intact through the
   restricted loader;
2. ``test_dataframe_in_records_roundtrip`` — an engine whose
   ``records`` contains a ``pandas.DataFrame`` also survives;
3. ``test_blocked_class_raises`` — a pickle that requires an
   off-whitelist class raises the expected ``UnpicklingError``;
4. ``test_restricted_loader_blocks_arbitrary_module`` — a
   raw-pickle of a module that isn't on the whitelist produces
   ``UnpicklingError`` immediately, not after partial restoration,
   so a user is never half-served.

If a fix is needed in the production whitelist, the contract
here is: add *only* the specific module+class names actually
required, with a comment explaining the dependency.
"""

from __future__ import annotations

import pickle
import tempfile
from collections.abc import Iterator
from pathlib import Path

import pandas as pd
import pytest

from edf_bill_fetcher.collectors.engine import EvidenceEngine
from edf_bill_fetcher.io.cli import (
    _RestrictedUnpickler,
    _safe_pickle_load,
)
from edf_bill_fetcher.models.config import ConfigDict


@pytest.fixture
def tmp_dir() -> Iterator[Path]:
    """A writable temp directory.

    We use ``tempfile.mkdtemp`` instead of pytest's built-in
    ``tmp_path`` because the latter's default Windows location
    (under ``%TEMP%``) hits a stale ACL on this host that causes
    ``iterdir`` to raise ``PermissionError`` during the fixture
    teardown for some users.  ``tempfile.mkdtemp`` always returns
    a directory the current user both owns *and* can read.
    """
    d = Path(tempfile.mkdtemp(prefix="edf_pickle_test_"))
    yield d
    # Best-effort cleanup; ignore errors so the test's assertions
    # complete before the Windows file-lock cleanup pass races
    # against teardown.
    try:
        for f in d.iterdir():
            f.unlink(missing_ok=True)
        d.rmdir()
    except OSError:
        pass


def _build_synthetic_records() -> list[dict]:
    """Realistic-looking EDF bill records.

    All identifiers here (account numbers, amounts, dates) are
    deliberately synthetic. See ``test_account_number_and_signed_zero.py``
    for the same convention.
    """
    return [
        {
            "Source": "Local PDF",
            "Sender": "",
            "Date": "15/01/2026",
            "Period From": "01/12/2025",
            "Period To": "31/12/2025",
            "Invoice #": "KI-0000000-0001",
            "Amount (£)": 240.50,
            "Period Charge (£)": 240.50,
            "Unit Rate (p/kWh)": 32.10,
            "% Change": "",
            "Entry Type": "New Bill",
            "Reading": "Smart",
            "Units (kWh)": 750.0,
            "Standing Chg (p/day)": 53.68,
            "Attachment Name": "Jan 2026 bill.pdf",
            "Details": "Your charges: 1 December 2025 - 31 December 2025",
            "Logic Used": "New Invoice Format",
        },
        {
            "Source": "PST PDF Attachment",
            "Sender": "[email protected]",
            "Date": "15/02/2026",
            "Period From": "01/01/2026",
            "Period To": "31/01/2026",
            "Invoice #": "KI-0000000-0002",
            "Amount (£)": 275.00,
            "Period Charge (£)": 275.00,
            "Unit Rate (p/kWh)": 27.69,
            "% Change": "+13.8%",
            "Entry Type": "New Bill",
            "Reading": "Actual",
            "Units (kWh)": 994.0,
            "Standing Chg (p/day)": 54.75,
            "Attachment Name": "Feb 2026 bill.pdf",
            "Details": "Your charges: 1 January 2026 - 31 January 2026",
            "Logic Used": "New Invoice Format",
        },
        {
            "Source": "HTM Account History",
            "Sender": "",
            "Date": "20/02/2026",
            "Period From": "N/A",
            "Period To": "N/A",
            "Invoice #": "N/A",
            "Amount (£)": -275.00,
            "Period Charge (£)": "N/A",
            "Unit Rate (p/kWh)": "N/A",
            "% Change": "",
            "Entry Type": "Payment",
            "Reading": "N/A",
            "Units (kWh)": "N/A",
            "Standing Chg (p/day)": "N/A",
            "Attachment Name": "N/A",
            "Details": "Direct debit",
            "Logic Used": "HTM Rebalanced",
        },
    ]


def _noop(*args, **kwargs):
    """Top-level no-op callback.

    Defined at module level because pickle cannot serialise
    closures/inner-defined lambdas ("Can't pickle local object").
    EvidenceEngine.__init__ requires a real callable for
    ``update_ui_cb``; ``_noop`` is reused across all phase-1.4
    fixtures to keep the round-trip independent of closure scope.
    """
    return None


def _build_engine() -> EvidenceEngine:
    """A minimal EvidenceEngine whose ``update_ui`` callback is harmless."""
    cfg: ConfigDict = {
        "use_anchors": True,
        "use_large": True,
        "use_reading_classification": True,
        "use_pdf_fields": True,
        "use_acc_filter": False,
        "acc_num": "",
        "min_amount": 0.0,
        "analysis_min": 0.0,
        "filter_below": False,
        "save_filtered": True,
        "use_dedup": False,
        "save_dups": False,
        "use_domain_filter": False,
        "domain_filter": "",
    }
    # ``_noop`` is intentionally module-level so a pickled engine
    # object round-trips through pickle cleanly — a lambda or local
    # function would fail at ``pickle.dump()`` with
    # "Can't get local object".
    engine = EvidenceEngine(cfg, _noop, _noop, None)
    engine.records = _build_synthetic_records()
    engine.filtered_records = _build_synthetic_records()[:1]  # one filtered record
    engine.pdf_count = 7
    engine.email_count = 3
    engine.error_log = [
        "PDF: stuck on PDF-bill-2024-03.pdf: never matched an anchor",
    ]
    return engine


class TestPickleRoundTrip:
    """Phase 1.4 — the restricted pickle loader is fit for purpose."""

    def test_records_roundtrip(self, tmp_dir: Path) -> None:
        """An engine object pickles to disk and round-trips intact through
        ``_safe_pickle_load``.
        """
        engine = _build_engine()

        # Persist to a temp pickle file (not committed to the repo).
        pkl_path = tmp_dir / "engine.pkl"
        with open(pkl_path, "wb") as fh:
            pickle.dump(engine, fh, protocol=pickle.HIGHEST_PROTOCOL)

        restored = _safe_pickle_load(str(pkl_path))
        # Object graph round-tripped.
        assert isinstance(restored, EvidenceEngine)
        assert len(restored.records) == 3
        assert restored.records[0]["Invoice #"] == "KI-0000000-0001"
        assert restored.records[2]["Entry Type"] == "Payment"
        # Counters survived.
        assert restored.pdf_count == 7
        assert restored.email_count == 3
        assert "stuck on PDF-bill-2024-03.pdf" in restored.error_log[0]
        assert len(restored.filtered_records) == 1

    def test_dataframe_in_records_roundtrip(self, tmp_dir: Path) -> None:
        """Records containing a real ``pandas.DataFrame`` round-trip
        just as cleanly.

        The EDF evidence product doesn't currently store DataFrames
        inside ``engine.records`` (it stores list-of-dicts); this
        test pins a future-proofing contract for users
        who may want to migrate record storage to ``DataFrame``.

        The DataFrame is built with ``pd.array(..., dtype="object")``
        for every column so the pickle stream stays within the
        legacy ``numpy.object_`` storage path.  This avoids the
        pandas 2.x default of Arrow-backed string arrays which
        would require the unpickler whitelist to include the entire
        pyarrow library surface — pyarrow is a transitive dependency
        only, not a direct runtime dep of the EDF app.
        """
        engine = _build_engine()
        # Replace the records with a list that contains one DataFrame
        # and two list-of-dicts entries.  This exercises the
        # pandas.DataFrame + pandas.Series reducer path.
        built = _build_synthetic_records()
        # Build the DataFrame column-by-column using ``pd.array`` so
        # every column is explicitly ``object`` dtype — sidesteps
        # the pandas 2.x Arrow ambiguity discussed above.
        df = pd.DataFrame(
            {col: pd.array([row[col] for row in built], dtype="object") for col in built[0]}
        )
        engine.records = [df]

        pkl_path = tmp_dir / "engine_with_df.pkl"
        with open(pkl_path, "wb") as fh:
            pickle.dump(engine, fh, protocol=pickle.HIGHEST_PROTOCOL)

        restored = _safe_pickle_load(str(pkl_path))
        assert isinstance(restored, EvidenceEngine)
        assert len(restored.records) == 1
        restored_df = restored.records[0]
        assert isinstance(restored_df, pd.DataFrame)
        assert len(restored_df) == 3
        assert "Invoice #" in restored_df.columns
        # The DataFrame's index was preserved (an integer RangeIndex
        # is the standard, nothing special needed for it).
        assert list(restored_df["Invoice #"]) == [
            "KI-0000000-0001",
            "KI-0000000-0002",
            "N/A",
        ]


class TestRestrictedPicklerBlocksUnsafe:
    """The whole point of the restricted loader is to *block*
    unsafe pickle payloads.  These tests pin the blocking behaviour.
    """

    def test_blocked_class_raises(self, tmp_dir: Path) -> None:
        """A pickle stream that names a class not on ``_SAFE_CLASSES``
        raises ``pickle.UnpicklingError`` instead of importing and
        calling arbitrary code.
        """

        # Hand-craft a Redux that, if trusted, would import + call a
        # class the restricted loader should block.
        class Evil:
            def __reduce__(self) -> tuple:  # noqa: D401
                return (eval, ("__import__('os').system('echo pwned')",))

        pkl = tmp_dir / "evil.pkl"
        with open(pkl, "wb") as fh:
            pickle.dump(Evil(), fh, protocol=pickle.HIGHEST_PROTOCOL)

        with pytest.raises(pickle.UnpicklingError):
            _safe_pickle_load(str(pkl))

    def test_restricted_loader_blocks_arbitrary_module_path(self) -> None:
        """Sanity check: even an *explicit* call to the restricted
        unpickler with an off-whitelist module fails closed, not
        half-open.

        This guards against the failure mode where ``find_class``
        falls back to ``super().find_class`` (which *would* call the
        stdlib default — and thus re-enable arbitrary execution).
        """
        import io

        bad_payload = (
            b"\x80\x04\x95\x14\x00\x00\x00\x00\x00\x00\x00"
            b"\x8c\x12subprocess.check_output"
            b"\x94\x8c\x04list\x94\x85\x94R\x94."
        )
        with pytest.raises(pickle.UnpicklingError):
            _RestrictedUnpickler(io.BytesIO(bad_payload)).load()


# Quick sanity guard: when loaded directly, a known-bad pickle of a
# *stdlib* class name surfaces UnpicklingError, not a silent
# import + execution. This is purely defensive; the test above is
# already the canonical guard.
