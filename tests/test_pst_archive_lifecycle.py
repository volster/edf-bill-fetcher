"""The public PST/OST facade closes archives even when opening fails."""

from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from edf_bill_fetcher.collectors.engine import EvidenceEngine


@pytest.mark.parametrize("extension", ["pst", "ost"])
@pytest.mark.parametrize("failure", ["open", "root", "crawl"])
def test_archive_error_is_logged_and_resource_closed(
    monkeypatch: pytest.MonkeyPatch, extension: str, failure: str
) -> None:
    archive = MagicMock()
    constructor = MagicMock(return_value=archive)
    monkeypatch.setattr("edf_bill_fetcher.collectors.engine.HAS_PYPFF", True)
    monkeypatch.setattr(
        "edf_bill_fetcher.collectors.engine.pypff", SimpleNamespace(file=constructor)
    )
    engine = EvidenceEngine({}, lambda message: None)
    crawl = MagicMock()
    monkeypatch.setattr(engine, "crawl_pst", crawl)
    failing_call = {"open": archive.open, "root": archive.get_root_folder, "crawl": crawl}[failure]
    failing_call.side_effect = OSError("injected " + failure)
    getattr(engine, "process_" + extension + "_file")("synthetic." + extension)
    archive.close.assert_called_once()
    assert len(engine.error_log) == 1
    assert "injected " + failure in engine.error_log[0]
    assert "synthetic." + extension in engine.error_log[0]
    if failure == "open":
        archive.get_root_folder.assert_not_called()
        crawl.assert_not_called()
