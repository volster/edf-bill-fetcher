"""Tests for the canonical OFGEM caps loader (helpers/ofgem_caps.py)."""

import json

import pytest

from edf_bill_fetcher.helpers.ofgem_caps import load_ofgem_caps

FIXTURE = "tests/fixtures/ofgem_caps_minimal.json"


def test_loader_returns_expected_shape_from_fixture():
    caps, latest = load_ofgem_caps(path=FIXTURE)
    assert set(caps) == {"2024-Q1", "2024-Q2", "2024-Q3"}
    assert caps["2024-Q1"] == {"unit_rate": 28.62, "standing_charge": 53.35}
    assert caps["2024-Q2"] == {"unit_rate": 24.50, "standing_charge": 60.10}
    assert caps["2024-Q3"] == {"unit_rate": 24.50, "standing_charge": 60.10}
    assert latest == caps["2024-Q3"]


def test_auto_carry_false_returns_none_latest():
    caps, latest = load_ofgem_caps(auto_carry=False, path=FIXTURE)
    assert set(caps) == {"2024-Q1", "2024-Q2", "2024-Q3"}
    assert latest is None


def test_latest_known_is_max_key_not_hardcoded():
    caps, latest = load_ofgem_caps(path=FIXTURE)
    assert latest == caps[max(caps.keys())]
    assert list(caps)[-1] == "2024-Q3"


def test_caps_has_no_sentinel_key():
    caps, _ = load_ofgem_caps(path=FIXTURE)
    assert "_LATEST_KNOWN" not in caps
    assert all("_LATEST_KNOWN" not in v for v in caps.values())


def test_malformed_json_raises(tmp_path):
    bad = tmp_path / "bad.json"
    bad.write_text("{not json")
    with pytest.raises(json.JSONDecodeError):
        load_ofgem_caps(path=str(bad))


def test_missing_quarters_key_raises(tmp_path):
    bad = tmp_path / "noquarters.json"
    bad.write_text(json.dumps({"source": "x", "quarters": {}}))
    with pytest.raises(ValueError, match="quarters"):
        load_ofgem_caps(path=str(bad))


def test_quarter_missing_rates_raises(tmp_path):
    bad = tmp_path / "norates.json"
    bad.write_text(
        json.dumps(
            {
                "quarters": {
                    "2024-Q1": {"unit_rate": 28.62},
                }
            }
        )
    )
    with pytest.raises(ValueError, match="2024-Q1"):
        load_ofgem_caps(path=str(bad))


def test_carry_mismatch_warns(tmp_path):
    mismatch = tmp_path / "mismatch.json"
    mismatch.write_text(
        json.dumps(
            {
                "quarters": {
                    "2024-Q1": {"unit_rate": 28.62, "standing_charge": 53.35},
                    "2024-Q2": {"unit_rate": 99.99, "standing_charge": 60.10, "is_carry": True},
                }
            }
        )
    )
    with pytest.warns(UserWarning, match="carry"):
        load_ofgem_caps(path=str(mismatch))


def test_packaged_resource_default_loads_real_table():
    caps, latest = load_ofgem_caps()
    assert len(caps) >= 30
    assert latest == caps[max(caps.keys())]
    assert caps["2026-Q3"] == {"unit_rate": 26.11, "standing_charge": 57.19}
