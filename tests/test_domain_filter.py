from edf_bill_fetcher.helpers.domain_filter import matches_domain_filter


def test_domain_filter_matches_exact_address_and_subdomain() -> None:
    assert matches_domain_filter("Billing@sub.edf.com", "billing@edf.com, *.edf.com")


def test_domain_filter_rejects_unlisted_sender() -> None:
    assert not matches_domain_filter("billing@other.example", "edf.com")
