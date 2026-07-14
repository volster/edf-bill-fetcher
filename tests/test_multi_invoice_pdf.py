from __future__ import annotations

from edf_collector import slice_pdf_pages


def test_empty_input_returns_single_empty_chunk():
    assert slice_pdf_pages([]) == [[]]


def test_no_markers_returns_single_chunk():
    pages = ["random text", "more text", "yet more"]
    assert slice_pdf_pages(pages) == [pages]


def test_single_invoice_number_on_page1_keeps_all_pages():
    pages = [
        "Invoice number: KI-31105244-0004\nDate issued: 07 October 2024",
        "Your charges in detail",
        "Make a complaint",
        "Get in touch",
    ]
    assert slice_pdf_pages(pages) == [pages]


def test_two_invoices_split_at_second_invoice_number():
    pages = [
        "Invoice number: T78701920069\nBill date: 16 Sep 2023\nPage 1 of 4",
        "Page 2 of 4",
        "Page 3 of 4",
        "Page 4 of 4",
        "Invoice number: T78701920070\nBill date: 17 Oct 2023\nPage 1 of 4",
        "Page 2 of 4",
        "Page 3 of 4",
        "Page 4 of 4",
    ]
    result = slice_pdf_pages(pages)
    assert len(result) == 2
    assert result[0] == pages[0:4]
    assert result[1] == pages[4:8]


def test_page1_of_n_marker_alone_splits_even_without_invoice_number():
    pages = [
        "Some cover letter\nPage 1 of 4",
        "Page 2 of 4",
        "Page 3 of 4",
        "Page 4 of 4",
        "Different invoice header\nPage 1 of 3",
        "Page 2 of 3",
        "Page 3 of 3",
    ]
    result = slice_pdf_pages(pages)
    assert len(result) == 2
    assert result[0] == pages[0:4]
    assert result[1] == pages[4:7]


def test_page_n_of_n_is_inclusive_stays_with_current_slice():
    pages = [
        "Invoice number: X\nPage 1 of 4",
        "Page 2 of 4",
        "Page 3 of 4",
        "Page 4 of 4",
        "Invoice number: Y\nPage 1 of 4",
        "Page 2 of 4",
        "Page 3 of 4",
        "Page 4 of 4",
    ]
    result = slice_pdf_pages(pages)
    assert len(result) == 2
    assert result[0] == pages[0:4]
    assert result[1] == pages[4:8]


def test_textual_page_marker_one_of_four_matches():
    pages = [
        "Invoice number: A\none of four",
        "two of four",
        "three of four",
        "four of four",
        "Invoice number: B\none of four",
        "two of four",
    ]
    result = slice_pdf_pages(pages)
    assert len(result) == 2
    assert result[0] == pages[0:4]
    assert result[1] == pages[4:6]


def test_slash_page_marker_1_4_matches():
    pages = [
        "Invoice number: A\n1/4",
        "2/4",
        "3/4",
        "4/4",
        "Invoice number: B\n1/3",
        "2/3",
        "3/3",
    ]
    result = slice_pdf_pages(pages)
    assert len(result) == 2
    assert result[0] == pages[0:4]
    assert result[1] == pages[4:7]


def test_blank_pages_are_ignored_as_boundaries_but_kept_in_chunks():
    pages = [
        "Invoice number: A\nPage 1 of 4",
        "",
        "Page 3 of 4",
        "Page 4 of 4",
        "Invoice number: B\nPage 1 of 4",
        "Page 2 of 4",
        "Page 3 of 4",
        "Page 4 of 4",
    ]
    result = slice_pdf_pages(pages)
    assert len(result) == 2
    assert result[0] == pages[0:4]
    assert result[1] == pages[4:8]


def test_invoice_number_takes_precedence_when_both_markers_on_same_page():
    pages = [
        "Invoice number: T78701920069\nBill date: 16 Sep 2023\nPage 1 of 4",
        "Page 2 of 4",
        "Page 3 of 4",
        "Page 4 of 4",
        "Invoice number: T78701920070\nPage 1 of 4",
        "Page 2 of 4",
    ]
    result = slice_pdf_pages(pages)
    assert len(result) == 2
    assert result[0] == pages[0:4]
    assert result[1] == pages[4:6]


def test_realistic_d2_merged_pdf_yields_8_invoices():
    pages = []
    for i in range(8):
        pages.append(f"Invoice number: T787019200{69 + i}\nPage 1 of 4")
        pages.append("Page 2 of 4")
        pages.append("Page 3 of 4")
        pages.append("Page 4 of 4")
    assert len(pages) == 32
    result = slice_pdf_pages(pages)
    assert len(result) == 8
    for chunk in result:
        assert len(chunk) == 4
