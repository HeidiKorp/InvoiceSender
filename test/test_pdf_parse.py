import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.data_classes import is_valid_period, is_valid_year
from pypdf import PdfReader

from src.pdf_extractor import (
    _looks_like_date_pair,
    embedded_page_text_trustworthy,
    extract_address_period_apartment,
    extract_period,
    extract_year,
    has_reasonable_apartment_digits,
    pick_address_and_apt,
    review_apartment_sequence,
    saved_invoice_stems,
    small_apartment_gaps,
    strip_address_prefix,
)

S43_OKT_PDF = ROOT / "data" / "okt test data" / "S43 okt 26.pdf"


def test_strip_address_prefix_without_colon():
    assert strip_address_prefix("Aadress Õismäe tee 48-5") == "Õismäe tee 48-5"


def test_single_digit_apartment_is_accepted():
    assert has_reasonable_apartment_digits("5") is True
    assert has_reasonable_apartment_digits("") is False


def test_date_like_house_apartment_pairs_are_detected():
    assert _looks_like_date_pair("01", "02") is True
    assert _looks_like_date_pair("2026", "08") is True
    assert _looks_like_date_pair("48", "5") is False


def test_period_must_be_estonian_month():
    assert is_valid_period("august") is True
    assert is_valid_period("August") is True
    assert is_valid_period("periood") is False
    assert is_valid_period("") is False


def test_year_must_be_between_2001_and_2999():
    assert is_valid_year("2026") is True
    assert is_valid_year("2000") is False
    assert is_valid_year("3000") is False
    assert is_valid_year("abcd") is False


def test_extract_period_skips_the_word_periood():
    rows = [
        "Arve periood august 2026",
        "Kuupäev: 15.08.2026",
    ]
    assert extract_period(rows) == "august"
    assert extract_year(rows) == "2026"


def test_extract_period_reads_month_from_next_row():
    rows = [
        "Periood:",
        "august 2026",
    ]
    assert extract_period(rows) == "august"


def test_invalid_carried_period_is_replaced():
    text = "\n".join(
        [
            "Aadress Õismäe tee 48-5",
            "Arve periood august 2026",
            "Kuupäev: 15.08.2026",
        ]
    )
    result = extract_address_period_apartment(
        text, period="periood", year="99"
    )
    assert result["period"] == "august"
    assert result["year"] == "2026"


def test_pick_address_prefers_address_line_with_street_token():
    rows = [
        "Aadress Õismäe tee 48-5",
        "Viitenumber 12345",
        "Muu rida 12-99",
    ]
    address, apartment = pick_address_and_apt(rows)
    assert "48" in address
    assert str(apartment).startswith("5")


def test_lettered_apartment_stays_on_the_address_line():
    for line, expected in (
        ("Aadress: Rännaku pst 16-44 a", "44A"),
        ("Aadress: Rännaku pst 16-44a", "44A"),
        ("Aadress: Rännaku pst 16-44A", "44A"),
        ("Aadress: Rännaku pst 16-53 a", "53A"),
    ):
        address, apartment = pick_address_and_apt([line])
        assert "16" in address
        assert apartment == expected


def test_review_keeps_consecutive_apartments():
    assert review_apartment_sequence(["1", "2", "3"]) == [
        ("1", False),
        ("2", False),
        ("3", False),
    ]


def test_review_keeps_forward_gap_without_review():
    assert review_apartment_sequence(["1", "2", "4"]) == [
        ("1", False),
        ("2", False),
        ("4", False),
    ]


def test_review_keeps_gap_when_the_sequence_continues():
    assert review_apartment_sequence(["1", "2", "4", "5"]) == [
        ("1", False),
        ("2", False),
        ("4", False),
        ("5", False),
    ]


def test_review_keeps_same_number_with_a_letter():
    assert review_apartment_sequence(["44", "44 a", "45"]) == [
        ("44", False),
        ("44A", False),
        ("45", False),
    ]
    assert review_apartment_sequence(["51", "52", "52A", "53"]) == [
        ("51", False),
        ("52", False),
        ("52A", False),
        ("53", False),
    ]


def test_review_flags_only_a_mid_sequence_spike():
    assert review_apartment_sequence(["51", "52", "86", "54", "55"]) == [
        ("51", False),
        ("52", False),
        ("86", True),
        ("54", False),
        ("55", False),
    ]


def test_review_keeps_a_large_number_at_the_end():
    assert review_apartment_sequence(["56", "58", "148"]) == [
        ("56", False),
        ("58", False),
        ("148", False),
    ]
    assert small_apartment_gaps(["56", "58", "148"]) == ["57"]


def test_small_holes_list_one_or_two_missing_numbers():
    assert small_apartment_gaps(["53A", "55", "56", "58", "148"]) == ["54", "57"]


def test_small_or_backwards_reads_stay_as_read():
    assert review_apartment_sequence(["47", "4"]) == [("47", False), ("4", True)]
    assert review_apartment_sequence(["63", "6"]) == [("63", False), ("6", True)]
    assert review_apartment_sequence(["63", "64"]) == [("63", False), ("64", False)]
    assert review_apartment_sequence([""]) == [("", True)]


def test_unreadable_pages_are_named_nimetu_and_suspicious_numbers_keep_theirs():
    apartments = ["56", "", "??", "58", "47", "4"]
    assert review_apartment_sequence(apartments)[-1] == ("4", True)
    assert saved_invoice_stems(apartments) == [
        "56",
        "nimetu-1",
        "nimetu-2",
        "58",
        "47",
        "4",
    ]


def test_saved_invoice_stems_disambiguate_repeated_apartments():
    assert saved_invoice_stems(["56", "56", "58"]) == [
        "56",
        "56-kontrolli",
        "58",
    ]
    assert saved_invoice_stems(["56", "56", "56"]) == [
        "56",
        "56-kontrolli",
        "56-kontrolli-2",
    ]


def test_small_holes_ignore_kontrolli_suffix_on_duplicate_stems():
    assert small_apartment_gaps(["56", "56-kontrolli", "58"]) == ["57"]


def test_embedded_text_trustworthy_accepts_clean_address_line():
    text = "\n".join(
        [
            "Aadress: Sütiste tee 43-48",
            "Arve periood oktoober 2026",
        ]
    )
    assert embedded_page_text_trustworthy(text) is True


def test_embedded_text_trustworthy_rejects_empty_and_merged_periood_line():
    assert embedded_page_text_trustworthy("") is False
    assert embedded_page_text_trustworthy("   ") is False
    merged = "Aadress: Sütiste tee 43-13-18 Palur periood oktoober 2026"
    assert embedded_page_text_trustworthy(merged) is False


def test_embedded_text_trustworthy_rejects_misread_house_apartment():
    text = "Aadress: Õismäe tee 1413-1541"
    assert embedded_page_text_trustworthy(text) is False


def test_embedded_text_trustworthy_accepts_lettered_apartment():
    line = "Aadress: Rännaku pst 16-44 a"
    text = "\n".join([line, "Arve periood august 2026"])
    assert embedded_page_text_trustworthy(text) is True
    parsed = extract_address_period_apartment(text)
    assert parsed["apartment"] == "44A"


def test_s43_page_48_embedded_text_reads_apartment_48_not_48a():
    if not S43_OKT_PDF.is_file():
        return
    reader = PdfReader(S43_OKT_PDF)
    text = (reader.pages[47].extract_text() or "").strip()
    assert embedded_page_text_trustworthy(text) is True
    parsed = extract_address_period_apartment(text)
    assert parsed["apartment"] == "48"
