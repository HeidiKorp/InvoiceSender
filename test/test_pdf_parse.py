import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.data_classes import is_valid_period, is_valid_year
from src.pdf_extractor import (
    _looks_like_date_pair,
    extract_address_period_apartment,
    extract_period,
    extract_year,
    has_reasonable_apartment_digits,
    pick_address_and_apt,
    strip_address_prefix,
)


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
        text, prev_apt=4, period="periood", year="99"
    )
    assert result["period"] == "august"
    assert result["year"] == "2026"


def test_pick_address_prefers_address_line_with_street_token():
    rows = [
        "Aadress Õismäe tee 48-5",
        "Viitenumber 12345",
        "Muu rida 12-99",
    ]
    address, apartment = pick_address_and_apt(rows, prev_apt=4)
    assert "48" in address
    assert str(apartment).startswith("5")
