import sys
from collections import Counter
from pathlib import Path
from unittest.mock import MagicMock

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

sys.modules.setdefault("win32com", MagicMock())
sys.modules.setdefault("win32com.client", MagicMock())
sys.modules.setdefault("pywintypes", MagicMock())

from src.data_classes import InvoiceItem, Person
from src.email_sender import (
    _build_validation_errors,
    _check_duplicate_invoices,
    _check_extra_invoices,
    _check_missing_invoices,
    apartments_from_invoices,
    apartments_from_persons,
    get_person_invoice,
    persons_with_invoices,
    validate_persons_vs_invoice_items,
    validate_persons_vs_invoices,
)


def _person(apartment, email="a@example.com") -> Person:
    return Person(apartment=apartment, address="tn 1", emails=[email])


def _invoice(apartment: str) -> InvoiceItem:
    return InvoiceItem(
        address="tn 1", period="jaanuar", apartment=apartment, year="2026"
    )


def _touch_pdf(directory: Path, apartment: str) -> Path:
    path = directory / f"{apartment}.pdf"
    path.write_bytes(b"%PDF-1.4")
    return path


def test_apartments_from_persons_strips_and_skips_empty():
    persons = [_person(" 1 "), _person(""), _person("2")]
    assert apartments_from_persons(persons) == {"1", "2"}


def test_apartments_from_invoices_counts_pdf_stems(tmp_path):
    _touch_pdf(tmp_path, "1")
    _touch_pdf(tmp_path, "2")
    (tmp_path / "notes.txt").write_text("ignore")
    assert apartments_from_invoices(tmp_path) == Counter({"1": 1, "2": 1})


def test_build_validation_errors_includes_each_mismatch_type():
    problems = _build_validation_errors(["3"], ["9"], ["1"])
    assert problems == [
        "Puuduvad arved korteritele: 3.\n",
        "Arved, millele ei leitud klienti: 9.\n",
        "Duplikaatsed arvefailid korteritele: 1.\n",
    ]


def test_build_validation_errors_empty_when_all_match():
    assert _build_validation_errors([], [], []) == []


def test_check_helpers_sort_apartment_differences():
    person_apts = {"1", "3"}
    invoice_apts = {"1", "2"}
    assert _check_missing_invoices(person_apts, invoice_apts) == ["3"]
    assert _check_extra_invoices(person_apts, invoice_apts) == ["2"]
    assert _check_duplicate_invoices(Counter({"1": 1, "4": 2})) == ["4"]


def test_validate_returns_only_matched_pairs(tmp_path):
    _touch_pdf(tmp_path, "1")
    _touch_pdf(tmp_path, "2")
    persons = [_person("1"), _person("2")]

    matched, problems = validate_persons_vs_invoices(persons, tmp_path)

    assert problems == []
    assert matched == persons


def test_person_without_invoice_is_reported_and_excluded(tmp_path):
    _touch_pdf(tmp_path, "1")
    persons = [_person("1"), _person("3")]

    matched, problems = validate_persons_vs_invoices(persons, tmp_path)

    assert matched == [persons[0]]
    assert "Puuduvad arved korteritele: 3.\n" in problems


def test_invoice_without_person_is_reported_and_excluded(tmp_path):
    _touch_pdf(tmp_path, "1")
    _touch_pdf(tmp_path, "9")
    persons = [_person("1")]

    matched, problems = validate_persons_vs_invoices(persons, tmp_path)

    assert matched == persons
    assert "Arved, millele ei leitud klienti: 9.\n" in problems
    assert get_person_invoice("9", tmp_path) is not None
    assert all(_apartment_of(person) != "9" for person in matched)


def test_missing_person_and_extra_invoice_are_both_excluded(tmp_path):
    _touch_pdf(tmp_path, "1")
    _touch_pdf(tmp_path, "9")
    persons = [_person("1"), _person("3")]

    matched, problems = validate_persons_vs_invoices(persons, tmp_path)

    assert [person.apartment for person in matched] == ["1"]
    assert "Puuduvad arved korteritele: 3.\n" in problems
    assert "Arved, millele ei leitud klienti: 9.\n" in problems


def test_duplicate_invoice_apartment_is_reported_and_excluded(tmp_path, monkeypatch):
    _touch_pdf(tmp_path, "1")
    _touch_pdf(tmp_path, "2")
    persons = [_person("1"), _person("2")]
    monkeypatch.setattr(
        "src.email_sender.apartments_from_invoices",
        lambda invoices_dir, exts=None: Counter({"1": 2, "2": 1}),
    )

    matched, problems = validate_persons_vs_invoices(persons, tmp_path)

    assert matched == [persons[1]]
    assert "Duplikaatsed arvefailid korteritele: 1.\n" in problems


def test_persons_with_invoices_uses_the_same_filtered_result(tmp_path):
    _touch_pdf(tmp_path, "2")
    persons = [_person("1"), _person("2")]
    assert persons_with_invoices(persons, tmp_path) == [persons[1]]


def test_get_person_invoice_returns_path_only_when_file_exists(tmp_path):
    invoice = _touch_pdf(tmp_path, "4")
    assert get_person_invoice("4", tmp_path) == str(invoice)
    assert get_person_invoice("5", tmp_path) is None


def test_invoice_items_exclude_person_without_extracted_invoice():
    persons = [_person("1"), _person("3")]
    invoices = [_invoice("1")]

    matched, problems = validate_persons_vs_invoice_items(persons, invoices)

    assert matched == [persons[0]]
    assert "Puuduvad arved korteritele: 3.\n" in problems


def test_invoice_items_exclude_invoice_without_person():
    persons = [_person("1")]
    invoices = [_invoice("1"), _invoice("9")]

    matched, problems = validate_persons_vs_invoice_items(persons, invoices)

    assert matched == persons
    assert "Arved, millele ei leitud klienti: 9.\n" in problems


def test_pairing_is_by_apartment_not_email():
    shared = "korpheidi@gmail.com"
    persons = [
        Person(apartment="5", address="tn 1", emails=[shared, "other@example.com"]),
        Person(apartment="6", address="tn 1", emails=[shared]),
    ]

    only_five, problems = validate_persons_vs_invoice_items(persons, [_invoice("5")])
    assert [person.apartment for person in only_five] == ["5"]
    assert only_five[0].emails == [shared, "other@example.com"]
    assert "Puuduvad arved korteritele: 6.\n" in problems

    both, no_problems = validate_persons_vs_invoice_items(
        persons, [_invoice("5"), _invoice("6")]
    )
    assert [person.apartment for person in both] == ["5", "6"]
    assert no_problems == []
    assert both[0].emails == [shared, "other@example.com"]
    assert both[1].emails == [shared]


def test_leftover_pdf_does_not_add_person_without_extracted_invoice(tmp_path):
    _touch_pdf(tmp_path, "1")
    _touch_pdf(tmp_path, "3")
    persons = [_person("1"), _person("3")]
    invoices = [_invoice("1")]

    file_matched, _file_problems = validate_persons_vs_invoices(persons, tmp_path)
    assert [person.apartment for person in file_matched] == ["1", "3"]

    matched, problems = validate_persons_vs_invoice_items(persons, invoices)
    assert [person.apartment for person in matched] == ["1"]
    assert "Puuduvad arved korteritele: 3.\n" in problems


def _apartment_of(person: Person) -> str:
    return str(person.apartment).strip()
