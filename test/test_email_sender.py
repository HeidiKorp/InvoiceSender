from pathlib import Path
import threading
from unittest.mock import MagicMock
import sys
import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

sys.modules.setdefault("win32com", MagicMock())
sys.modules.setdefault("win32com.client", MagicMock())
sys.modules.setdefault("pywintypes", MagicMock())

from src.data_classes import InvoiceItem, InvoiceType, Person, ValidationError
from src.invoice_batch import InvoiceBatch


def _person(apartment, email="a@example.com") -> Person:
    return Person(apartment=apartment, address="tn 1", emails=[email])


def _invoice(apartment: str) -> InvoiceItem:
    return InvoiceItem(
        address="tn 1", period="jaanuar", apartment=apartment, year="2026"
    )


def _type() -> InvoiceType:
    return InvoiceType(
        key="kommunaal",
        label="Kommunaalarved",
        subject="{address} arve {period} {year}",
        body="text",
    )


def _batch(persons, invoices) -> InvoiceBatch:
    return InvoiceBatch(
        invoice_path="invoices.pdf",
        clients_path="clients.xls",
        invoice_type=_type(),
        cancel_event=threading.Event(),
        persons=persons,
        invoices=invoices,
    )


def _touch_pdf(directory: Path, apartment: str) -> Path:
    path = directory / f"{apartment}.pdf"
    path.write_bytes(b"%PDF-1.4")
    return path


def test_apartments_from_persons_strips_and_skips_empty():
    persons = [_person(" 1 "), _person(""), _person("2")]
    batch = _batch(persons, [_invoice("1"), _invoice("2")])
    assert batch.match_apartments() == []
    assert [person.apartment_key() for person in batch.persons] == ["1", "2"]


def test_apartments_from_invoice_files_counts_pdf_stems(tmp_path):
    _touch_pdf(tmp_path, "1")
    _touch_pdf(tmp_path, "2")
    (tmp_path / "notes.txt").write_text("ignore")
    persons = [_person("1"), _person("2")]
    batch = _batch(persons, [_invoice("1"), _invoice("2")])
    matched, problems = batch.match_against_saved_pdfs(tmp_path)
    assert problems == []
    assert [person.apartment for person in matched] == ["1", "2"]


def test_build_validation_errors_includes_each_mismatch_type():
    persons = [_person("3")]
    batch = _batch(persons, [_invoice("9"), _invoice("1"), _invoice("1")])
    problems = batch.match_apartments()
    assert "Puuduvad arved korteritele: 3.\n" in problems
    assert "Arved, millele ei leitud klienti: 1, 9.\n" in problems
    assert "Duplikaatsed arvefailid korteritele: 1.\n" in problems


def test_build_validation_errors_empty_when_all_match():
    persons = [_person("1"), _person("2")]
    batch = _batch(persons, [_invoice("1"), _invoice("2")])
    assert batch.match_apartments() == []


def test_match_returns_only_matched_pairs():
    persons = [_person("1"), _person("2")]
    batch = _batch(persons, [_invoice("1"), _invoice("2")])
    assert batch.match_apartments() == []
    assert batch.persons == persons


def test_person_without_invoice_is_reported_and_excluded():
    persons = [_person("1"), _person("3")]
    batch = _batch(persons, [_invoice("1")])
    problems = batch.match_apartments()
    assert batch.persons == [persons[0]]
    assert "Puuduvad arved korteritele: 3.\n" in problems


def test_invoice_without_person_is_reported_and_excluded(tmp_path):
    persons = [_person("1")]
    batch = _batch(persons, [_invoice("1"), _invoice("9")])
    problems = batch.match_apartments()
    assert batch.persons == persons
    assert "Arved, millele ei leitud klienti: 9.\n" in problems
    leftover = _touch_pdf(tmp_path, "9")
    assert leftover.exists()
    assert all(person.apartment_key() != "9" for person in batch.persons)


def test_missing_person_and_extra_invoice_are_both_excluded():
    persons = [_person("1"), _person("3")]
    batch = _batch(persons, [_invoice("1"), _invoice("9")])
    problems = batch.match_apartments()
    assert [person.apartment for person in batch.persons] == ["1"]
    assert "Puuduvad arved korteritele: 3.\n" in problems
    assert "Arved, millele ei leitud klienti: 9.\n" in problems


def test_duplicate_invoice_apartment_is_reported_and_excluded():
    persons = [_person("1"), _person("2")]
    batch = _batch(persons, [_invoice("1"), _invoice("1"), _invoice("2")])
    problems = batch.match_apartments()
    assert batch.persons == [persons[1]]
    assert "Duplikaatsed arvefailid korteritele: 1.\n" in problems


def test_person_invoice_pdf_path_only_when_file_exists(tmp_path):
    invoice = _touch_pdf(tmp_path, "4")
    assert _person("4").invoice_pdf_path(tmp_path) == str(invoice)
    assert _person("5").invoice_pdf_path(tmp_path) is None


def test_pairing_is_by_apartment_not_email():
    shared = "korpheidi@gmail.com"
    persons = [
        Person(apartment="5", address="tn 1", emails=[shared, "other@example.com"]),
        Person(apartment="6", address="tn 1", emails=[shared]),
    ]
    batch = _batch(persons, [_invoice("5")])
    problems = batch.match_apartments()
    assert [person.apartment for person in batch.persons] == ["5"]
    assert batch.persons[0].emails == [shared, "other@example.com"]
    assert "Puuduvad arved korteritele: 6.\n" in problems

    batch_both = _batch(persons, [_invoice("5"), _invoice("6")])
    assert batch_both.match_apartments() == []
    assert [person.apartment for person in batch_both.persons] == ["5", "6"]


def test_leftover_pdf_does_not_add_person_without_extracted_invoice(tmp_path):
    _touch_pdf(tmp_path, "1")
    _touch_pdf(tmp_path, "3")
    persons = [_person("1"), _person("3")]
    batch_files = _batch(persons, [_invoice("1")])
    file_matched, _file_problems = batch_files.match_against_saved_pdfs(tmp_path)
    assert [person.apartment for person in file_matched] == ["1", "3"]

    batch = _batch(persons, [_invoice("1")])
    problems = batch.match_apartments()
    assert [person.apartment for person in batch.persons] == ["1"]
    assert "Puuduvad arved korteritele: 3.\n" in problems


def test_apply_email_templates_uses_address_period_year():
    batch = _batch([_person("1")], [_invoice("1")])
    batch.invoices[0].address = "Õismäe tee 48"
    batch.invoices[0].period = "august"
    batch.invoices[0].year = "2026"
    batch.apply_email_templates()
    assert batch.subject == "Õismäe tee 48 arve august 2026"


def test_apply_email_templates_uses_other_invoice_when_first_period_invalid():
    first = _invoice("1")
    first.address = "Õismäe tee 48"
    first.period = "periood"
    first.year = "2026"
    second = _invoice("2")
    second.address = ""
    second.period = "august"
    second.year = "99"
    batch = _batch([_person("1"), _person("2")], [first, second])
    batch.apply_email_templates()
    assert batch.subject == "Õismäe tee 48 arve august 2026"


def test_apply_email_templates_raises_when_no_valid_period():
    invoice = _invoice("1")
    invoice.period = "periood"
    invoice.year = "2026"
    invoice.address = "Õismäe tee 48"
    batch = _batch([_person("1")], [invoice])
    with pytest.raises(ValidationError, match="periood"):
        batch.apply_email_templates()
