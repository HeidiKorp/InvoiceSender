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
    batch = _batch(persons, [_invoice("9"), _invoice("1"), _invoice("1-kontrolli")])
    problems = batch.match_apartments()
    assert "Puuduvad arved korteritele: 3.\n" in problems
    assert "Arved, millele ei leitud klienti: 1, 9.\n" in problems
    assert (
        "Sama korteri number loeti mitu korda. "
        "Kontrolli faili nimi üle: 1-kontrolli.pdf.\n"
    ) in problems


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


def test_duplicate_invoice_apartment_keeps_first_and_lists_recheck_stem():
    persons = [_person("1"), _person("2")]
    batch = _batch(
        persons, [_invoice("1"), _invoice("1-kontrolli"), _invoice("2")]
    )
    problems = batch.match_apartments()
    assert batch.persons == persons
    assert "Duplikaatsed arvefailid korteritele: 1.\n" not in problems
    assert (
        "Sama korteri number loeti mitu korda. "
        "Kontrolli faili nimi üle: 1-kontrolli.pdf.\n"
    ) in problems


def test_duplicate_stems_still_warn_missing_apartment_in_sequence():
    persons = [_person("56"), _person("57"), _person("58")]
    batch = _batch(
        persons,
        [_invoice("56"), _invoice("56-kontrolli"), _invoice("58")],
    )
    problems = batch.match_apartments()
    assert [person.apartment for person in batch.persons] == ["56", "58"]
    assert "Arvete järjestuses puuduvad numbrid: 57.\n" in problems
    assert "Puuduvad arved korteritele: 57.\n" in problems
    assert (
        "Sama korteri number loeti mitu korda. "
        "Kontrolli faili nimi üle: 56-kontrolli.pdf.\n"
    ) in problems


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


def test_small_holes_are_warned_and_existing_invoices_still_match():
    persons = [_person("56"), _person("58"), _person("148"), _person("57")]
    invoices = [_invoice("56"), _invoice("58"), _invoice("148")]
    batch = _batch(persons, invoices)
    problems = batch.match_apartments()
    assert "Arvete järjestuses puuduvad numbrid: 57.\n" in problems
    assert "Puuduvad arved korteritele: 57.\n" in problems
    assert [person.apartment for person in batch.persons] == ["56", "58", "148"]


def test_lettered_apartment_keys_match():
    persons = [_person("44 a")]
    batch = _batch(persons, [_invoice("44A")])
    assert batch.match_apartments() == []
    assert batch.persons[0].apartment_key() == "44A"


def test_unreadable_invoice_file_asks_for_a_rename():
    unnamed = _invoice("nimetu-1")
    unnamed.needs_review = True
    reviewed = _invoice("4")
    reviewed.needs_review = True
    persons = [_person("56"), _person("57"), _person("58"), _person("4")]
    batch = _batch(persons, [_invoice("56"), _invoice("58"), unnamed, reviewed])
    problems = batch.match_apartments()
    assert "Puuduvad arved korteritele: 57.\n" in problems
    assert "Kontrolli üle arved: 4.\n" in problems
    assert (
        "Korteri numbrit ei õnnestunud lugeda. "
        "Nimeta fail ise ümber korteri numbriks: nimetu-1.pdf.\n"
    ) in problems
    assert not any("ei leitud klienti" in problem for problem in problems)
    assert [person.apartment for person in batch.persons] == ["56", "58", "4"]


def test_two_unreadable_invoices_are_listed_together():
    batch = _batch(
        [_person("1")],
        [_invoice("1"), _invoice("nimetu-1"), _invoice("nimetu-2")],
    )
    problems = batch.match_apartments()
    assert (
        "Korteri numbrit ei õnnestunud lugeda. "
        "Nimeta fail ise ümber korteri numbriks: nimetu-1.pdf, nimetu-2.pdf.\n"
    ) in problems
    assert not any("ei leitud klienti" in problem for problem in problems)


def test_review_apartments_are_reported_and_not_excluded():
    reviewed = _invoice("48")
    reviewed.needs_review = True
    persons = [_person("1"), _person("48")]
    batch = _batch(persons, [_invoice("1"), reviewed])
    problems = batch.match_apartments()
    assert "Kontrolli üle arved: 48.\n" in problems
    assert [person.apartment for person in batch.persons] == ["1", "48"]


def test_apply_email_templates_raises_when_no_valid_period():
    invoice = _invoice("1")
    invoice.period = "periood"
    invoice.year = "2026"
    invoice.address = "Õismäe tee 48"
    batch = _batch([_person("1")], [invoice])
    with pytest.raises(ValidationError, match="periood"):
        batch.apply_email_templates()
