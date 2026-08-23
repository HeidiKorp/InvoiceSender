from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
import threading

from src.data_classes import (
    Cancelled,
    InvoiceItem,
    InvoiceType,
    Person,
    ValidationError,
)
from src.clients_extractor import ClientsExtractor
from utils.file_utils import create_invoice_dir

INVOICE_FILE_EXTENSIONS = frozenset({".pdf"})


@dataclass
class InvoiceBatch:
    invoice_path: str
    clients_path: str
    invoice_type: InvoiceType
    cancel_event: threading.Event
    persons: list[Person] = field(default_factory=list)
    invoices: list[InvoiceItem] = field(default_factory=list)
    dest_dir: Path | None = None
    subject: str = ""
    body: str = ""

    @property
    def invoice_type_key(self) -> str:
        return self.invoice_type.key

    def load_clients(self) -> None:
        self.persons = ClientsExtractor().extract(self.clients_path)
        self._raise_if_cancelled()

    def load_invoices(self, on_progress=None) -> None:
        extractor = self._extractor()
        self.invoices = extractor.load(
            self.invoice_path,
            on_progress=on_progress,
            cancel_event=self.cancel_event,
        )
        self._raise_if_cancelled()
        if not self.invoices:
            raise ValidationError("Arveid ei leitud.")

    def apply_email_templates(self) -> None:
        example = self._representative_invoice()
        self.subject = example.format_template(self.invoice_type.subject)
        self.body = example.format_template(self.invoice_type.body)

    def prepare_destination(self) -> Path:
        parent = Path(self.invoice_path).resolve().parent
        dest = parent / "arved"
        try:
            dest.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            raise ValidationError(f"Kausta loomine ebaõnnestus:\n{dest}\n\n{e}") from e
        if not dest.exists() or not dest.is_dir():
            raise ValidationError(f"Kausta ei õnnestunud luua:\n{dest}")
        self.dest_dir = create_invoice_dir(dest, self._representative_invoice())
        return self.dest_dir

    def save(self, on_progress=None) -> Path:
        if self.dest_dir is None:
            self.prepare_destination()
        extractor = self._extractor()
        extractor.save(self, on_progress=on_progress)
        return self.dest_dir

    def match_apartments(self) -> list[str]:
        matched, problems = self._pair_persons(self._apartments_from_invoice_items())
        self.persons = matched
        return problems

    def match_against_saved_pdfs(
        self, invoices_dir=None
    ) -> tuple[list[Person], list[str]]:
        directory = invoices_dir or self.dest_dir
        return self._pair_persons(self._apartments_from_invoice_files(directory))

    def create_drafts(self) -> None:
        from src.email_sender import OutlookMailer

        OutlookMailer().save_drafts(self)

    def _representative_invoice(self) -> InvoiceItem:
        if not self.invoices:
            raise ValidationError("Arveid ei leitud.")
        address = next(
            (invoice.address for invoice in self.invoices if invoice.has_valid_address()),
            "",
        )
        period = next(
            (invoice.period for invoice in self.invoices if invoice.has_valid_period()),
            "",
        )
        year = next(
            (invoice.year for invoice in self.invoices if invoice.has_valid_year()),
            "",
        )
        ky_name = next(
            (invoice.ky_name for invoice in self.invoices if invoice.ky_name),
            None,
        )
        missing = []
        if not address:
            missing.append("aadress")
        if not period:
            missing.append("periood")
        if not year:
            missing.append("aasta")
        if missing:
            raise ValidationError(
                "Arvete andmetest ei õnnestunud leida kehtivat "
                + ", ".join(missing)
                + "."
            )
        first = self.invoices[0]
        return InvoiceItem(
            address=address,
            period=period,
            apartment=first.apartment,
            year=year,
            ky_name=ky_name,
        )

    def _raise_if_cancelled(self) -> None:
        if self.cancel_event.is_set():
            raise Cancelled()

    def _extractor(self):
        if self.invoice_type_key == "kommunaal":
            from src.pdf_extractor import PdfInvoiceExtractor

            return PdfInvoiceExtractor()
        if self.invoice_type_key == "kyte":
            from src.excel_invoice_extractor import ExcelInvoiceExtractor

            return ExcelInvoiceExtractor()
        raise ValidationError(f"Tundmatu arve tüüp: {self.invoice_type_key}")

    def _pair_persons(self, invoice_counts: Counter) -> tuple[list[Person], list[str]]:
        person_apts = self._apartments_from_persons()
        invoice_apts = set(invoice_counts.keys())

        missing = sorted(person_apts - invoice_apts, key=str)
        extra = sorted(invoice_apts - person_apts, key=str)
        duplicates = sorted(
            [apt for apt, count in invoice_counts.items() if count > 1], key=str
        )

        problems = self._build_validation_errors(missing, extra, duplicates)
        excluded_apts = set(missing) | set(duplicates)
        matched_persons = [
            person
            for person in self.persons
            if person.apartment_key() in invoice_apts
            and person.apartment_key() not in excluded_apts
        ]
        return matched_persons, problems

    def _apartments_from_persons(self) -> set[str]:
        return {person.apartment_key() for person in self.persons if person.apartment_key()}

    def _apartments_from_invoice_items(self) -> Counter:
        return Counter(
            invoice.apartment_key()
            for invoice in self.invoices
            if invoice.apartment_key()
        )

    def _apartments_from_invoice_files(self, invoices_dir) -> Counter:
        counts = Counter()
        for path in Path(invoices_dir).iterdir():
            if path.is_file() and path.suffix.lower() in INVOICE_FILE_EXTENSIONS:
                apt = path.stem.strip()
                if apt:
                    counts[apt] += 1
        return counts

    def _build_validation_errors(self, missing, extra, duplicates) -> list[str]:
        problems = []
        if missing:
            problems.append(f"Puuduvad arved korteritele: {', '.join(missing)}.\n")
        if extra:
            problems.append(f"Arved, millele ei leitud klienti: {', '.join(extra)}.\n")
        if duplicates:
            problems.append(
                f"Duplikaatsed arvefailid korteritele: {', '.join(duplicates)}.\n"
            )
        return problems
