from typing import Callable, List, Optional
from pathlib import Path
from dataclasses import dataclass
import threading


class ValidationError(ValueError):
    pass


class Person:
    def __init__(self, apartment, address, emails=None):
        self.apartment = apartment
        self.address = address
        self.emails = emails or []

    def __repr__(self):
        return f"Person(\nemails={self.emails}, \naddress={self.address}, \napartment={self.apartment}\n)"


@dataclass
class InvoiceItem:
    address: str
    period: str
    apartment: str
    year: str
    pdf_page: Optional[object] = None # Placeholder for PDF page object
    excel_sheet_name: Optional[str] = None # Placeholder for Excel sheet object

    def __repr__(self):
        return f"Invoice(address={self.address}, period={self.period}, apartment={self.apartment})"



@dataclass
class InvoiceBatch:
    parent: object
    persons: list[Person]
    invoices: list[InvoiceItem]
    invoice_path: str
    invoice_type_key: str
    dest_dir: Path
    subject: str
    body: str
    cancel_event: threading.Event


def create_invoice_batch(
    *,
    parent: object,
    persons: list[Person],
    invoices: list[InvoiceItem],
    invoice_path: str,
    invoice_type_key: str,
    dest_dir: Path,
    subject: str,
    body: str,
    cancel_event: threading.Event,
) -> InvoiceBatch:
    return InvoiceBatch(
        parent=parent,
        persons=persons,
        invoices=invoices,
        invoice_path=invoice_path,
        invoice_type_key=invoice_type_key,
        dest_dir=dest_dir,
        subject=subject,
        body=body,
        cancel_event=cancel_event,
    )


@dataclass(frozen=True)
class InvoiceType:
    key: str
    label: str
    subject: str
    body: str



class Cancelled(Exception):
    # "Operation cancelled by user."
    pass