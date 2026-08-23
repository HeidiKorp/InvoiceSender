from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
import string

ESTONIAN_MONTHS = {
    1: "jaanuar",
    2: "veebruar",
    3: "märts",
    4: "aprill",
    5: "mai",
    6: "juuni",
    7: "juuli",
    8: "august",
    9: "september",
    10: "oktoober",
    11: "november",
    12: "detsember",
}
ESTONIAN_MONTH_NAMES = frozenset(ESTONIAN_MONTHS.values())
MIN_INVOICE_YEAR = 2001
MAX_INVOICE_YEAR = 2999


class ValidationError(ValueError):
    pass


def is_valid_period(period: str | None) -> bool:
    if not period:
        return False
    return period.strip().lower() in ESTONIAN_MONTH_NAMES


def is_valid_year(year: str | None) -> bool:
    if year is None:
        return False
    text = str(year).strip()
    if not text.isdigit():
        return False
    value = int(text)
    return MIN_INVOICE_YEAR <= value <= MAX_INVOICE_YEAR


def is_valid_address(address: str | None) -> bool:
    return bool(address and str(address).strip())


@dataclass
class Person:
    apartment: str
    address: str
    emails: list[str] = field(default_factory=list)

    def apartment_key(self) -> str:
        return str(self.apartment).strip()

    def invoice_pdf_path(self, invoices_dir: str | Path) -> str | None:
        path = Path(invoices_dir) / f"{self.apartment_key()}.pdf"
        return str(path) if path.exists() else None


@dataclass
class InvoiceItem:
    address: str
    period: str
    apartment: str
    year: str
    ky_name: Optional[str] = None
    pdf_page: Optional[object] = None
    excel_sheet_name: Optional[str] = None

    def apartment_key(self) -> str:
        return str(self.apartment).strip()

    def template_context(self) -> dict[str, str]:
        return {
            "address": self.address or "",
            "period": self.period or "",
            "apartment": self.apartment or "",
            "year": self.year or "",
            "ky_name": self.ky_name or "",
        }

    def format_template(self, template: str) -> str:
        context = self.template_context()
        placeholders = {
            field_name
            for _, field_name, _, _ in string.Formatter().parse(template)
            if field_name
        }
        safe_context = {key: context.get(key, "") for key in placeholders}
        safe_context.update(context)
        return template.format(**safe_context)

    def has_valid_period(self) -> bool:
        return is_valid_period(self.period)

    def has_valid_year(self) -> bool:
        return is_valid_year(self.year)

    def has_valid_address(self) -> bool:
        return is_valid_address(self.address)


@dataclass(frozen=True)
class InvoiceType:
    key: str
    label: str
    subject: str
    body: str


class Cancelled(Exception):
    pass
