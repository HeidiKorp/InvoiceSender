from pypdf import PdfReader, PdfWriter
import pandas as pd
from pathlib import Path
import re
from xls_extractor import ValidationError

class Invoice:
    def __init__(self, page, address, period, apartment, year):
        self.page = page
        self.address = address
        self.period = period
        self.apartment = apartment
        self.year = year

    def __repr__(self):
        return f"Invoice(address={self.address}, period={self.period}, apartment={self.apartment})"

# Only splity the files here, extract information in another function
def separate_invoices(pdf_path):
    reader = PdfReader(pdf_path)
    invoices = []
    for _, page in enumerate(reader.pages):
        text = page.extract_text()
        client_data = extract_address_period_apartment(text)
        invoice = Invoice(page, client_data["address"], client_data["period"], client_data["apartment"], client_data["year"])
        invoices.append(invoice)
    return invoices

# test with faulty addresses
def extract_address_period_apartment(text):
    rows = text.splitlines()
    address_parts = extract_parts(rows, "aadress", r'[:-]')
    address = address_parts[1] if len(address_parts) > 1 else ""
    apartment = address_parts[-1] if len(address_parts) > 2 else ""

    period_parts = extract_parts(rows, "periood")
    period = period_parts[1] if len(period_parts) > 1 else ""

    for i in rows:
        print(i)
    year = extract_parts(rows, "kuupäev", pattern=r'[:\-\. ]+')[-1]

    return {"address": address, "apartment": apartment, "period": period, "year": year}


# Find row keyword, split it, return list of stripped parts
def extract_parts(rows, keyword, pattern=r'[:\- ]+'):
    row = next((row for row in rows if keyword in row.lower()), None)
    if row is None:
        raise ValidationError(f"Keyword '{keyword}' not found in rows")
    return [part.strip().lower() for part in re.split(pattern, row) if part]


def save_each_invoice_as_file(invoices, dest):
    invoices_dir = dest / invoices[0].address.replace(" ", "_") / invoices[0].period
    invoices_dir.mkdir(parents=True, exist_ok=True)

    for invoice in invoices:
        writer = PdfWriter()
        writer.add_page(invoice.page)
        with open(invoices_dir / f'{invoice.apartment}.pdf', "wb") as f:
            writer.write(f)
    return invoices_dir
