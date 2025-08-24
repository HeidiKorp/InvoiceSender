from pypdf import PdfReader, PdfWriter
from pathlib import Path
import re

class Invoice:
    def __init__(self, page, address, period, apartment):
        self.page = page
        self.address = address
        self.period = period
        self.apartment = apartment

    def __repr__(self):
        return f"Invoice(address={self.address}, period={self.period}, apartment={self.apartment})"

# Only splity the files here, extract information in another function
def separate_invoices(pdf_path):
    reader = PdfReader(pdf_path)
    invoices = []
    for _, page in enumerate(reader.pages):
        text = page.extract_text()
        extract_address_period_apartment(text)

        # invoice = Invoice(i, text)
        # invoices.append(invoice)
    return invoices

# test with faulty addresses
def extract_address_period_apartment(text):
    rows = text.splitlines()
    address_row = [row for row in rows if "aadress" in row.lower()][0]
    address_split = re.split(r'[:-]', address_row)
    address = address_split[1].strip().lower() if address_split else ""
    apartment = address_split[-1] if address_split else ""
    print(apartment)
    


# def save_each_invoice_as_file(invoices):

#     for invoice in invoices:
#         with open(f"invoice_{invoice.id}.txt", "w", encoding="utf-8") as f:
#             f.write(invoice.text)

# TODO: Save each file with the corresponding apartment number as the file name
# To find the apartment number, search for address + " " + maj_number + "-" + apartment