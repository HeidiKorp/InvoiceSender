from pypdf import PdfReader, PdfWriter
import pandas as pd
from pathlib import Path
import re
import io, gc, fitz, logging, sys, os
from PIL import Image, ImageOps, ImageFilter
import pytesseract
import traceback

from src.xls_extractor import ValidationError
from utils.logging_helper import log_line
from utils.ocr_helper import (
    check_tesseract_lang,
    render_page_to_image,
    preprocess_for_ocr,
    run_ocr_on_image,
)


logging.basicConfig(
    level=logging.INFO,  # 👈 enables INFO and above
    format="[%(levelname)s] %(message)s",
)


class Invoice:
    def __init__(self, page, address, period, apartment, year):
        self.page = page
        self.address = address
        self.period = period
        self.apartment = apartment
        self.year = year

    def __repr__(self):
        return f"Invoice(address={self.address}, period={self.period}, apartment={self.apartment})"


def ocr_pdf_all_pages(
    pdf_path: str,
    lang: str = "est",
    dpi: int = 300,
    psm: int = 6,
    oem: int = 1,
    timeout_sec: int = 120,
    on_progress=None,  # callback: on_progress(page_number: int, total_pages: int)
    cancel_flag=None,  # optional threading.Event to signal cancellation
) -> list[str]:
    """
    OCR all pages from a PDF file using PyMuPDF and Tesseract.
    Returns a list of extracted text strings, one per page (may be empty).
    """

    log_line(f"Using tesseract_cmd={pytesseract.pytesseract.tesseract_cmd}")

    texts: list[str] = []

    # Render scaling for given DPI (72 is default)
    scale = dpi / 72
    matrix = fitz.Matrix(scale, scale)

    check_tesseract_lang(lang)

    ocr_config = f"--oem {oem} --psm {psm}"

    with fitz.open(pdf_path) as doc:
        total_pages = doc.page_count

        for i, page in enumerate(doc, start=1):
            if cancel_flag and cancel_flag.is_set():
                logging.info("OCR process cancelled by user.")
                break

            if on_progress:
                try:
                    on_progress(i, doc.page_count)
                except Exception:
                    logging.debug(
                        "on_progress callback raised an exception:", exc_info=True
                    )

            logging.info(f"OCR processing page {i}/{doc.page_count} of '{pdf_path}'")
            img = None
            pix = None
            try:
                img = render_page_to_image(page, matrix)
                img = preprocess_for_ocr(img)
                text = run_ocr_on_image(img, lang, ocr_config, i, pdf_path, timeout_sec)
                texts.append(text)
            finally:
                try:
                    if img is not None:
                        img.close()
                except Exception:
                    pass
                
        gc.collect()
    return texts


# Only splity the files here, extract information in another function
def separate_invoices(pdf_path, on_progress=None, cancel_flag=None):
    """
    OCRib kogu PDF-i ja kasutab saadud teksti sinu olemasoleva parseriga.
    Säilitab sinu varasema 'Invoice(page, ...)' signatuuri, andes kaasa pypdf page-objekti.
    """
    print(on_progress)
    if on_progress:
        page_texts = ocr_pdf_all_pages(
            pdf_path, "est", dpi=300, on_progress=on_progress, cancel_flag=cancel_flag
        )
    else:
        page_texts = ocr_pdf_all_pages(
            pdf_path, "est", dpi=300, cancel_flag=cancel_flag
        )
    reader = PdfReader(pdf_path)

    if len(page_texts) != len(reader.pages) and not cancel_flag:
        raise ValidationError(
            f"PDF faili '{pdf_path}' OCR-tulemus on ebajärjekindel (lehtede arv ei klapi)."
        )

    invoices = []
    for idx, (page, text) in enumerate(zip(reader.pages, page_texts), start=1):
        rows = (text or "").splitlines()
        if len(text.strip()) == 0:
            logging.error(
                f"SEPARATE_INVOICES: EMPTY text for page {idx} of '{pdf_path}'"
            )
            raise ValidationError(
                f"PDF faili '{pdf_path}' leheküljelt {idx} ei õnnestunud teksti lugeda ka pärast OCR-i. "
                "PDF võib olla vigane."
            )
        client_data = extract_address_period_apartment(text)
        invoice = Invoice(
            page,
            client_data["address"],
            client_data["period"],
            client_data["apartment"],
            client_data["year"],
        )
        invoices.append(invoice)
    return invoices


# test with faulty addresses
def extract_address_period_apartment(text):
    rows = text.splitlines()
    address_parts = extract_parts(rows, "aadress", r"[:-]")
    address = address_parts[1] if len(address_parts) > 1 else ""
    apartment = address_parts[-1] if len(address_parts) > 2 else ""

    period_parts = extract_parts(rows, "periood")
    period = period_parts[1] if len(period_parts) > 1 else ""

    year = extract_parts(rows, "kuupäev", pattern=r"[:\-\. ]+")[-1]

    return {"address": address, "apartment": apartment, "period": period, "year": year}


# Find row keyword, split it, return list of stripped parts
def extract_parts(rows, keyword, pattern=r"[:\- ]+"):
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
        with open(invoices_dir / f"{invoice.apartment}.pdf", "wb") as f:
            writer.write(f)
    return invoices_dir
