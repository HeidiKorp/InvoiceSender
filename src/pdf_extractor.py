from pypdf import PdfReader, PdfWriter
import pandas as pd
from pathlib import Path
import re
import io, gc, fitz, logging, sys, os
from PIL import Image, ImageOps, ImageFilter
import pytesseract
import traceback

from src.data_classes import ValidationError
from utils.logging_helper import log_line
from utils.ocr_helper import (
    check_tesseract_lang,
    render_page_to_image,
    preprocess_for_ocr,
    run_ocr_on_image,
)
from src.data_classes import InvoiceItem
from utils.file_utils import create_invoice_dir


logging.basicConfig(
    level=logging.INFO,  # 👈 enables INFO and above
    format="[%(levelname)s] %(message)s",
)


def _validate_page_text(text: str, page_number: int, pdf_path: str):
    if not text or not text.strip():
        logging.error(
            f"SEPARATE_INVOICES: EMPTY text for page {page_number} of '{pdf_path}'"
        )
        raise ValidationError(
            f"PDF faili '{pdf_path}' leheküljelt {page_number} ei õnnestunud teksti lugeda ka pärast OCR-i. "
            "PDF võib olla vigane."
        )


def _parse_invoice_page(page, text: str, page_number: int, pdf_path: str) -> dict:
    _validate_page_text(text, page_number, pdf_path)
    client_data = extract_address_period_apartment(text)
    return InvoiceItem(
        pdf_page=page,
        address=client_data["address"],
        period=client_data["period"],
        apartment=client_data["apartment"],
        year=client_data["year"],
    )


def _ocr_single_page(page, page_idx, doc, pdf_path, lang, ocr_config, matrix, timeout_sec, on_progress, cancel_flag):
    """OCR a single page and return the extracted text."""
    if cancel_flag and cancel_flag.is_set():
        logging.info("OCR process cancelled by user.")
        return None

    if on_progress:
        try:
            on_progress(page_idx, doc.page_count)
        except Exception:
            logging.debug(
                "on_progress callback raised an exception:", exc_info=True
            )

    logging.info(f"OCR processing page {page_idx}/{doc.page_count} of '{pdf_path}'")
    img = None
    try:
        img = render_page_to_image(page, matrix)
        img = preprocess_for_ocr(img)
        text = run_ocr_on_image(img, lang, ocr_config, page_idx, pdf_path, timeout_sec)
        return text
    finally:
        if img is not None:
            try:
                img.close()
            except Exception:
                pass


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
            text = _ocr_single_page(
                page,
                i,
                doc,
                pdf_path,
                lang,
                ocr_config,
                matrix,
                timeout_sec,
                on_progress,
                cancel_flag,
            )
            if text is not None:
                texts.append(text)
            elif cancel_flag and cancel_flag.is_set():
                logging.info("OCR process cancelled by user.")
                break

        gc.collect()
    return texts


# Only splity the files here, extract information in another function
def separate_invoices(pdf_path, on_progress=None, cancel_flag=None):
    """
    Separate a multi-invoice PDF into individual invoices by OCRing each page and extracting relevant data.
    Returns a list of Invoice objects.
    """
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
        invoices.append(_parse_invoice_page(page, text, idx, pdf_path))
    return invoices


def build_address_block(rows: list[str]) -> str:
    """
    Build a text block containing "Aadress" line and (optionally) the next line if it looks like part of the address.
    """

    for i, row in enumerate(rows):
        if "aadress" in row.lower():
            address_block = row.strip()

            # Check next row for possible continuation
            if i + 1 < len(rows):
                next_row = rows[i + 1].strip()
                if re.search(r"\d", next_row) and "reg. kood" not in next_row:
                    address_block += " " + next_row
            return address_block
    raise ValidationError("Keyword 'aadress' not found in rows")


def _extract_apartment_from_address(address_block: str) -> tuple[str, str]:
    APARTMENT_RE = re.compile(r"\b(\d{1,3})-(\d+)\b")

    # Find apartment matches like '113-64' in that block
    matches = list(APARTMENT_RE.finditer(address_block))

    if matches:
        last_match = matches[-1]
        house_number, apt_number = last_match.groups()
        apartment = apt_number

        # Everything before the apartment number is the address
        before_apt = address_block[: last_match.start()].strip()
        address = f"{before_apt} {house_number}".strip()
    else:
        # No apartment match found, fallback to last part after splitting
        apartment = ""
        address = address_block
        logging.info(
            f"No apartment match found. Extracted address='{address}', apartment='{apartment}'"
        )
    return apartment, address


def extract_address_period_apartment(text):
    rows = text.splitlines()

    # --- Address & apartment ---
    address_block = build_address_block(rows)

    # Strip "Aadress" prefix
    after_label = re.split(r"aadress\s*[:\- ]\s*", address_block, flags=re.IGNORECASE)[
        -1
    ].strip()

    # Extract apartment number
    apartment, address = _extract_apartment_from_address(after_label)

    # Period
    period_parts = extract_parts(rows, "periood")
    period = period_parts[1] if len(period_parts) > 1 else ""

    # Year
    year_parts = extract_parts(rows, "kuupäev", pattern=r"[:\-\. ]+")
    year = year_parts[-1] if len(year_parts) > 1 else ""

    return {"address": address, "apartment": apartment, "period": period, "year": year}


# Find row keyword, split it, return list of stripped parts
def extract_parts(rows, keyword, pattern=r"[:\- ]+"):
    for i, row in enumerate(rows):
        if keyword in row.lower():
            parts = [part.strip().lower() for part in re.split(pattern, row) if part]

            if keyword == "aadress" and i + 1 < len(rows):
                next_row = rows[i + 1].strip().lower()
                if re.search(r"\d", next_row):
                    extra_parts = [
                        part.strip().lower()
                        for part in re.split(pattern, next_row)
                        if part
                    ]
                    parts.extend(extra_parts)
            return parts
    raise ValidationError(f"Keyword '{keyword}' not found in rows")


def save_each_invoice_as_file(invoices, dest):
    for invoice in invoices:
        writer = PdfWriter()
        writer.add_page(invoice.pdf_page)
        with open(dest / f"{invoice.apartment}.pdf", "wb") as f:
            writer.write(f)
    return dest
