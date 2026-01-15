from pypdf import PdfReader, PdfWriter
import pandas as pd
from pathlib import Path
import re
import io, gc, fitz, logging, sys, os
from PIL import Image, ImageOps, ImageFilter
import pytesseract
import traceback
import json

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

BLOCK_NEXTLINE_RE = re.compile(r"\b(reg\.?\s*kood|raamatupidamine|arveldus|iban|tel|e-?post)\b", re.IGNORECASE)
STREET_TOKEN_PATTERN = r"(tee|tn|tänav|pst|puiestee|mnt|küla|maantee)"
HOUSE_APT_RE = re.compile(
    rf"\b(?P<street>[A-Za-zÕÄÖÜõäöüšž\.\- ]+)\s+(?P<house>\d{{1,4}}[A-Za-z]?)\s*-\s*(?P<apartment>\d{{1,4}})\b",
    re.IGNORECASE
)


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


def save_debug_data(entry):
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))

    debug_file = os.path.join(base, "debug.log")
    # print(f"Debug file is at: {debug_file}")

    with open(debug_file, "a", encoding="utf-8") as f:
        if isinstance(entry, dict):
            f.write(json.dumps(entry, ensure_ascii=False, indent=2))
        else:
            f.write(str(entry))
        f.write("\n\n---\n\n")


def _parse_invoice_page(page, text: str, page_number: int, pdf_path: str, prev_apt) -> dict:
    # Debug part
    # debug_nr = [5, 52, 60, 61]

    _validate_page_text(text, page_number, pdf_path)
    client_data = extract_address_period_apartment(text, prev_apt)
    # if page_number in debug_nr:
    #     save_debug_data(text)
    #     save_debug_data(client_data)
        
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

    # test_invoices = [39, 40, 41]

    with fitz.open(pdf_path) as doc:
        total_pages = doc.page_count

        for i, page in enumerate(doc, start=1):
            # if i in test_invoices:
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
            else:
                print("Text is none! idx={i}")
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

    invoices: list[InvoiceItem] = []
    prev_apt: int | None = None

    for idx, (page, text) in enumerate(zip(reader.pages, page_texts), start=1):
        invoice = _parse_invoice_page(page, text, idx, pdf_path, prev_apt)
        invoices.append(invoice)

        try:
            prev_apt = int(invoice.apartment) if invoice.apartment else prev_apt
        except ValueError:
            pass
    return invoices


def _norm_row(row: str) -> str:
    row = row.replace("\u2014", "-").replace("\u2013", "-") # em/en dash -> hypen
    row = row.replace("\xa0", " ")
    row = re.sub(r"\s+", " ", row)
    return row.strip()


def contains_street_token(street_text: str) -> bool:
    return re.search(rf"\b{STREET_TOKEN_PATTERN}\b", street_text, re.IGNORECASE) is not None


def parse_house_apartment_from_row(row_text: str) -> dict | None:
    match = HOUSE_APT_RE.search(row_text)
    if not match:
        return None

    street_name = match.group("street").strip()
    house_number = match.group("house").strip()
    apartment_number = match.group("apartment").strip()

    if not contains_street_token(street_name):
        return None

    return {
        "street": street_name,
        "house": house_number,
        "apartment": apartment_number,
    }


def find_address_block(rows: list[str]) -> str:
    """
    Build a text block containing "Aadress" line and (optionally) the next line if it looks like part of the address.
    """
    for row_idx, row in enumerate(rows):
        norm_row = _norm_row(row)
        if "aadress" in norm_row.lower():
            return norm_row
    return None


def strip_address_prefix(address_block: str) -> str:
    parts = re.split(r"\baadress\b\s*[:\- ]\s*", address_block, flags=re.IGNORECASE, maxsplit=1)
    return parts[-1].strip() if parts else address_block.strip()


def score_candidate_row(row_text: str, apartment_nr: str, previous_apt: int | None) -> int:
    score_value = 0

    if BLOCK_NEXTLINE_RE.search(row_text):
        score_value -= 50

    if re.search(r"\btel\b", row_text, re.IGNORECASE):
        score_value += 8

    if len(apartment_nr) >= 2:
        score_value += 5
    else:
        score_value -= 5

    if previous_apt:
        try:
            apt_int = int(apartment_nr)
            score_value += 3 if apt_int >= previous_apt else -3
        except ValueError:
            score_value -= 2

    return score_value


def find_best_house_apartment_candidate(rows: list[str], previous_apt: int | None) -> dict | None:
    best_candidate = None
    best_score = -10_000

    for row_idx, row_text in enumerate(rows):
        norm_row = _norm_row(row_text)
        parsed = parse_house_apartment_from_row(norm_row)
        if not parsed:
            continue
        
        row_score = score_candidate_row(norm_row, parsed["apartment"], previous_apt)

        if row_score > best_score:
            print(f'Getting here! small score yayy')
            best_score = row_score
            best_candidate = {
                "row_idx": row_idx,
                "row_text": norm_row,
                **parsed,
                "score": row_score,
            }
    return best_candidate


def pick_address_and_apt(rows: list[str], prev_apt: int | None) -> tuple[str, str]:
    address_block = find_address_block(rows)

    address_candidate = None
    if address_block:
        address_payload = strip_address_prefix(address_block)
        address_candidate = parse_house_apartment_from_row(address_payload)

    best_candidate = find_best_house_apartment_candidate(rows, prev_apt)

    if best_candidate and (not address_candidate):
        return f"{best_candidate['street']} {best_candidate['house']}", best_candidate["apartment"]
    
    if best_candidate and address_candidate:
        # If mismatch, prefer global best (handles your “Aadress line garbled” cases)
        if address_candidate["apartment"] != best_candidate["apartment"]:
            return f"{best_candidate['street']} {best_candidate['house']}", best_candidate["apartment"]

    # Otherwise trust Aadress
    return f"{address_candidate['street']} {address_candidate['house']}", address_candidate["apartment"]

    if address_candidate:
        return f"{address_candidate['street']} {address_candidate['house']}", address_candidate["apartment"]

    return "", ""


def extract_address_period_apartment(text, prev_apt):
    rows = text.splitlines()

    # --- Address & apartment ---
    address, apartment = pick_address_and_apt(rows, prev_apt)

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
