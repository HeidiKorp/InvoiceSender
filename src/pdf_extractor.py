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

ADMIN_LINE = r"\b(reg\.?\s*kood|raamatupidamine|iban|e-?post)\b"

STREET_TOKEN_PATTERN = r"(tee|tn|tänav|pst|puiestee|mnt|küla|maantee)"
HOUSE_APT_RE = re.compile(
    r"\b(?P<street>[A-Za-zÕÄÖÜõäöüšž\.\- ]+)\s+"
    r"(?P<house>\d{1,4}[A-Za-z]?)\s*-\s*"
    r"(?P<apartment>\d{1,4})"
    r"(?:\s*(?P<apt_suffix>[A-Za-z]))?\b",
    re.IGNORECASE,
)
HOUSE_ONLY_APT_RE = re.compile(
    r"\b(?P<house>\d{1,4}[A-Za-z]?)\s*-\s*"
    r"(?P<apartment>\d{1,4})"
    r"(?:\s*(?P<apt_suffix>[A-Za-z]))?\b",
    re.IGNORECASE,
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

    with open(debug_file, "a", encoding="utf-8") as f:
        if isinstance(entry, dict):
            f.write(json.dumps(entry, ensure_ascii=False, indent=2))
        else:
            f.write(str(entry))
        f.write("\n\n---\n\n")


def _parse_invoice_page(
    page, text: str, page_number: int, pdf_path: str, prev_apt
) -> dict:

    _validate_page_text(text, page_number, pdf_path)
    client_data = extract_address_period_apartment(text, prev_apt)

    return InvoiceItem(
        pdf_page=page,
        address=client_data["address"],
        period=client_data["period"],
        apartment=client_data["apartment"],
        year=client_data["year"],
    )


def _ocr_single_page(
    page,
    page_idx,
    doc,
    pdf_path,
    lang,
    ocr_config,
    matrix,
    timeout_sec,
    on_progress,
    cancel_flag,
):
    """OCR a single page and return the extracted text."""
    if cancel_flag and cancel_flag.is_set():
        logging.info("OCR process cancelled by user.")
        return None

    if on_progress:
        try:
            on_progress(page_idx, doc.page_count)
        except Exception:
            logging.debug("on_progress callback raised an exception:", exc_info=True)

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

    # Debug
    # test_invoices = [159, 188, 216, 187, 149, 157, 159, 164, 189, 211, 24, 47, 52, 74]
    # test_invoices = [5, 52, 139, 164]

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
            # print(f"Full text:\n{text}")
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


    invoices: list[InvoiceItem] = []
    prev_apt: int | None = None

    for idx, (page, text) in enumerate(zip(reader.pages, page_texts), start=1):
        # print(f"Full text: \n{text}\n")
        invoice = _parse_invoice_page(page, text, idx, pdf_path, prev_apt)
        invoices.append(invoice)
        # print("\n")
        try:
            prev_apt = apartment_numeric_part(str(invoice.apartment)) or prev_apt
        except ValueError:
            pass
    return invoices


def _norm_row(row: str) -> str:
    row = row.replace("\u2014", "-").replace("\u2013", "-")  # em/en dash -> hypen
    row = row.replace("\xa0", " ")
    row = re.sub(r"\s+", " ", row)
    return row.strip()


def _norm_apartment(apartment_nr: str, apartment_suffix: str | None) -> str:
    if not apartment_nr:
        return ""
    if apartment_suffix:
        return f"{apartment_nr}{apartment_suffix.upper()}"
    return apartment_nr


def contains_street_token(street_text: str) -> bool:
    return (
        re.search(rf"\b{STREET_TOKEN_PATTERN}\b", street_text, re.IGNORECASE)
        is not None
    )

def infer_likely_house(rows: list[str]) -> str | None:
    best_house = None
    best_score = -10_000

    for row_text in rows:
        normalized_row = _norm_row(row_text)
        parsed = parse_house_apartment_from_loose(normalized_row)
        if not parsed:
            continue

        score = 0
        score += 10 if parsed.get("has_street") else 0
        score += 6 if re.search(r"\btel\b", normalized_row, re.IGNORECASE) else 0

        # mild penalty for crazy-large house numbers (e.g. 413 vs 43 for this dataset)
        if parsed.get("house", "").isdigit() and int(parsed["house"]) >= 300:
            score -= 8

        if score > best_score:
            best_score = score
            best_house = parsed["house"]

    return best_house


def should_strip_apartment_suffix(row_text: str) -> bool:
    """
    Return True only when suffix is likely OCR noise (e.g. 'A P 11').
    Keep suffix for real apartments like '73 A'.
    """
    match = re.search(r"\b-\s*(\d{1,4})\s*([A-Za-z])\b(.*)$", _norm_row(row_text))
    if not match:
        return False

    trailing_text = (match.group(3) or "").strip()

    # Noise patterns: "A P ..." or "A AP ..." etc.
    if re.match(r"^(p|ap)\b", trailing_text, flags=re.IGNORECASE):
        return True
    if re.match(r"^[A-Za-z]\b", trailing_text):  # "A P ..." (next token is a letter)
        return True

    return False

def parse_house_apartment_from_row(row_text: str) -> dict | None:
    match = HOUSE_APT_RE.search(row_text)

    if not match:
        return None

    street_name = match.group("street").strip()
    house_number = match.group("house").strip()
    apartment_number = match.group("apartment").strip()
    apartment_suffix = match.group("apt_suffix")

    if not contains_street_token(street_name):
        return None
    
    apartment_value = _norm_apartment(apartment_number, apartment_suffix)
    
    return {
        "street": street_name,
        "house": house_number,
        "apartment": apartment_value,
        "has_street": True,
    }


def parse_house_apartment_from_loose(row_text: str) -> dict | None:
    match = HOUSE_APT_RE.search(row_text)

    if match:
        street_name = match.group("street").strip()
        house_number = match.group("house").strip()
        apartment_number = match.group("apartment").strip()
        apartment_suffix = match.group("apt_suffix")

        return {
            "street": street_name,
            "house": house_number,
            "apartment": _norm_apartment(apartment_number, apartment_suffix),
            "has_street": contains_street_token(street_name),
        }

    match2 = HOUSE_ONLY_APT_RE.search(row_text)
    if not match2:
        return None

    return {
        "street": "",
        "house": match2.group("house").strip(),
        "apartment": _norm_apartment(match2.group("apartment").strip(), match2.group("apt_suffix")),
        "has_street": False,
    }


def get_house_from_address_candidate(address_candidate: dict | None) -> str | None:
    # Prefer candidates with the same house; extract expected house from address line even if apt is wrong
    if not address_candidate:
        return None
    return address_candidate.get("house") or None


def find_best_address_block(rows: list[str], prev_apt: int | None = None) -> str | None:
    best_row_text = None
    best_score = -10_000

    for row_text in rows:
        normalized_row = _norm_row(row_text)
        if "aadress" not in normalized_row.lower():
            continue

        address_payload = strip_address_prefix(normalized_row)
        parsed = (
            parse_house_apartment_from_row(address_payload)
            or parse_house_apartment_from_loose(address_payload)
        )
        if not parsed:
            continue

        score = 0

        # Strong preference: street token recognized
        score += 10 if parsed.get("has_street") else 0

        # Avoid obvious OCR house-number explosions (e.g. 413 instead of 43)
        house_value = parsed.get("house", "")
        if house_value.isdigit() and int(house_value) >= 300:
            score -= 8

        # If prev_apt available: prefer small forward jumps, avoid huge jumps
        apartment_int = apartment_numeric_part(parsed.get("apartment", ""))
        if prev_apt is not None and apartment_int is not None:
            jump = apartment_int - prev_apt
            if 0 <= jump <= 3:
                score += 12
            elif jump > 10:
                score -= 8
            elif jump < 0:
                score -= 12

        if score > best_score:
            best_score = score
            best_row_text = normalized_row

    return best_row_text


def strip_address_prefix(address_block: str) -> str:
    parts = re.split(
        r"\baadress\b\s*[:;\-\. ]\s*", address_block, flags=re.IGNORECASE, maxsplit=1
    )
    return parts[-1].strip() if parts else address_block.strip()


def apartment_numeric_part(apartment_value: str) -> int | None:
    match = re.match(r"^\s*(\d+)", apartment_value)
    return int(match.group(1)) if match else None


def _iter_parsed_candidates(rows: list[str]) -> list[dict]:
    """
    Parse every row into a candidate dict.
    Returns a list of dicts that include row_idx + row_text + parsed fields.
    """
    parsed_candidates: list[dict] = []

    for row_index, row_text in enumerate(rows):
        normalized_row_text = _norm_row(row_text)

        parsed = parse_house_apartment_from_loose(normalized_row_text)
        if not parsed:
            continue

        parsed_candidates.append(
            {
                "row_idx": row_index,
                "row_text": normalized_row_text,
                **parsed,
            }
        )

    return parsed_candidates


def _is_better_candidate(new_candidate: dict, new_score: int, best_candidate: dict | None, best_score: int) -> bool:
    """
    Decide if new candidate should replace current best.
    Tie-breaker: earlier row wins if scores equal.
    """
    if best_candidate is None:
        return True
    if new_score > best_score:
        return True
    if new_score < best_score:
        return False

    # Tie: prefer earlier occurrence
    return new_candidate["row_idx"] < best_candidate["row_idx"]


def _pick_best_candidate(
    parsed_candidates: list[dict],
    previous_apartment: int | None,
    expected_house: str | None,
) -> dict | None:
    best_candidate = None
    best_score = -10_000

    for candidate in parsed_candidates:
        candidate_score = score_candidate_row(
            candidate["row_text"], candidate, previous_apartment, expected_house
        )

        if _is_better_candidate(candidate, candidate_score, best_candidate, best_score):
            best_candidate = {**candidate, "score": candidate_score}
            best_score = candidate_score

    return best_candidate


def score_candidate_row(
    row_text: str, parsed: dict, previous_apt: int | None, expected_house: str | None
) -> int:
    score = 0

    if re.search(ADMIN_LINE, row_text, re.IGNORECASE):
        score -= 5

    apartment_nr = parsed["apartment"]
    numeric_part = apartment_numeric_part(apartment_nr)

    if numeric_part is not None:
        score += 6 if numeric_part >= 10 else 2
    else:
        score -= 2

    # if we know the expected house nr, prefer is strongly
    if expected_house and parsed["house"] == expected_house:
        score += 20

    # if full street token present, small bonus
    if parsed.get("has_street"):
        score += 4

    if previous_apt is not None and numeric_part is not None:
        jump = numeric_part - previous_apt

        # Prefer small forward steps (4->5, 10->11)
        if 0 <= jump <= 3:
            score += 12

        # Still allow forward jumps, but penalize big ones
        if jump > 10:
            score -= 8

        # Penalize going backwards
        if jump < 0:
            score -= 12

    return score


def has_reasonable_apartment_digits(apartment_value: str) -> bool:
    digits = apartment_numeric_part(apartment_value)
    return digits is not None and digits >= 10  # 2+ digits


def find_best_house_apartment_candidate(
    rows: list[str],
    previous_apartment: int | None,
    expected_house: str | None,
) -> dict | None:
    parsed_candidates = _iter_parsed_candidates(rows)
    return _pick_best_candidate(parsed_candidates, previous_apartment, expected_house)


def format_address(candidate) -> str:
    if candidate is None:
        return ""

    if isinstance(candidate, dict):
        street = (candidate.get("street") or "").strip()
        house = (candidate.get("house") or "").strip()
    else:
        street = (getattr(candidate, "street", "") or "").strip()
        house = (getattr(candidate, "house", "") or "").strip()
    return f"{street} {house}".strip() if street else house


def choose_expected_house(rows: list[str], address_candidate: dict | None) -> str | None:
    if not address_candidate:
        return infer_likely_house(rows)

    house_value = address_candidate.get("house")
    has_street = bool(address_candidate.get("has_street"))

    if not house_value:
        return infer_likely_house(rows)

    if house_value.isdigit() and int(house_value) >= 300:
        return infer_likely_house(rows)

    if not has_street:
        return infer_likely_house(rows)

    return house_value


def choose_apartment_value(address_candidate: dict, best_candidate: dict) -> str:
    """
    If numeric parts match (24 vs 24A), keep suffix unless it looks like noise.
    Otherwise, prefer best candidate's apartment.
    """
    best_apartment = best_candidate.get("apartment", "") or ""
    address_apartment = address_candidate.get("apartment", "") or ""

    best_numeric = apartment_numeric_part(best_apartment)
    address_numeric = apartment_numeric_part(address_apartment)

    if best_numeric and address_numeric and best_numeric == address_numeric:
        # Only strip suffix when it looks like OCR noise in the *best candidate row*
        if should_strip_apartment_suffix(best_candidate.get("row_text", "")):
            return best_numeric
        return best_apartment  # keep e.g. 73A

    return best_apartment or address_apartment
    

def pick_address_and_apt(rows: list[str], prev_apt: int | None) -> tuple[str, str]:
    address_block = find_best_address_block(rows, prev_apt)
    # print(f"Address block: {address_block}")
    
    address_candidate = None
    if address_block:
        address_payload = strip_address_prefix(address_block)
        address_candidate = parse_house_apartment_from_row(address_payload) or parse_house_apartment_from_loose(address_payload)

    expected_house = choose_expected_house(rows, address_candidate)    
    best_candidate = find_best_house_apartment_candidate(rows, prev_apt, expected_house)

    # logging.info(f"Aadress candidate: {address_candidate}")
    # logging.info(f"Best candidate: {best_candidate}")

    if best_candidate and not address_candidate:
        return format_address(best_candidate), best_candidate["apartment"]

    if best_candidate and address_candidate:
        chosen_apartment = choose_apartment_value(address_candidate, best_candidate)

        if address_candidate["apartment"] != best_candidate["apartment"]:
            return format_address(best_candidate), chosen_apartment
        return format_address(address_candidate), chosen_apartment

    # print(f"Address candidate: {address_candidate}\n")
    if address_candidate:
        return format_address(address_candidate), address_candidate["apartment"]

    return "", ""


def extract_address_period_apartment(text, prev_apt):
    rows = text.splitlines()

    # --- Address & apartment ---
    address, apartment = pick_address_and_apt(rows, prev_apt)
    # print(f"Chose address and apt: {address} {apartment}")

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
