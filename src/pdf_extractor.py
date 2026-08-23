from pypdf import PdfReader, PdfWriter
import re
from typing import Optional
import gc, fitz, logging, sys, os
import pytesseract
import json

from src.data_classes import ValidationError, is_valid_period, is_valid_year
from utils.logging_helper import log_exception
from utils.ocr_helper import (
    check_tesseract_lang,
    render_page_to_image,
    preprocess_for_ocr,
    run_ocr_on_image,
)
from src.data_classes import InvoiceItem

ADMIN_LINE = r"\b(reg\.?\s*kood|raamatupidamine|iban|e-?post)\b"

STREET_TOKEN_PATTERN = r"(tee|tn|tänav|pst|puiestee|mnt|küla|maantee|põik|allee)"
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
KW_RE = re.compile(r"(?i)\b(kü|korteriühistu)\b")
REMOVE_KW_RE = re.compile(r"(?i)\b(?:kü|korteriühistu)\b")
EDGE_CLEAN_RE = re.compile(r"^[\s,;:\-–—]+|[\s,;:\-–—]+$")

CUTOFF_RE = re.compile(
    r"(?i)\b(?:aadress|address|viitenumber|kuupäev|kp|arve|arve\s*nr|telefon|tel|e-?post|email)\s*:"
)

logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(message)s",
)


class PdfInvoiceExtractor:
    def load(self, pdf_path, on_progress=None, cancel_event=None) -> list[InvoiceItem]:
        return separate_invoices(
            pdf_path, on_progress=on_progress, cancel_flag=cancel_event
        )

    def save(self, invoice_batch, on_progress=None):
        dest = invoice_batch.dest_dir
        for invoice in invoice_batch.invoices:
            writer = PdfWriter()
            writer.add_page(invoice.pdf_page)
            with open(dest / f"{invoice.apartment}.pdf", "wb") as f:
                writer.write(f)
        return dest


def separate_invoices(pdf_path, on_progress=None, cancel_flag=None):
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
    period, year = None, None

    for idx, (page, text) in enumerate(zip(reader.pages, page_texts), start=1):
        invoice = _parse_invoice_page(page, text, idx, pdf_path, prev_apt, period, year)
        period = invoice.period
        year = invoice.year
        invoices.append(invoice)
        try:
            prev_apt = apartment_numeric_part(str(invoice.apartment)) or prev_apt
        except ValueError:
            pass
    return invoices


def ocr_pdf_all_pages(
    pdf_path: str,
    lang: str = "est",
    dpi: int = 300,
    psm: int = 6,
    oem: int = 1,
    timeout_sec: int = 120,
    on_progress=None,
    cancel_flag=None,
) -> list[str]:
    logging.info("Using tesseract_cmd=%s", pytesseract.pytesseract.tesseract_cmd)

    texts: list[str] = []
    scale = dpi / 72
    matrix = fitz.Matrix(scale, scale)

    check_tesseract_lang(lang)
    ocr_config = f"--oem {oem} --psm {psm}"

    with fitz.open(pdf_path) as doc:
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


def extract_address_period_apartment(text, prev_apt, period=None, year=None):
    rows = text.splitlines()

    address, apartment = pick_address_and_apt(rows, prev_apt)
    if not address or not apartment:
        _log_weak_address_parse(rows, address, apartment)

    if not is_valid_period(period):
        period = extract_period(rows)

    if not is_valid_year(year):
        year = extract_year(rows)

    return {"address": address, "apartment": apartment, "period": period, "year": year}


def pick_address_and_apt(rows: list[str], prev_apt: int | None) -> tuple[str, str]:
    address_block = find_best_address_block(rows, prev_apt)

    address_candidate = None
    if address_block:
        address_payload = strip_address_prefix(address_block)
        address_candidate = parse_house_apartment_from_row(address_payload) or parse_house_apartment_from_loose(address_payload)

    expected_house = choose_expected_house(rows, address_candidate)
    best_candidate = find_best_house_apartment_candidate(rows, prev_apt, expected_house)

    if best_candidate and not address_candidate:
        return format_address(best_candidate), best_candidate["apartment"]

    if best_candidate and address_candidate:
        chosen_apartment = choose_apartment_value(address_candidate, best_candidate)

        if address_candidate["apartment"] != best_candidate["apartment"]:
            if address_candidate.get("has_street"):
                return format_address(address_candidate), (
                    address_candidate["apartment"] or chosen_apartment
                )
            return format_address(best_candidate), chosen_apartment
        return format_address(address_candidate), chosen_apartment

    if address_candidate:
        return format_address(address_candidate), address_candidate["apartment"]

    return "", ""


def extract_ky_name_from_text(text: str) -> Optional[str]:
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if not KW_RE.search(line):
            continue

        name = extract_ky_name_from_line(line)
        if name:
            return name

    return None


def extract_ky_name_from_line(line: str) -> Optional[str]:
    if not KW_RE.search(line):
        return None

    line = line.splitlines()[0]

    if "," in line:
        line = line.split(",")[-1]

    remainder = REMOVE_KW_RE.sub("", line)

    part = CUTOFF_RE.search(remainder)
    if part:
        remainder = remainder[: part.start()]

    remainder = re.sub(r"\s+", " ", remainder).strip()
    remainder = EDGE_CLEAN_RE.sub("", remainder).strip()

    if not remainder:
        return None

    return remainder


def extract_period(rows: list[str]) -> str:
    return _first_matching_value(rows, "periood", is_valid_period, pattern=r"[:\- ]+")


def extract_year(rows: list[str]) -> str:
    return _first_matching_value(
        rows, "kuupäev", is_valid_year, pattern=r"[:\-\. ]+"
    )


def extract_parts(rows, keyword, pattern=r"[:\- ]+"):
    for i, row in enumerate(rows):
        if keyword in row.lower():
            parts = [part.strip().lower() for part in re.split(pattern, row) if part]

            if i + 1 < len(rows):
                next_row = rows[i + 1].strip().lower()
                if _should_include_next_row(keyword, next_row):
                    extra_parts = [
                        part.strip().lower()
                        for part in re.split(pattern, next_row)
                        if part
                    ]
                    parts.extend(extra_parts)
            return parts
    return []


def _first_matching_value(rows, keyword, is_valid, pattern=r"[:\- ]+") -> str:
    parts = extract_parts(rows, keyword, pattern=pattern)
    for part in parts:
        if is_valid(part):
            return part
    return ""


def _should_include_next_row(keyword: str, next_row: str) -> bool:
    if not next_row:
        return False
    if keyword == "aadress":
        return re.search(r"\d", next_row) is not None
    return keyword in {"periood", "kuupäev"}


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
        score += 10 if parsed.get("has_street") else 0

        house_value = parsed.get("house", "")
        if house_value.isdigit() and int(house_value) >= 300:
            score -= 8

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


def find_best_house_apartment_candidate(
    rows: list[str],
    previous_apartment: int | None,
    expected_house: str | None,
) -> dict | None:
    parsed_candidates = _iter_parsed_candidates(rows)
    return _pick_best_candidate(parsed_candidates, previous_apartment, expected_house)


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
    best_apartment = best_candidate.get("apartment", "") or ""
    address_apartment = address_candidate.get("apartment", "") or ""

    best_numeric = apartment_numeric_part(best_apartment)
    address_numeric = apartment_numeric_part(address_apartment)

    if best_numeric and address_numeric and best_numeric == address_numeric:
        if should_strip_apartment_suffix(best_candidate.get("row_text", "")):
            return best_numeric
        return best_apartment

    return best_apartment or address_apartment


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

        if parsed.get("house", "").isdigit() and int(parsed["house"]) >= 300:
            score -= 8

        if score > best_score:
            best_score = score
            best_house = parsed["house"]

    return best_house


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


def strip_address_prefix(address_block: str) -> str:
    parts = re.split(
        r"\baadress\b\s*[:;\-\.]?\s*",
        address_block,
        flags=re.IGNORECASE,
        maxsplit=1,
    )
    return parts[-1].strip() if parts else address_block.strip()


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

    if expected_house and parsed["house"] == expected_house:
        score += 20

    if parsed.get("has_street"):
        score += 4

    if _looks_like_date_pair(parsed.get("house", ""), parsed.get("apartment", "")):
        score -= 15

    if previous_apt is not None and numeric_part is not None:
        jump = numeric_part - previous_apt

        if 0 <= jump <= 3:
            score += 12

        if jump > 10:
            score -= 8

        if jump < 0:
            score -= 12

    return score


def should_strip_apartment_suffix(row_text: str) -> bool:
    match = re.search(r"\b-\s*(\d{1,4})\s*([A-Za-z])\b(.*)$", _norm_row(row_text))
    if not match:
        return False

    trailing_text = (match.group(3) or "").strip()

    if re.match(r"^(p|ap)\b", trailing_text, flags=re.IGNORECASE):
        return True
    if re.match(r"^[A-Za-z]\b", trailing_text):
        return True

    return False


def contains_street_token(street_text: str) -> bool:
    return (
        re.search(rf"\b{STREET_TOKEN_PATTERN}\b", street_text, re.IGNORECASE)
        is not None
    )


def apartment_numeric_part(apartment_value: str) -> int | None:
    match = re.match(r"^\s*(\d+)", apartment_value)
    return int(match.group(1)) if match else None


def has_reasonable_apartment_digits(apartment_value: str) -> bool:
    digits = apartment_numeric_part(apartment_value)
    return digits is not None and digits >= 1


def get_house_from_address_candidate(address_candidate: dict | None) -> str | None:
    if not address_candidate:
        return None
    return address_candidate.get("house") or None


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
    page, text: str, page_number: int, pdf_path: str, prev_apt, period=None, year=None
) -> dict:
    _validate_page_text(text, page_number, pdf_path)
    client_data = extract_address_period_apartment(text, prev_apt, period, year)
    ky_name = extract_ky_name_from_text(text)

    return InvoiceItem(
        pdf_page=page,
        address=client_data["address"],
        period=client_data["period"],
        apartment=client_data["apartment"],
        year=client_data["year"],
        ky_name=ky_name
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


def _validate_page_text(text: str, page_number: int, pdf_path: str):
    if not text or not text.strip():
        logging.error(
            f"SEPARATE_INVOICES: EMPTY text for page {page_number} of '{pdf_path}'"
        )
        raise ValidationError(
            f"PDF faili '{pdf_path}' leheküljelt {page_number} ei õnnestunud teksti lugeda ka pärast OCR-i. "
            "PDF võib olla vigane."
        )


def _iter_parsed_candidates(rows: list[str]) -> list[dict]:
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


def _is_better_candidate(new_candidate: dict, new_score: int, best_candidate: dict | None, best_score: int) -> bool:
    if best_candidate is None:
        return True
    if new_score > best_score:
        return True
    if new_score < best_score:
        return False

    return new_candidate["row_idx"] < best_candidate["row_idx"]


def _looks_like_date_pair(house: str, apartment: str) -> bool:
    house_text = str(house or "").strip()
    apt_text = str(apartment or "").strip()
    if not house_text.isdigit() or not apt_text.isdigit():
        return False
    house_num = int(house_text)
    apt_num = int(apt_text)
    if len(house_text) == 4 and 1900 <= house_num <= 2100:
        return True
    if len(apt_text) == 4 and 1900 <= apt_num <= 2100:
        return True
    if house_text.startswith("0") and apt_text.startswith("0"):
        return True
    return False


def _log_weak_address_parse(rows: list[str], address: str, apartment: str) -> None:
    snippets: list[str] = []
    for index, row in enumerate(rows):
        if "aadress" in row.lower():
            start = max(0, index - 1)
            end = min(len(rows), index + 3)
            snippets = rows[start:end]
            break
    if not snippets:
        snippets = rows[:8]
    excerpt = "\n".join(snippets)
    log_exception(
        RuntimeError(
            f"Nõrk aadressi parsimine: address={address!r} apartment={apartment!r}\n{excerpt}"
        ),
        operation="ocr_address_parse",
    )


def _norm_row(row: str) -> str:
    row = row.replace("\u2014", "-").replace("\u2013", "-")
    row = row.replace("\xa0", " ")
    row = re.sub(r"\s+", " ", row)
    return row.strip()


def _norm_apartment(apartment_nr: str, apartment_suffix: str | None) -> str:
    if not apartment_nr:
        return ""
    if apartment_suffix:
        return f"{apartment_nr}{apartment_suffix.upper()}"
    return apartment_nr
