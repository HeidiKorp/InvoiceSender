from pypdf import PdfReader, PdfWriter
import pandas as pd
from pathlib import Path
import re
import io, gc, fitz, logging, sys, os
from PIL import Image, ImageOps, ImageFilter
import pytesseract
import traceback

from src.xls_extractor import ValidationError


logging.basicConfig(
    level=logging.INFO,                # 👈 enables INFO and above
    format='[%(levelname)s] %(message)s'
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
    on_progress=None, # callback: on_progress(page_number: int, total_pages: int)
    cancel_flag=None # optional threading.Event to signal cancellation
    ) -> list[str]:
    """
    OCR a single page from an open fitz.Document (PyMuPDF).
    Returns extracted text (may be empty).
    """

    # log_line(f"Using tesseract_cmd={pytesseract.pytesseract.tesseract_cmd}")

    texts: list[str] = []
    scale = dpi / 72  # 72 is the default resolution
    matrix = fitz.Matrix(scale, scale)

    try:
        pytesseract.get_languages(config='')
        if lang not in pytesseract.get_languages(config=''):
            logging.warning(
                f"'{lang}' language data not found in Tesseract. "
                "Install it (e.g., 'tesseract-ocr-est') for best results."
            )
    except Exception as e:
        pass

    ocr_config = f"--oem {oem} --psm {psm}"
    with fitz.open(pdf_path) as doc:
        for i, page in enumerate(doc, start=1):
            if cancel_flag and cancel_flag.is_set():
                logging.info("OCR process cancelled by user.")
                break

            if on_progress:
                try:
                    on_progress(i, doc.page_count)
                except Exception:
                    pass


            logging.info(f"OCR processing page {i}/{doc.page_count} of '{pdf_path}'")
            img = None
            pix = None
            try:
                pix = page.get_pixmap(matrix=matrix, alpha=False)
                png_bytes = pix.tobytes("png")

                # PIL load
                img = Image.open(io.BytesIO(png_bytes))

                # Preprocess: grayscale -> slight denoise -> autocontrast -> binarize
                img = img.convert("L")  # Grayscale\
                img = img.filter(ImageFilter.MedianFilter(size=3))  # Denoise
                img = ImageOps.autocontrast(img, cutoff=1)  # Autocontrast

                # Binarize 
                img = img.point(lambda x: 255 if x > 180 else 0, mode='1')

                # OCR with timeout so a single page can't block the whole process
                text = pytesseract.image_to_string(
                    img, lang=lang, config=ocr_config, timeout=timeout_sec
                    ) or ""
                texts.append(text)
            except pytesseract.TesseractError as e:
                # Show stderr from tesseract - helpful for missing lang and bad params
                logging.error(f"Tesseract failed on page {i}: {e}\n{getattr(e, 'stderr', '')}")
                stderr = getattr(e, "stderr", "")
                if stderr:
                    log_line("--- Tesseract stderr ---")
                    log_line(stderr)
                texts.append("") # Append empty text on error
                
            except RuntimeError as e:
                # pytesseract timeout raises RuntimeError
                if "Timeout" in str(e):
                    logging.error(f"OCR timeout on page {i} of '{pdf_path}' after {timeout_sec} seconds")
                else:
                    raise

            except Exception as e:
                logging.error(f"Unexpected error on page {i} of '{pdf_path}': {e}")
                log_line(traceback.format_exc())
                texts.append("")
            finally:
                try:
                    if img is not None:
                        img.close()
                except Exception:
                    pass
            
                if "png_bytes" in locals():
                    del png_bytes
                del pix
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
        page_texts = ocr_pdf_all_pages(pdf_path, "est", dpi=300, on_progress=on_progress, cancel_flag=cancel_flag)
    else:
        page_texts = ocr_pdf_all_pages(pdf_path, "est", dpi=300, cancel_flag=cancel_flag)
    reader = PdfReader(pdf_path)

    if len(page_texts) != len(reader.pages) and not cancel_flag:
        raise ValidationError(f"PDF faili '{pdf_path}' OCR-tulemus on ebajärjekindel (lehtede arv ei klapi).")
    
    invoices = []
    for idx, (page, text) in enumerate(zip(reader.pages, page_texts), start=1):       
        rows = (text or "").splitlines()
        if len(text.strip()) == 0:
            log_line(f"SEPARATE_INVOICES: EMPTY text for page {idx} of '{pdf_path}'")
            raise ValidationError(
                f"PDF faili '{pdf_path}' leheküljelt {idx} ei õnnestunud teksti lugeda ka pärast OCR-i. "
                "PDF võib olla vigane.")
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
