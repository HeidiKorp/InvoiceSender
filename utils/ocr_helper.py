import os, sys, shutil, logging, fitz, io
import pytesseract
from tkinter import messagebox
from PIL import Image, ImageOps, ImageFilter

from utils.logging_helper import log_line


def get_tesseract_cmd():
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
        return os.path.join(base_dir, "_internal", "tesseract", "tesseract.exe")
    else:
        base_dir = shutil.which("tesseract")
        return base_dir
        

def check_ocr_environment():
    try:
        v = pytesseract.get_tesseract_version()
        log_line(f"Tesseract version={v}")
    except Exception as e:
        log_exception(e)
        messagebox.showerror(
            "Tesseract puudub",
            "Tesseract OCR ei ole selles arvutis paigaldatud või ei leitud teekonda.\n\n"
            f"Viga: {e}"
        )
        return False

    try:
        langs = pytesseract.get_languages(config="")
    except Exception as e:
        messagebox.showerror(
            "Tesseract viga",
            f"Tesseract on paigaldatud (versioon {v}), aga keelte nimekirja ei saanud lugeda.\n\nViga: {e}"
        )
        return False

    if "est" not in langs:
        messagebox.showerror(
            "Puuduv keel",
            "Tesseract OCR on paigaldatud, kuid 'est' (eesti) keeleandmed puuduvad.\n\n"
            "Paigalda Tesseract'i eesti keele toetus."
        )
        return False

    return True


def check_tesseract_lang(lang: str) -> None:
    # Check if language data is available
    try:
        available_langs = pytesseract.get_languages(config='')
        if lang not in available_langs:
            logging.warning(
                "Tesseract language '%s' not found in available languages: %s. "
                "Install this language (e.g. 'tesseract-ocr-%s') for best results.",
                lang,
                available_langs,
                lang,
            )
    except Exception as e:
        # Unable to query languages - log and continue
        logging.debug("Failed to query Tesseract languages.", exc_info=True)


def render_page_to_image(page: fitz.Page, matrix: fitz.Matrix) -> Image.Image:
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    png_bytes = pix.tobytes("png")

    # PIL load
    img = Image.open(io.BytesIO(png_bytes))
    del pix, png_bytes
    return img


def preprocess_for_ocr(img: Image.Image) -> Image.Image:
    # Preprocess: grayscale -> slight denoise -> autocontrast -> binarize
    img = img.convert("L")  # Grayscale\
    img = img.filter(ImageFilter.MedianFilter(size=3))  # Denoise
    img = ImageOps.autocontrast(img, cutoff=1)  # Autocontrast

    # Binarize 
    img = img.point(lambda x: 255 if x > 180 else 0, mode='1')
    return img


def run_ocr_on_image(img: Image.Image, lang: str, ocr_config: str, page_index: int, pdf_path: str, timeout_sec: int) -> str:
    # Run OCR on a preprocessed image and handle errors
    try:
        # OCR with timeout so a single page can't block the whole process
        text = pytesseract.image_to_string(
            img, lang=lang, config=ocr_config, timeout=timeout_sec
            ) or ""
        return text
    except pytesseract.TesseractError as e:
        # Show stderr from tesseract - helpful for missing lang and bad params
        logging.error(f"Tesseract failed on page {i}: {e}\n{getattr(e, 'stderr', '')}")
        stderr = getattr(e, "stderr", "")
        if stderr:
            logging.error("--- Tesseract stderr ---")
            logging.error(stderr)
        return "" # Return empty text on error
        
    except RuntimeError as e:
        # pytesseract timeout raises RuntimeError
        if "Timeout" in str(e):
            logging.error(f"OCR timeout on page {i} of '{pdf_path}' after {timeout_sec} seconds")
            return ""
        # Reraise other runtime errors
        raise

    except Exception as e:
        logging.error(f"Unexpected error on page {i} of '{pdf_path}': {e}")
        logging.error(traceback.format_exc())
        return ""