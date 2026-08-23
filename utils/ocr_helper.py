import os, sys, shutil, logging, fitz, io, traceback
import pytesseract
from tkinter import messagebox
from PIL import Image, ImageOps, ImageFilter

from utils.logging_helper import log_exception


def check_ocr_environment():
    version = _check_tesseract_version()
    if not version:
        return False

    langs = _check_tesseract_languages(version)
    if not langs:
        return False

    if "est" not in langs:
        messagebox.showerror(
            "Puuduv keel",
            "Tesseract OCR on paigaldatud, kuid 'est' (eesti) keeleandmed puuduvad.\n\n"
            "Paigalda Tesseract'i eesti keele toetus.",
        )
        return False

    return True


def get_tesseract_cmd():
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
        return os.path.join(base_dir, "_internal", "tesseract", "tesseract.exe")
    return shutil.which("tesseract")


def check_tesseract_lang(lang: str) -> None:
    try:
        available_langs = pytesseract.get_languages(config="")
        if lang not in available_langs:
            logging.warning(
                "Tesseract language '%s' not found in available languages: %s.",
                lang,
                available_langs,
            )
    except Exception as e:
        log_exception(e, operation="tesseract_lang_check")


def render_page_to_image(page: fitz.Page, matrix: fitz.Matrix) -> Image.Image:
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    png_bytes = pix.tobytes("png")
    img = Image.open(io.BytesIO(png_bytes))
    del pix, png_bytes
    return img


def preprocess_for_ocr(img: Image.Image) -> Image.Image:
    img = img.convert("L")
    img = img.filter(ImageFilter.MedianFilter(size=3))
    img = ImageOps.autocontrast(img, cutoff=1)
    return img.point(lambda x: 255 if x > 180 else 0, mode="1")


def run_ocr_on_image(
    img: Image.Image,
    lang: str,
    ocr_config: str,
    page_index: int,
    pdf_path: str,
    timeout_sec: int,
) -> str:
    try:
        return (
            pytesseract.image_to_string(
                img, lang=lang, config=ocr_config, timeout=timeout_sec
            )
            or ""
        )
    except pytesseract.TesseractError as e:
        log_exception(e, operation=f"ocr_page:{page_index}:{pdf_path}")
        return ""
    except RuntimeError as e:
        if "Timeout" in str(e):
            log_exception(e, operation=f"ocr_timeout:{page_index}:{pdf_path}")
            return ""
        raise
    except Exception as e:
        log_exception(e, operation=f"ocr_page:{page_index}:{pdf_path}")
        logging.error(traceback.format_exc())
        return ""


def _check_tesseract_version():
    try:
        return pytesseract.get_tesseract_version()
    except Exception as e:
        log_exception(e, operation="tesseract_version")
        messagebox.showerror(
            "Tesseract puudub",
            "Tesseract OCR ei ole selles arvutis paigaldatud või ei leitud teekonda.\n\n"
            f"Viga: {e}",
        )
        return None


def _check_tesseract_languages(version):
    try:
        return pytesseract.get_languages(config="")
    except Exception as e:
        log_exception(e, operation="tesseract_languages")
        messagebox.showerror(
            "Tesseract viga",
            f"Tesseract on paigaldatud (versioon {version}), aga keelte nimekirja ei saanud lugeda.\n\nViga: {e}",
        )
        return None
