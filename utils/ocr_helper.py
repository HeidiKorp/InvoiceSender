import os, sys, shutil
import pytesseract
from utils.logging_helper import log_line
from tkinter import messagebox


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