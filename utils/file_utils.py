from pathlib import Path
from tkinter import messagebox
import shutil, os, sys
import configparser
from dataclasses import dataclass

@dataclass(frozen=True)
class InvoiceType:
    key: str
    label: str
    subject: str
    body: str


def load_app_version(config):
    return config.get("app", "VERSION", fallback="1.0.0")


def load_app_name(config):
    config.get("app", "NAME", fallback="Arvete Saatja")
    

def load_invoice_types(config):
    """Loads two types from config.cfg"""
    hint = config.get("ui", "TYPE_HINT")

    def read_section(section: str) -> InvoiceType:
        return InvoiceType(
            key=config.get(section, "KEY"),
            label=config.get(section, "LABEL"),
            subject=config.get(section, "SUBJECT"),
            body=config.get(section, "BODY").replace("\\n", "\n")
        )
    t1 = read_section("invoice_type_kommunaal")
    t2 = read_section("invoice_type_kyte")

    types = {t1.key: t1, t2.key: t2}
    return types, hint


def read_config():
    config = configparser.ConfigParser()
    config_path = Path(__file__).parent.parent / "config.cfg"

    try:
        with config_path.open("r", encoding="utf-8") as f:
            config.read_file(f)
        return config
    except UnicodeDecodeError:
        # Fallback if the file was saved in legacy Windows encoding
        with config_path.open("r", encoding="cp1252") as f:
            config.read_file(f)
        return config


def delete_folder(root, path_str):
    path = Path(path_str)

    if not path.exists() or not path.is_dir():
        messagebox.showerror("Viga", f"Kausta ei leitud:\n{path_str}")
        return
    
    # Ask the user first
    confirm = messagebox.askyesno("Kustuta kaust", f"Oled kindel, et soovid kustutada kausta ja failid?\n\n{path_str}")
    if not confirm:
        return
    try:
        shutil.rmtree(path)
    except Exception as e:
        messagebox.showerror("Viga", f"Kausta kustutamine ebaõnnestus:\n{path_str}\n\n{e}")
        return
    messagebox.showinfo("Kustutatud", f"Kaust on kustutatud:\n{path_str}")
    root.hide_delete_button()


def get_log_path():
    # same dir as exe
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "error.log")



def get_field(row, name, default="") -> str:
    if hasattr(row, name):
        val = getattr(row, name)
    else:
        try:
            val = row[name]
        except Exception:
            val = default
    return ("" if val is None else str(val)).strip()