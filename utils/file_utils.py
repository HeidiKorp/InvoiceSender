from pathlib import Path
from tkinter import messagebox
import shutil, os, sys

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

