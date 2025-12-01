import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
# from pdf_extractor import ocr_pdf_all_pages, log_line
import sys, threading, os
import shutil
import traceback
import pytesseract

from ttkbootstrap.style import Bootstyle
from src.xls_extractor import extract_person_data, ValidationError
# from pdf_extractor import separate_invoices, save_each_invoice_as_file
from src.email_sender import clear_outlook_cache, send_drafts
from utils.logging_helper import (
    log_exc_triple,
    log_exception,
    delete_old_error_log,
    _thread_excepthook,
)
from utils.file_utils import delete_folder
from utils.ocr_helper import get_tesseract_cmd, check_ocr_environment
from utils.gui_helpers import (
    select_file,
    center_window,
    validate_files,
    get_data_ready,
    open_outlook,
    open_email_editor,
    call_error,
    _Cancelled,
    cancel_current_job,
)


threading.excepthook = _thread_excepthook


def main():
    # Clean Outlook cache on startup
    clear_outlook_cache()

    # Set up default subject, body
    subject = "Arve"
    body = (
        "Lugupeetud KÜ korteri omanik. Kü edastab järjekordse korteri "
        "kuu kulude arve. See on automaatteavitus, palume mitte vastata."
    )

    # # --- Window setup ---
    root = tb.Window(themename="superhero")

    delete_old_error_log()

    pytesseract.pytesseract.tesseract_cmd = get_tesseract_cmd()

    print(f"Using tesseract version={pytesseract.get_tesseract_version()}")

    if not check_ocr_environment():
        messagebox.showerror(
            "Tesseract puudub",
            "Tesseract OCR ei ole selles arvutis paigaldatud või ei leitud teekonda.\n\n"
            "Palun paigalda Tesseract OCR ja proovi uuesti.",
        )
        root.destroy()
        return

    def tk_report_callback_exception(exc_type, exc_value, exc_tb):
        # Log it
        log_exc_triple(exc_type, exc_value, exc_tb)
        # Also show a more user-friendly error
        messagebox.showerror(
            "Viga", f"Kasutajaliidese viga:\n{exc_type.__name__}: {exc_value}"
        )

    root.report_callback_exception = tk_report_callback_exception

    root.title("Invoice Sender")
    root.resizable(True, True)
    center_window(root, 1100, 800)
    root.update_idletasks()
    root.deiconify()
    root.lift()
    root.focus_force()

    invoice_var = tb.StringVar()
    clients_var = tb.StringVar()
    invoices_dir_var = tb.StringVar()

    style = tb.Style()

    # Define a custom font and size
    style.configure("TButton", font=("Helvetica", 15))
    style.configure("success.TButton", font=("Helvetica", 15))
    style.configure("TLabel", font=("Helvetica", 15))
    style.configure("info.TLabel", font=("Helvetica", 15))

    # shared state
    root.cancel_event = threading.Event()
    root.current_worker = None

    # --- Bottom bar with Next (right corner) ---
    bottom_bar = tb.Frame(root)
    bottom_bar.pack(fill=X, side=BOTTOM)
    tb.Button(
        bottom_bar,
        text="Koosta meilid",
        bootstyle="success",
        command=lambda: get_data_ready(
            root, invoice_var, clients_var, root, subject, body
        ),
    ).pack(side=RIGHT, padx=12, pady=12)

    # Cancel button
    root.btn_cancel = tb.Button(
        bottom_bar,
        text="Katkesta",
        bootstyle="danger",
        command=lambda: cancel_current_job(root),
    )
    root.btn_cancel.pack(side=RIGHT, padx=0, pady=12)
    root.btn_cancel.configure(state=DISABLED)

    # --- Status bar (bottom, above the button row) ---
    status_bar = tb.Frame(root)
    status_bar.pack(fill=X, side=BOTTOM)

    root.status_label = tb.Label(status_bar, text="Valmis", bootstyle=INFO)
    root.status_label.pack(side=LEFT, padx=10, pady=8)

    root.page_progress = tb.Progressbar(
        status_bar, orient="horizontal", mode="determinate", maximum=100, bootstyle=INFO
    )

    root.page_progress.pack(side=LEFT, fill=X, expand=True, padx=(10, 12), pady=8)

    # Hide status bar initially
    status_bar.pack_forget()
    root.status_bar = status_bar  # store reference on root so we can show/hide later

    root.invoices_dir_var = tb.StringVar(value="")
    btn_delete_invoices = tb.Button(
        bottom_bar,
        text="Kustuta arvekaust",
        bootstyle=DANGER,
        command=lambda: delete_folder(root, root.invoices_dir_var.get()),
    )

    # Keep state/handles somewhere accessible (closure or attributes)
    root.btn_delete_invoices = btn_delete_invoices

    def on_folder_created(path: str):
        root.invoices_dir_var.set(path)
        if not getattr(root, "_delete_packed", False):
            root.btn_delete_invoices.pack(side=RIGHT, padx=(0, 12), pady=12)
            root._delete_packed = True

    def hide_delete_button():
        if getattr(root, "_delete_packed", False):
            root.btn_delete_invoices.pack_forget()
            root._delete_packed = False
        root.invoices_dir_var.set("")

    root.on_folder_created = on_folder_created
    root.hide_delete_button = hide_delete_button

    # --- Send drafts button
    btn_send_drafts = tb.Button(
        bottom_bar,
        text="Saada mustandid",
        bootstyle=SUCCESS,
        command=lambda: send_drafts(root),
    )
    root.btn_send_drafts = btn_send_drafts

    def on_emails_saved():
        if not getattr(root, "_send_drafts_packed", False):
            root.btn_send_drafts.pack(side=RIGHT, padx=(0, 12), pady=12)
            root._send_drafts_packed = True

    def hide_send_drafts_button():
        if getattr(root, "_send_drafts_packed", False):
            root.btn_send_drafts.pack_forget()
            root._send_drafts_packed = False

    root.on_emails_saved = on_emails_saved
    root.hide_send_drafts_button = hide_send_drafts_button

    # --- Center content ---
    content = tb.Frame(root)
    content.pack(expand=True)

    # Invoice
    btn_text_invoice = tk.StringVar(value="Vali arvete fail")
    btn_text_clients = tk.StringVar(value="Vali klientide fail")

    btn_invoice = tb.Button(
        content,
        textvariable=btn_text_invoice,
        bootstyle=INFO,
        command=lambda: select_file(
            invoice_var,
            [("PDF files", "*.pdf")],
            btn_text_invoice,
            "Muuda arvete faili",
        ),
    )
    btn_invoice.grid(row=0, column=0, padx=22, pady=22)
    lbl_invoice = tb.Label(
        content, textvariable=invoice_var, wraplength=680, foreground="#9aa0a6"
    )
    lbl_invoice.grid(row=1, column=0, padx=12, pady=12)

    # Clients
    btn_clients = tb.Button(
        content,
        textvariable=btn_text_clients,
        bootstyle=INFO,
        command=lambda: select_file(
            clients_var,
            [("XLS files", "*.xls"), ("XLSX files", "*.xlsx")],
            btn_text_clients,
            "Muuda klientide faili",
        ),
    )
    btn_clients.grid(row=2, column=0, padx=12, pady=12)
    lbl_clients = tb.Label(
        content, textvariable=clients_var, wraplength=680, foreground="#9aa0a6"
    )
    lbl_clients.grid(row=3, column=0, padx=12, pady=12)

    # Center the column
    content.grid_columnconfigure(0, weight=1)

    root.mainloop()
