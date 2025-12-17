import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
import threading, os
import pytesseract
from ttkbootstrap.style import Bootstyle

from src.email_sender import clear_outlook_cache, send_drafts
from utils.logging_helper import (
    log_exc_triple,
    delete_old_error_log,
    _thread_excepthook,
)
from utils.file_utils import delete_folder, read_config
from utils.ocr_helper import get_tesseract_cmd, check_ocr_environment
from utils.gui_helpers import (
    select_file,
    center_window,
    cancel_current_job,
    get_data_ready,
)


threading.excepthook = _thread_excepthook


def _perform_startup_checks() -> bool:
    """
    Perform startup checks: delete old error log, set up pytesseract, check Tesseract environment.
    Returns True if all checks pass, False otherwise.
    """
    # Remove old error log if present
    delete_old_error_log()

    # Ensure pytesseract is set up correctly
    pytesseract.pytesseract.tesseract_cmd = get_tesseract_cmd()

    # Check Tesseract environment - if not OK, exit
    if not check_ocr_environment():
        messagebox.showerror(
            "Tesseract puudub",
            "Tesseract OCR ei ole selles arvutis paigaldatud või ei leitud teekonda.\n\n"
            "Palun paigalda Tesseract OCR ja proovi uuesti.",
        )
        return False
    return True


def _setup_exception_handler(root):
    """
    Set up a global exception handler for the TK root to log uncaught exceptions.
    Args:
        root: The main application window (ttkbootstrap Window).
    """
    # Attach a global exception handler to the TK root so uncaught exceptions are logged
    def tk_report_callback_exception(exc_type, exc_value, exc_tb):
        # Log it
        log_exc_triple(exc_type, exc_value, exc_tb)
        # Also show a more user-friendly error
        try:
            messagebox.showerror(
                "Viga", f"Kasutajaliidese viga:\n{exc_type.__name__}: {exc_value}"
            )
        except Exception:
            # UI might be in an inconsistent state; don't raise again
            pass

    # All uncaught exceptions to be logged
    root.report_callback_exception = tk_report_callback_exception


def _configure_styles(style):
    """
    Configure global styles for buttons and labels.
    Args:
        style: The ttkbootstrap Style object.
    """
    # Centralize style configuration
    style.configure("TButton", font=("Helvetica", 15))
    style.configure("success.TButton", font=("Helvetica", 15))
    style.configure("TLabel", font=("Helvetica", 15))
    style.configure("info.TLabel", font=("Helvetica", 15))


def _setup_window_properties(root, config):
    """
    Set up window title, resizability, and shared state on the root window.
    Args:
        root: The main application window (ttkbootstrap Window).
        config: The configuration object.
    """
    # --- Window properties ---
    root.title(config.get("app", "NAME", fallback="Arvete Saatja"))
    root.resizable(True, True)

    # --- Shared state ---
    root.cancel_event = threading.Event()
    root.current_worker = None


def _create_version_label(root, config):
    """
    Create the version label at the top-right corner of the window.
    Returns the created label.
    """
    version_label = tb.Label(
        root,
        text="Versioon " + config.get("app", "VERSION", fallback="1.0.0"),
        bootstyle=INFO,
        font=("Helvetica", 10),
    )
    version_label.pack(anchor="ne", padx=10, pady=5)
    return version_label


def _create_status_bar(root):
    """
    Create the status bar at the bottom of the window with a label and progress bar.
    Returns the created status bar frame.
    """
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


def _create_bottom_bar(root, invoice_var, clients_var, subject, body):
    """
    Create the bottom bar with "Koosta meilid" and "Katkesta" buttons.
    Returns the created bottom bar frame.
    """
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

    # --- Cancel button ---
    root.btn_cancel = tb.Button(
        bottom_bar,
        text="Katkesta",
        bootstyle="danger",
        command=lambda: cancel_current_job(root),
    )
    root.btn_cancel.pack(side=RIGHT, padx=0, pady=12)
    root.btn_cancel.configure(state=DISABLED)
    return bottom_bar


def _setup_delete_button_handlers(root, parent):
    """
    Create the "Kustuta arvekaust" button attached to 'parent' and set up on_folder_created and hide_delete_button handlers on 'root'.
    'root.invoices_dir_var' is used to store the current invoices directory path.
    The button is initially hidden and only shown when a folder is created.
    """
    if not hasattr(root, "invoices_dir_var"):
        root.invoices_dir_var = tb.StringVar(value="")

    btn_delete_invoices = tb.Button(
        parent,
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
            try:
                root.btn_delete_invoices.pack_forget()
            except Exception:
                pass
            root._delete_packed = False
        root.invoices_dir_var.set("")

    root.on_folder_created = on_folder_created
    root.hide_delete_button = hide_delete_button


def _setup_send_drafts_button_handlers(root, parent):
    """
    Create the "Saada mustandid" button attached to 'parent' and 
    set up on_emails_saved / hide_send_drafts_button handlers on 'root'.
    The button is initially hidden and only shown when emails are saved.
    """
    btn_send_drafts = tb.Button(
        parent,
        text="Saada mustandid",
        bootstyle=SUCCESS,
        command=lambda: send_drafts(root),
    )
    root.btn_send_drafts = btn_send_drafts

    def on_emails_saved():
        if not getattr(root, "_send_drafts_packed", False):
            root.btn_send_drafts.pack(side=RIGHT, padx=(0, 12), pady=12)
            root._send_drafts_packed = True
            try:
                center_window(root)  # Re-center after adding button
            except Exception:
                pass

    def hide_send_drafts_button():
        if getattr(root, "_send_drafts_packed", False):
            try:
                root.btn_send_drafts.pack_forget()
            except Exception:
                pass
            root._send_drafts_packed = False

    root.on_emails_saved = on_emails_saved
    root.hide_send_drafts_button = hide_send_drafts_button

    return btn_send_drafts

def _create_content_area(root, invoice_var, clients_var):
    """
    Create the main content area with file selection buttons and labels.
    Returns the created content frame.
    """
    content = tb.Frame(root)
    content.pack(expand=True)

    # Invoice
    btn_text_invoice = tk.StringVar(value="Vali arvete fail")
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
    btn_text_clients = tk.StringVar(value="Vali klientide fail")
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
    return content


def main():
    # --- Clean Outlook cache on startup ---
    clear_outlook_cache()
    config = read_config()

    if not _perform_startup_checks():
        return

    # --- Set up default subject, body ---
    subject = "Arve"
    body = (
        "Lugupeetud KÜ korteri omanik. Kü edastab järjekordse korteri "
        "kuu kulude arve. See on automaatteavitus, palume mitte vastata."
    )

    # --- Start window setup ---
    root = tb.Window(themename="superhero")

    # --- Set up exception handler for main thread ---
    _setup_exception_handler(root)

    # --- Configure styles ---
    style = tb.Style()
    _configure_styles(style)

    # --- Set up window properties ---
    _setup_window_properties(root, config)

    invoice_var = tb.StringVar()
    clients_var = tb.StringVar()

    # --- Version label (top right) ---
    _create_version_label(root, config)

    # --- Bottom bar ---
    bottom_bar = _create_bottom_bar(root, invoice_var, clients_var, subject, body)

    # --- Status bar (bottom, above the button row) ---
    _create_status_bar(root)

    # -- Delete button handlers ---
    btn_delete_invoices = _setup_delete_button_handlers(root, bottom_bar)

    # --- Send drafts button handlers ---
    btn_send_drafts = _setup_send_drafts_button_handlers(root, bottom_bar)

    # --- Content area ---
    content = _create_content_area(root, invoice_var, clients_var)

    center_window(root, min_w=500, min_h=500, max_w=700)
    # root.update_idletasks()
    root.deiconify()
    root.lift()
    root.focus_force()

    root.mainloop()
