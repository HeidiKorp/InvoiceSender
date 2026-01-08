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
from utils.file_utils import (
    delete_folder,
    read_config,
    load_invoice_types,
    load_app_version,
    load_app_name,
)
from utils.ocr_helper import get_tesseract_cmd, check_ocr_environment
from utils.gui_helpers import (
    select_file,
    center_window,
    cancel_current_job,
    get_data_ready,
    get_selected_invoice_type,
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
    style.configure("TButton", font=("Helvetica", 14))
    style.configure("success.TButton", font=("Helvetica", 14))
    style.configure("TLabel", font=("Helvetica", 14))
    style.configure("info.TLabel", font=("Helvetica", 14))

    style.configure("Path.TLabel", font=("Helvetica", 11))


def _setup_window_properties(root, app_name):
    """
    Set up window title, resizability, and shared state on the root window.
    Args:
        root: The main application window (ttkbootstrap Window).
        config: The configuration object.
    """
    # --- Window properties ---
    root.title(app_name)
    root.resizable(True, True)

    # --- Shared state ---
    root.cancel_event = threading.Event()
    root.current_worker = None


def _create_version_label(root, version):
    """
    Create the version label at the top-right corner of the window.
    Returns the created label.
    """
    version_label = tb.Label(
        root,
        text="Versioon " + version,
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

    root.status_label = tb.Label(status_bar, text="Valmis", bootstyle=INFO, anchor="w")
    root.status_label.pack(side=LEFT, padx=10, pady=8)

    root.page_progress = tb.Progressbar(
        status_bar, orient="horizontal", mode="determinate", maximum=100, bootstyle=INFO
    )

    root.page_progress.pack(side=LEFT, fill=X, expand=True, padx=(10, 12), pady=8)

    def enforce_layout(_event=None):
        window_width = root.winfo_width()
        if window_width <= 1:
            return
        
        min_width = int(window_width * 0.5)
        root.page_progress.configure(length=min_width)

        root.status_label.configure(wraplength=int(root.winfo_width() * 0.45))

    root.bind("<Configure>", enforce_layout)

    # Set once after initial layout
    root.after(0, enforce_layout)

    # Hide status bar initially
    status_bar.pack_forget()
    root.status_bar = status_bar  # store reference on root so we can show/hide later


def _create_bottom_bar(root):
    """
    Create the bottom bar with "Koosta meilid" and "Katkesta" buttons.
    Returns the created bottom bar frame.
    """
    bottom_bar = tb.Frame(root)
    bottom_bar.pack(fill=X, side=BOTTOM)

    # --- Cancel button ---
    root.btn_cancel = tb.Button(
        bottom_bar,
        text="Katkesta",
        bootstyle="danger",
        command=lambda: cancel_current_job(root),
    )
    root.btn_cancel.pack(side=LEFT, padx=(12, 6), pady=12)
    root.btn_cancel.configure(state=DISABLED)
    return bottom_bar


def _set_invoice_type_style(root, left_style, right_style):
    root.btn_type_left.configure(bootstyle=left_style)
    root.btn_type_right.configure(bootstyle=right_style)


def _apply_content_type_gate(root):
    key = root.content_type_var.get()
    enabled = key in root.invoice_types

    # Hint text only when nothing selected
    root.lbl_type_hint.configure(text="" if enabled else root.type_hint)

    # Styling: selected card is filled info, other is outline
    if key == getattr(root, "type_left_key", None):
        _set_invoice_type_style(root, INFO, "info-outline")
    elif key == getattr(root, "type_right_key", None):
        _set_invoice_type_style(root, "info-outline", INFO)
    else:
        _set_invoice_type_style(root, "info-outline", "info-outline")

    # Disable/enable other controls (but keep them visible)
    state = NORMAL if enabled else DISABLED

    for widget in (
        root.btn_invoice,
        root.btn_clients,
        root.btn_compose,
    ):
        widget.configure(state=state)

    for name in ("btn_delete_invoices", "btn_send_drafts"):
        if hasattr(root, name):
            try:
                getattr(root, name).configure(state=state)
            except Exception:
                pass


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


def _create_section_header(parent, title: str):
    # Header row: title + separation line
    row = tb.Frame(parent)
    row.pack(fill=X, pady=(0, 10))
    tb.Label(row, text=title, bootstyle=INFO, font=("Helvetica", 18, "bold")).pack(
        side=LEFT
    )
    tb.Separator(row, orient=HORIZONTAL).pack(
        side=LEFT, fill=X, expand=True, padx=(16, 0)
    )


def _create_file_buttons(root, parent, invoice_var, clients_var):
    # Invoice row
    root.btn_text_invoice = tk.StringVar(value="Vali arvete fail")

    root.btn_invoice = tb.Button(
        parent,
        textvariable=root.btn_text_invoice,
        bootstyle=INFO,
        command=lambda: select_file(
            root,
            invoice_var,
            root.btn_text_invoice,
            "Muuda arvete faili",
        ),
    )
    root.btn_invoice.grid(
        row=0, column=0, sticky="w", padx=(0, 12), pady=(0, 10), ipady=8
    )

    root.lbl_invoice = tb.Label(
        parent,
        textvariable=invoice_var,
        style="Path.TLabel",
        foreground="#9aa0a6",
        wraplength=650,
        anchor="w",
        justify="left",
    )
    root.lbl_invoice.grid(row=0, column=1, sticky="w", pady=(0, 10))

    # Clients row
    root.btn_text_clients = tk.StringVar(value="Vali klientide fail")
    root.btn_clients = tb.Button(
        parent,
        textvariable=root.btn_text_clients,
        bootstyle=INFO,
        command=lambda: select_file(
            root,
            clients_var,
            root.btn_text_clients,
            "Muuda klientide faili",
            formats=[("Excel failid", "*.xls *.xlsx")],
        ),
    )
    root.btn_clients.grid(
        row=1, column=0, sticky="w", padx=(0, 12), pady=(0, 0), ipady=8
    )
    root.lbl_clients = tb.Label(
        parent,
        textvariable=clients_var,
        style="Path.TLabel",
        foreground="#9aa0a6",
        wraplength=650,
        anchor="w",
        justify="left",
    )
    root.lbl_clients.grid(row=1, column=1, sticky="w")


def _create_files_section(root, parent, invoice_var, clients_var):
    section = tb.Frame(parent)
    section.pack(fill=X, pady=(0, 18))

    _create_section_header(section, "Failid")

    grid = tb.Frame(section)
    grid.pack(fill=X)

    grid.grid_columnconfigure(0, weight=0)  # buttons
    grid.grid_columnconfigure(1, weight=1)  # paths

    _create_file_buttons(root, grid, invoice_var, clients_var)


def _set_type(root, key: str):
    root.content_type_var.set(key)
    _apply_content_type_gate(root)


def _create_invoice_type_section(root, parent):
    _create_section_header(parent, "Arvete tüüp")

    row = tb.Frame(parent)
    row.pack(fill=X)

    invoice_types = list(root.invoice_types.values())
    type_left, type_right = invoice_types[0], invoice_types[1]

    # Store keys so gate styling can compare by key
    root.type_left_key = type_left.key
    root.type_right_key = type_right.key

    # "Radio-card" style buttons
    root.btn_type_left = tb.Button(
        row,
        text=type_left.label,
        bootstyle="info-outline",
        command=lambda: _set_type(root, type_left.key),
    )
    root.btn_type_left.pack(side=LEFT, fill=X, expand=True, ipady=10, padx=(0, 12))

    root.btn_type_right = tb.Button(
        row,
        text=type_right.label,
        bootstyle="info-outline",
        command=lambda: _set_type(root, type_right.key),
    )
    root.btn_type_right.pack(side=LEFT, fill=X, expand=True, ipady=10)

    # Helper text
    root.lbl_type_hint = tb.Label(parent, text=root.type_hint, bootstyle="secondary")
    root.lbl_type_hint.pack(anchor=W, pady=(10, 0))

    # root.content_type_var = content_type_var


def _create_mail_section(root, parent, invoice_var, clients_var):
    section = tb.Frame(parent)
    section.pack(fill=X, pady=(0, 8))

    _create_section_header(section, "Meil")

    root.btn_compose = tb.Button(
        section,
        text="Koosta meilid",
        bootstyle="success",
        command=lambda: get_data_ready(
            root, invoice_var, clients_var, root, root.content_type_var
        ),
    )
    root.btn_compose.pack(anchor=W, ipady=10, ipadx=18)

    tb.Label(
        section,
        text="Muuda teemat ja sisu enne saatmist",
        bootstyle="secondary",
        font=("Helvetica", 12),
    ).pack(anchor=W, pady=(10, 0))


def _setup_ui_components(root, version, invoice_var, clients_var, content_type_var):
    """Set up all UI components."""
    # --- Version label (top right) ---
    _create_version_label(root, version)

    # --- One main container for consistent padding ---
    container = tb.Frame(root, padding=24)
    container.pack(fill=BOTH, expand=True)

    # --- Invoice type selector ---
    _create_invoice_type_section(root, container)

    # --- Middle: files + mail sections
    _create_files_section(root, container, invoice_var, clients_var)

    _create_mail_section(root, container, invoice_var, clients_var)

    # --- Bottom bar ---
    bottom_bar = _create_bottom_bar(root)

    # --- Status bar (bottom, above the button row) ---
    _create_status_bar(root)

    # -- Delete button handlers ---
    _setup_delete_button_handlers(root, bottom_bar)

    # --- Send drafts button handlers ---
    _setup_send_drafts_button_handlers(root, bottom_bar)

    # --- Apply initial disabled state until type chosen
    _apply_content_type_gate(root)


def main():
    # --- Clean Outlook cache on startup ---
    clear_outlook_cache()
    config = read_config()

    if not _perform_startup_checks():
        return

    # --- Start window setup ---
    root = tb.Window(themename="superhero")

    # --- Set up exception handler for main thread ---
    _setup_exception_handler(root)

    # --- Configure styles ---
    style = tb.Style()
    _configure_styles(style)

    # --- Set up window properties ---
    version = load_app_version(config)
    app_name = load_app_name(config)
    _setup_window_properties(root, app_name)

    invoice_var = tb.StringVar()
    clients_var = tb.StringVar()
    root.invoice_types, root.type_hint = load_invoice_types(config)
    root.content_type_var = tb.StringVar(value="")  # "", "kommunaal", "kyte"

    # --- Create UI components ---
    _setup_ui_components(root, version, invoice_var, clients_var, root.content_type_var)

    center_window(root, min_w=800, min_h=650, max_w=980)
    root.deiconify()
    root.lift()
    root.focus_force()

    root.mainloop()
