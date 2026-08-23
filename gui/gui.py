import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import messagebox
import threading
import pytesseract

from src.email_sender import clear_outlook_cache, send_drafts
from utils.logging_helper import log_exc_triple, delete_old_error_log
from utils.file_utils import delete_folder
from utils.config_helpers import (
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
)


class MainWindow(tb.Window):
    def __init__(self, app_name, version, invoice_types, type_hint):
        super().__init__(themename="superhero")
        self.title(app_name)
        self.resizable(True, True)
        self.cancel_event = threading.Event()
        self.invoice_types = invoice_types
        self.type_hint = type_hint
        self.content_type_var = tb.StringVar(value="")
        self.invoice_var = tb.StringVar()
        self.clients_var = tb.StringVar()
        self._setup_exception_handler()
        self._configure_styles()
        self._build_ui(version)
        self._apply_content_type_gate()

    def _setup_exception_handler(self):
        def tk_report_callback_exception(exc_type, exc_value, exc_tb):
            log_exc_triple(exc_type, exc_value, exc_tb, operation="tk_callback")
            try:
                messagebox.showerror(
                    "Viga",
                    f"Kasutajaliidese viga:\n{exc_type.__name__}: {exc_value}",
                )
            except Exception:
                pass

        self.report_callback_exception = tk_report_callback_exception

    def _configure_styles(self):
        style = tb.Style()
        style.configure("TButton", font=("Helvetica", 14))
        style.configure("success.TButton", font=("Helvetica", 14))
        style.configure("TLabel", font=("Helvetica", 14))
        style.configure("info.TLabel", font=("Helvetica", 14))
        style.configure("Path.TLabel", font=("Helvetica", 11))

    def _build_ui(self, version):
        tb.Label(
            self,
            text="Versioon " + version,
            bootstyle=INFO,
            font=("Helvetica", 10),
        ).pack(anchor="ne", padx=10, pady=5)
        container = tb.Frame(self, padding=24)
        container.pack(fill=BOTH, expand=True)
        self._create_invoice_type_section(container)
        self._create_files_section(container)
        self._create_mail_section(container)
        bottom_bar = self._create_bottom_bar()
        self._create_status_bar()
        self._setup_delete_button_handlers(bottom_bar)
        self._setup_send_drafts_button_handlers(bottom_bar)

    def _create_section_header(self, parent, title: str):
        row = tb.Frame(parent)
        row.pack(fill=X, pady=(0, 10))
        tb.Label(
            row, text=title, bootstyle=INFO, font=("Helvetica", 18, "bold")
        ).pack(side=LEFT)
        tb.Separator(row, orient=HORIZONTAL).pack(
            side=LEFT, fill=X, expand=True, padx=(16, 0)
        )

    def _set_type(self, key: str):
        self.content_type_var.set(key)
        self._apply_content_type_gate()

    def _create_invoice_type_section(self, parent):
        self._create_section_header(parent, "Arvete tüüp")
        row = tb.Frame(parent)
        row.pack(fill=X)
        invoice_types = list(self.invoice_types.values())
        type_left, type_right = invoice_types[0], invoice_types[1]
        self.type_left_key = type_left.key
        self.type_right_key = type_right.key
        self.btn_type_left = tb.Button(
            row,
            text=type_left.label,
            bootstyle="info-outline",
            command=lambda: self._set_type(type_left.key),
        )
        self.btn_type_left.pack(
            side=LEFT, fill=X, expand=True, ipady=10, padx=(0, 12)
        )
        self.btn_type_right = tb.Button(
            row,
            text=type_right.label,
            bootstyle="info-outline",
            command=lambda: self._set_type(type_right.key),
        )
        self.btn_type_right.pack(side=LEFT, fill=X, expand=True, ipady=10)
        self.lbl_type_hint = tb.Label(
            parent, text=self.type_hint, bootstyle="secondary"
        )
        self.lbl_type_hint.pack(anchor=W, pady=(10, 0))

    def _create_files_section(self, parent):
        section = tb.Frame(parent)
        section.pack(fill=X, pady=(0, 18))
        self._create_section_header(section, "Failid")
        grid = tb.Frame(section)
        grid.pack(fill=X)
        grid.grid_columnconfigure(0, weight=0)
        grid.grid_columnconfigure(1, weight=1)

        self.btn_text_invoice = tk.StringVar(value="Vali arvete fail")
        self.btn_invoice = tb.Button(
            grid,
            textvariable=self.btn_text_invoice,
            bootstyle=INFO,
            command=lambda: select_file(
                self, self.invoice_var, self.btn_text_invoice, "Muuda arvete faili"
            ),
        )
        self.btn_invoice.grid(
            row=0, column=0, sticky="w", padx=(0, 12), pady=(0, 10), ipady=8
        )
        tb.Label(
            grid,
            textvariable=self.invoice_var,
            style="Path.TLabel",
            foreground="#9aa0a6",
            wraplength=650,
            anchor="w",
            justify="left",
        ).grid(row=0, column=1, sticky="w", pady=(0, 10))

        self.btn_text_clients = tk.StringVar(value="Vali klientide fail")
        self.btn_clients = tb.Button(
            grid,
            textvariable=self.btn_text_clients,
            bootstyle=INFO,
            command=lambda: select_file(
                self,
                self.clients_var,
                self.btn_text_clients,
                "Muuda klientide faili",
                formats=[("Excel failid", "*.xls *.xlsx *.xlsm")],
            ),
        )
        self.btn_clients.grid(
            row=1, column=0, sticky="w", padx=(0, 12), ipady=8
        )
        tb.Label(
            grid,
            textvariable=self.clients_var,
            style="Path.TLabel",
            foreground="#9aa0a6",
            wraplength=650,
            anchor="w",
            justify="left",
        ).grid(row=1, column=1, sticky="w")

    def _create_mail_section(self, parent):
        section = tb.Frame(parent)
        section.pack(fill=X, pady=(0, 8))
        self._create_section_header(section, "Meil")
        self.btn_compose = tb.Button(
            section,
            text="Koosta meilid",
            bootstyle="success",
            command=lambda: get_data_ready(
                self, self.invoice_var, self.clients_var, self, self.content_type_var
            ),
        )
        self.btn_compose.pack(anchor=W, ipady=10, ipadx=18)
        tb.Label(
            section,
            text="Muuda teemat ja sisu enne saatmist",
            bootstyle="secondary",
            font=("Helvetica", 12),
        ).pack(anchor=W, pady=(10, 0))

    def _create_bottom_bar(self):
        bottom_bar = tb.Frame(self)
        bottom_bar.pack(fill=X, side=BOTTOM)
        self.btn_cancel = tb.Button(
            bottom_bar,
            text="Katkesta",
            bootstyle="danger",
            command=lambda: cancel_current_job(self),
        )
        self.btn_cancel.pack(side=LEFT, padx=(12, 6), pady=12)
        self.btn_cancel.configure(state=DISABLED)
        return bottom_bar

    def _create_status_bar(self):
        status_bar = tb.Frame(self)
        self.status_label = tb.Label(
            status_bar, text="Valmis", bootstyle=INFO, anchor="w"
        )
        self.status_label.pack(side=LEFT, padx=10, pady=8)
        self.page_progress = tb.Progressbar(
            status_bar,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            bootstyle=INFO,
        )
        self.page_progress.pack(
            side=LEFT, fill=X, expand=True, padx=(10, 12), pady=8
        )
        self.status_bar = status_bar

        def enforce_layout(_event=None):
            window_width = self.winfo_width()
            if window_width <= 1:
                return
            self.page_progress.configure(length=int(window_width * 0.5))
            self.status_label.configure(wraplength=int(window_width * 0.45))

        self.bind("<Configure>", enforce_layout)

    def _setup_delete_button_handlers(self, parent):
        self.invoices_dir_var = tb.StringVar(value="")
        self.btn_delete_invoices = tb.Button(
            parent,
            text="Kustuta arvekaust",
            bootstyle=DANGER,
            command=lambda: delete_folder(self, self.invoices_dir_var.get()),
        )
        self._delete_packed = False

        def on_folder_created(path: str):
            self.invoices_dir_var.set(path)
            if not self._delete_packed:
                self.btn_delete_invoices.pack(side=RIGHT, padx=(0, 12), pady=12)
                self._delete_packed = True

        def hide_delete_button():
            if self._delete_packed:
                try:
                    self.btn_delete_invoices.pack_forget()
                except Exception:
                    pass
                self._delete_packed = False
            self.invoices_dir_var.set("")

        self.on_folder_created = on_folder_created
        self.hide_delete_button = hide_delete_button

    def _setup_send_drafts_button_handlers(self, parent):
        self.btn_send_drafts = tb.Button(
            parent,
            text="Saada mustandid",
            bootstyle=SUCCESS,
            command=lambda: send_drafts(self),
        )
        self._send_drafts_packed = False

        def on_emails_saved():
            if not self._send_drafts_packed:
                self.btn_send_drafts.pack(side=RIGHT, padx=(0, 12), pady=12)
                self._send_drafts_packed = True
                try:
                    center_window(self)
                except Exception:
                    pass

        def hide_send_drafts_button():
            if self._send_drafts_packed:
                try:
                    self.btn_send_drafts.pack_forget()
                except Exception:
                    pass
                self._send_drafts_packed = False

        self.on_emails_saved = on_emails_saved
        self.hide_send_drafts_button = hide_send_drafts_button

    def _apply_content_type_gate(self):
        key = self.content_type_var.get()
        enabled = key in self.invoice_types
        self.lbl_type_hint.configure(text="" if enabled else self.type_hint)
        if key == getattr(self, "type_left_key", None):
            self.btn_type_left.configure(bootstyle=INFO)
            self.btn_type_right.configure(bootstyle="info-outline")
        elif key == getattr(self, "type_right_key", None):
            self.btn_type_left.configure(bootstyle="info-outline")
            self.btn_type_right.configure(bootstyle=INFO)
        else:
            self.btn_type_left.configure(bootstyle="info-outline")
            self.btn_type_right.configure(bootstyle="info-outline")
        state = NORMAL if enabled else DISABLED
        for widget in (self.btn_invoice, self.btn_clients, self.btn_compose):
            widget.configure(state=state)
        for name in ("btn_delete_invoices", "btn_send_drafts"):
            if hasattr(self, name):
                try:
                    getattr(self, name).configure(state=state)
                except Exception:
                    pass


def main():
    clear_outlook_cache()
    config = read_config()
    if not _perform_startup_checks():
        return
    version = load_app_version(config)
    app_name = load_app_name(config)
    invoice_types, type_hint = load_invoice_types(config)
    window = MainWindow(app_name, version, invoice_types, type_hint)
    center_window(window, min_w=800, min_h=650, max_w=980)
    window.deiconify()
    window.lift()
    window.focus_force()
    window.mainloop()


def _perform_startup_checks() -> bool:
    delete_old_error_log()
    pytesseract.pytesseract.tesseract_cmd = get_tesseract_cmd()
    if not check_ocr_environment():
        messagebox.showerror(
            "Tesseract puudub",
            "Tesseract OCR ei ole selles arvutis paigaldatud või ei leitud teekonda.\n\n"
            "Palun paigalda Tesseract OCR ja proovi uuesti.",
        )
        return False
    return True
