import ttkbootstrap as tb
import tkinter as tk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from pathlib import Path
import threading, os, re
import pythoncom

from utils.logging_helper import log_exception
from src.data_classes import ValidationError, Cancelled
from src.invoice_batch import InvoiceBatch

HUNDRED_PERCENT = 100
REFIT_REGEX = r"(\d+)x(\d+)\+(\d+)\+(\d+)"


def get_data_ready(parent, invoice_var, clients_var, template_root, content_type_var):
    invoice_type = get_selected_invoice_type(parent)
    if invoice_type is None:
        return
    try:
        invoice_path, clients_path = validate_and_prepare_ui(
            parent, invoice_var, clients_var
        )
    except ValidationError as ve:
        messagebox.showerror("Viga", str(ve))
        return
    parent.after(
        10,
        lambda: start_processing_thread(
            worker,
            parent,
            invoice_type,
            invoice_path,
            clients_path,
        ),
    )


def select_file(root, label, btn_text_var, new_text, formats=None):
    if formats:
        invoice_formats = formats
    else:
        invoice_type = get_selected_invoice_type(root)
        if invoice_type is None:
            return
        invoice_formats = _invoice_file_extension(invoice_type.key)

    path = filedialog.askopenfilename(title="Vali fail", filetypes=invoice_formats)
    if path:
        label.set(path)
        btn_text_var.set(new_text)


def cancel_current_job(root):
    root.cancel_event.set()


def center_window(win, min_w=800, min_h=600, max_w=900, max_h=None, margin=40):
    width, height, screen_w, screen_h = get_window_size(
        win, min_w, min_h, max_w, max_h, margin
    )
    x = (screen_w - width) // 2
    y = (screen_h - height) // 2
    win.geometry(f"{width}x{height}+{x}+{y}")


def refit_window(win, min_w=800, min_h=600, max_w=900, max_h=None, margin=40):
    width, height, screen_w, screen_h = get_window_size(
        win, min_w, min_h, max_w, max_h, margin
    )
    m = re.match(REFIT_REGEX, win.geometry())
    x, y = (int(m.group(3)), int(m.group(4))) if m else (100, 100)
    win.geometry(f"{width}x{height}+{x}+{y}")


def get_window_size(win, min_w=800, min_h=600, max_w=900, max_h=None, margin=40):
    win.update_idletasks()
    req_w = win.winfo_reqwidth()
    req_h = win.winfo_reqheight()
    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()
    max_w = min(max_w, screen_w - margin) if max_w else screen_w - margin
    max_h = min(max_h, screen_h - margin) if max_h else screen_h - margin
    width = max(min_w, min(req_w, max_w))
    height = max(min_h, min(req_h, max_h))
    return width, height, screen_w, screen_h


def get_selected_invoice_type(parent):
    key = parent.content_type_var.get()
    return parent.invoice_types.get(key)


def validate_and_prepare_ui(parent, invoice_var, clients_var):
    invoice_path, clients_path = validate_files(invoice_var.get(), clients_var.get())
    parent.status_bar.pack(fill=X, side=BOTTOM)
    refit_window(parent)
    parent.cancel_event.clear()
    parent.btn_cancel.configure(state=NORMAL)
    parent.update_idletasks()
    return invoice_path, clients_path


def validate_files(invoice_path: str, clients_path: str):
    invoice = validate_file_exists(invoice_path, "Arvete fail")
    clients = validate_file_exists(clients_path, "Klientide fail")
    return invoice, clients


def validate_file_exists(path: str, label: str) -> str:
    if not path:
        raise ValidationError(f"{label} on kohustuslik.")
    if not Path(path).is_file():
        raise ValidationError(f"{label} faili ei eksisteeri: {path}")
    return str(path)


def start_processing_thread(target, *args):
    threading.Thread(target=lambda: target(*args), daemon=True).start()


def worker(parent, invoice_type, invoice_path, clients_path):
    try:
        fname = os.path.basename(invoice_path)
        parent.after(0, lambda: parent.status_bar.pack(fill=X, side=BOTTOM))
        parent.after(0, lambda: parent.status_label.configure(text="Alustan..."))
        parent.after(
            0, lambda: parent.page_progress.configure(value=0, mode="determinate")
        )

        batch = InvoiceBatch(
            invoice_path=invoice_path,
            clients_path=clients_path,
            invoice_type=invoice_type,
            cancel_event=parent.cancel_event,
        )
        pythoncom.CoInitialize()
        try:
            batch.load_clients()

            def on_progress(page_number, total_pages, message=None):
                if parent.cancel_event.is_set():
                    raise Cancelled()
                text = message or (
                    f"Loen {page_number}/{total_pages} - {fname}"
                )
                on_task_progress_ui(parent, page_number, total_pages, text)

            batch.load_invoices(on_progress=on_progress)
            if parent.cancel_event.is_set():
                parent.after(0, lambda: on_cancel_ui(parent))
                return

            batch.apply_email_templates()
            batch.prepare_destination()
            batch.save(on_progress=on_progress)
        finally:
            pythoncom.CoUninitialize()

        if parent.cancel_event.is_set():
            parent.after(0, lambda: on_cancel_ui(parent))
            return

        parent.after(0, lambda: finalize_after_saved(parent, batch))
    except Exception as e:
        _handle_worker_error(parent, e)
    finally:

        def cleanup():
            parent.page_progress.configure(value=0, mode="determinate")
            parent.btn_cancel.configure(state=DISABLED)

        parent.after(0, cleanup)


def on_task_progress_ui(parent, page_number, total_pages, message):
    pct = int(page_number / total_pages * HUNDRED_PERCENT) if total_pages else 0

    def apply():
        try:
            parent.status_bar.pack(fill=X, side=BOTTOM)
        except Exception:
            pass
        parent.page_progress.configure(value=pct, mode="determinate")
        parent.status_label.configure(text=message)
        parent.update_idletasks()

    parent.after(0, apply)


def finalize_after_saved(parent, batch: InvoiceBatch):
    try:
        parent.page_progress.configure(value=HUNDRED_PERCENT)
        parent.status_label.configure(text="Valmis")
        parent.on_folder_created(str(batch.dest_dir))
        messagebox.showinfo("Info", f"Arved salvestatakse kausta: {batch.dest_dir}")

        if parent.cancel_event.is_set():
            parent.after(0, lambda: on_cancel_ui(parent))
            return

        problems = batch.match_apartments()
        if problems:
            messagebox.showerror("Viga", " ".join(problems))
        if not batch.persons:
            return

        from gui.email_editor import EmailEditor

        EmailEditor(parent, batch)
        parent.status_bar.pack_forget()
    except Exception as e:
        log_exception(e, operation="finalize_after_saved")
        try:
            messagebox.showerror("Viga", f"Töö ebaõnnestus:\n{e}")
        except Exception as e2:
            log_exception(e2, operation="finalize_error_dialog")


def on_cancel_ui(parent):
    parent.status_label.config(text="Katkestatud")
    parent.page_progress.config(value=0, mode="determinate")
    parent.btn_cancel.configure(state=DISABLED)
    parent.status_bar.pack_forget()


def _invoice_file_extension(key):
    if key == "kommunaal":
        return [("PDF files", "*.pdf")]
    if key == "kyte":
        return [("Excel failid", "*.xls *.xlsx *.xlsm")]
    return [("Kõik failid", "*.*")]


def _handle_worker_error(parent, err):
    if isinstance(err, Cancelled):
        parent.after(0, lambda: on_cancel_ui(parent))
        log_exception(Cancelled("Operation cancelled by user."), operation="cancelled")
    elif isinstance(err, ValidationError):
        parent.after(0, lambda err=err: messagebox.showerror("Viga", str(err)))
        log_exception(err, operation="validation")
    else:
        parent.after(
            0,
            lambda err=err: messagebox.showerror("Viga", f"Töö ebaõnnestus:\n{err}"),
        )
        log_exception(err, operation="worker")
