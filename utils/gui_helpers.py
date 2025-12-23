import ttkbootstrap as tb
import tkinter as tk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from pathlib import Path
import threading, os, re
import pytesseract

from utils.logging_helper import log_exception
from src.pdf_extractor import separate_invoices, save_each_invoice_as_file
from src.xls_extractor import extract_person_data, ValidationError
from src.email_sender import (
    save_emails_with_invoices,
    ensure_outlook_ready,
    validate_persons_vs_invoices,
)

HUNDRED_PERCENT = 100
REFIT_REGEX = r"(\d+)x(\d+)\+(\d+)\+(\d+)"


def select_file(label, filetypes, btn_text_var, new_text):
    """Open file dialog and set label and button text."""
    path = filedialog.askopenfilename(title="Vali fail", filetypes=filetypes)
    if path:
        label.set(path)
        btn_text_var.set(new_text)


def get_window_size(win, min_w=800, min_h=600, max_w=900, max_h=None, margin=40):
    """Get window size clamped to min/max/screen size."""
    win.update_idletasks()

    req_w = win.winfo_reqwidth()
    req_h = win.winfo_reqheight()

    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()

    # Clamp to min/max/screen size
    max_w = min(max_w, screen_w - margin) if max_w else screen_w - margin
    max_h = min(max_h, screen_h - margin) if max_h else screen_h - margin

    width = max(min_w, min(req_w, max_w))
    height = max(min_h, min(req_h, max_h))

    return width, height, screen_w, screen_h


def center_window(win, min_w=800, min_h=600, max_w=900, max_h=None, margin=40):
    """Center a window on the screen with given width/height."""
    width, height, screen_w, screen_h = get_window_size(
        win, min_w, min_h, max_w, max_h, margin
    )

    x = (screen_w - width) // 2
    y = (screen_h - height) // 2

    win.geometry(f"{width}x{height}+{x}+{y}")


def refit_window(win, min_w=800, min_h=600, max_w=900, max_h=None, margin=40):
    """Refit window to new size but keep current position."""
    width, height, screen_w, screen_h = get_window_size(
        win, min_w, min_h, max_w, max_h, margin
    )

    # keep current position, do not recenter
    m = re.match(REFIT_REGEX, win.geometry())
    x, y = (int(m.group(3)), int(m.group(4))) if m else (100, 100)

    win.geometry(f"{width}x{height}+{x}+{y}")


def _on_progress_ui(parent, page_number, total_pages, fname):
    """Update progress bar and status label in the GUI thread."""
    pct = int(page_number / total_pages * HUNDRED_PERCENT) if total_pages else 0
    parent.after(
        0,
        lambda: (
            parent.page_progress.config(value=pct),
            parent.status_label.configure(
                text=f"PDF leht {page_number}/{total_pages} - {fname}"
            ),
        ),
    )


def validate_file_exists(path: str, label: str) -> str:
    """Validate that a file exists at the given path."""
    if not path:
        raise ValidationError(f"{label} on kohustuslik.")
    if not Path(path).is_file():
        raise ValidationError(f"{label} faili ei eksisteeri: {path}")
    return str(path)


def validate_files(invoice_path: str, clients_path: str):
    """Validate that both invoice and clients files exist."""
    invoice = validate_file_exists(invoice_path, "Arvete fail")
    clients = validate_file_exists(clients_path, "Klientide fail")
    return invoice, clients


def call_error(text):
    """Show an error message box."""
    messagebox.showerror("Viga", text)


def _create_dest_directory(invoice_path: str):
    """Create a destination directory for processed invoices."""
    parent = Path(invoice_path).resolve().parent
    dest = parent / "arved"

    # Try to create the directory (with parent, ignore if already exists)
    try:
        dest.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        log_exception(e)
        raise ValidationError(f"Kausta loomine ebaõnnestus:\n{dest}\n\n{e}")
    if not dest.exists() or not dest.is_dir():
        raise ValidationError(f"Kausta ei õnnestunud luua:\n{dest}")
    return dest


def _process_ocr(invoice_path: str, cancel_flag, on_progress):
    """Process the invoice PDF with OCR and return extracted invoices."""
    try:
        invoices = separate_invoices(
            invoice_path,
            on_progress=on_progress,
            cancel_flag=cancel_flag,
        )
    except pytesseract.TesseractError as e:
        log_exception(e)
        raise ValidationError(
            f"OCR töötlemine ebaõnnestus. Kontrolli, kas Tesseract on õigesti paigaldatud.\n{e}"
        )
    return invoices


def _save_and_open_invoices(
    parent, invoices, dest: Path, template_root, persons, subject, body
):
    """Save invoices to files and open email editor."""
    # After OK -> open the email editor
    invoices_dir = save_each_invoice_as_file(
        invoices, dest
    )  # returns the full folder path to all individual invoices
    parent.on_folder_created(str(invoices_dir))
    open_email_editor(template_root, persons, invoices_dir, subject, body)


def validate_and_prepare_ui(parent, invoice_var: str, clients_var: str):
    """Validate input files and prepare the UI for processing."""
    invoice_path, clients_path = validate_files(invoice_var.get(), clients_var.get())
    
    # Show status bar
    parent.status_bar.pack(fill=X, side=BOTTOM)
    refit_window(parent)

    # Reset status UI
    parent.status_label.config(text="Alustan...")
    parent.page_progress.config(value=0, mode="determinate")

    # prep the cancel event
    parent.cancel_event.clear()
    parent.btn_cancel.configure(state=NORMAL)
    return invoice_path, clients_path


def on_cancel_ui(parent):
    """Update the UI to reflect cancellation."""
    parent.status_label.config(text="Katkestatud")
    parent.page_progress.config(value=0, mode="determinate")
    parent.btn_cancel.configure(state=DISABLED)
    parent.status_bar.pack_forget()


def extract_person(clients_path: str, cancel_flag):
    """Extract person data from the clients file."""
    persons = extract_person_data(clients_path) # raise ValidationError on error
    if cancel_flag.is_set():
        raise _Cancelled()
    return persons


def finalize_and_open(parent, invoices, dest, template_root, persons, subject, body):
    """Finalize the process and open the email editor."""
    try:
        parent.page_progress.configure(value=HUNDRED_PERCENT)
        parent.status_label.configure(text="Valmis")

        # This blocks until the user clicks OK
        messagebox.showinfo("Info", f"Arved salvestatakse kausta: {dest}")

        if parent.cancel_event.is_set():
            parent.after(0, lambda: on_cancel_ui(parent))
            return

        # Save and open invoices
        _save_and_open_invoices(
            parent, invoices, dest, template_root, persons, subject, body
        )

        # Hide status bar again
        parent.status_bar.pack_forget()
    except Exception as e:
        log_exception(e)
        try:
            call_error(f"Töö ebaõnnestus:\n{e}")
        except Exception as e2:
            log_exception(e2)


def _handle_worker_error(parent, err):
    """Handle errors from the worker thread in the GUI thread."""
    if isinstance(err, _Cancelled):
        parent.after(0, lambda: on_cancel_ui(parent))
        log_exception(_Cancelled("Operation cancelled by user."))
    elif isinstance(err, ValidationError):
        parent.after(0, lambda err=err: messagebox.showerror("Viga", str(err)))
        log_exception(err)
    else:
        parent.after(
            0,
            lambda err=err: messagebox.showerror("Viga", f"Töö ebaõnnestus:\n{err}"),
        )
        log_exception(err)


def _worker_extract_and_process(parent, invoice_path, clients_path, cancel_flag):
    """Extract invoices and persons data."""
    # Extract people
    persons = extract_person(clients_path, parent.cancel_event)
    if parent.cancel_event.is_set():
        raise _Cancelled()

    def on_progress(page_number, total_pages):
        if parent.cancel_event.is_set():
            raise _Cancelled()
        _on_progress_ui(parent, page_number, total_pages, fname)

    # Process ocr
    invoices = _process_ocr(invoice_path, parent.cancel_event, on_progress)

    if parent.cancel_event.is_set():
        raise _Cancelled()
    return persons, invoices


def _worker_finalize_invoices(parent, invoices, invoice_path):
    """Finalize invoices and create destination folder."""
    # Continue processing (back in main thread)
    parent.after(
        0, lambda: parent.status_label.configure(text="Töötlen andmeid...")
    )

    # 4) Create a destination folder
    dest = _create_dest_directory(invoice_path)

    if parent.cancel_event.is_set():
        raise _Cancelled()

    example_invoice = invoices[0]
    subject = f"Arve {example_invoice.period} {example_invoice.year}"
    return dest, subject


def worker(parent, invoice_path, clients_path, template_root, subject, body):
    """Worker thread function to process invoices and open email editor."""
    try:
        # OCR read-through (emits per-page progress)
        fname = os.path.basename(invoice_path)

        persons, invoices = _worker_extract_and_process(
            parent, invoice_path, clients_path, parent.cancel_event
        )

        dest, subject = _worker_finalize_invoices(
            parent, invoices, invoice_path
        )

        # 5) Finalize and open email editor
        parent.after(0, lambda: finalize_and_open(parent, invoices, dest, template_root, persons, subject, body))

    except Exception as e:
        _handle_worker_error(parent, e)
    finally:
        parent.page_progress.config(value=0, mode="determinate")


def start_processing_thread(target, *args):
    """Start a worker thread to process invoices."""
    threading.Thread(target=lambda: target(*args), daemon=True).start()


def get_data_ready(
    parent, invoice_var: str, clients_var: str, template_root, subject, body
):
    """Validate inputs and start processing thread."""
    try:
        invoice_path, clients_path = validate_and_prepare_ui(parent, invoice_var, clients_var)
    except ValidationError as ve:
        messagebox.showerror("Viga", str(ve))
        return

    # Start worker
    start_processing_thread(worker, parent, invoice_path, clients_path, template_root, subject, body)


def open_outlook(persons, invoices_dir, subject, body):
    """Open Outlook email editor with prepared emails."""
    # Compose emails and send them
    ensure_outlook_ready()
    try:
        validate_persons_vs_invoices(persons, invoices_dir)
    except ValidationError as e:
        messagebox.showerror("Viga", str(e))
    save_emails_with_invoices(persons, invoices_dir, subject, body)


def _create_email_subject_section(top, subject):
    """Create the email subject entry section."""
    subject_var = tb.StringVar(value=subject)
    tb.Label(top, text="Meili teema:", bootstyle=INFO).pack(
        anchor="w", padx=12, pady=(12, 4)
    )
    tb.Entry(top, textvariable=subject_var, font=("Helvetica", 15)).pack(
        fill=X, padx=12
    )
    return subject_var


def _create_email_body_section(top, body):
    """Create the email body text section."""
    tb.Label(top, text="Meili sisu:", bootstyle=INFO).pack(
        anchor="w", padx=12, pady=(12, 4)
    )

    # Text + scrollbar in a frame that stretches
    body_frame = tb.Frame(top)
    body_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))

    body_text = tk.Text(body_frame, wrap=tk.WORD, font=("Helvetica", 15), height=10)
    body_text.pack(side="left", fill="both", expand=True)

    yscroll = tb.Scrollbar(body_frame, orient="vertical", command=body_text.yview)
    yscroll.pack(side="right", fill="y")
    body_text.configure(yscrollcommand=yscroll.set)

    body_text.insert(tk.END, body)
    return body_text


def _create_email_buttons_section(top, parent, persons, invoices_dir, subject_var, body_text):
    """Create the email buttons section."""
    btns_frame = tb.Frame(top)
    btns_frame.pack(pady=12, padx=12, fill=X, side=BOTTOM)

    def save_and_close(subject_var):
        subject_val = subject_var.get()
        body_val = body_text.get("1.0", tk.END).strip()
        # body = body_val
        # subject = subject_val
        top.destroy()
        open_outlook(persons, invoices_dir, subject_val, body_val)
        parent.on_emails_saved()

    tb.Button(
        btns_frame,
        text="Salvesta",
        bootstyle=SUCCESS,
        command=lambda: save_and_close(subject_var),
    ).pack(side=RIGHT, padx=6)
    tb.Button(
        btns_frame, text="Tühista", bootstyle=SECONDARY, command=top.destroy
    ).pack(side=RIGHT, padx=6)


def open_email_editor(parent, persons, invoices_dir, subject, body):
    """Open a window to edit email subject and body before sending."""

    parent.btn_cancel.configure(state=DISABLED)

    top = tb.Toplevel(parent)
    top.title("Muuda meili malli")
    top.transient(parent)
    top.grab_set()

    style = tb.Style()
    style.configure("info.TLabel", font=("Helvetica", 15))

    # Subject
    subject_var = _create_email_subject_section(top, subject)

    # Body
    body_text = _create_email_body_section(top, body)

    # Buttons
    _create_email_buttons_section(
        top, parent, persons, invoices_dir, subject_var, body_text
    )

    center_window(top, min_w=500, min_h=450, max_w=750)

class _Cancelled(Exception):
    # "Operation cancelled by user."
    pass


def cancel_current_job(root):
    """Set the cancel event to stop the current job."""
    root.cancel_event.set()  # callable from anywhere
