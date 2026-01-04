import ttkbootstrap as tb
import tkinter as tk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from pathlib import Path
import threading, os, re
import pytesseract
import traceback
import pythoncom

from utils.logging_helper import log_exception
from src.pdf_extractor import separate_invoices, save_each_invoice_as_file
from src.xls_extractor import extract_person_data, ValidationError
from src.email_sender import (
    save_emails_with_invoices,
    ensure_outlook_ready,
    validate_persons_vs_invoices,
)
from src.excel_invoice_extractor import export_excel_to_pdfs

HUNDRED_PERCENT = 100
REFIT_REGEX = r"(\d+)x(\d+)\+(\d+)\+(\d+)"


def _get_invoice_file_extension(key):
    if key == "kommunaal":
        return [("PDF files", "*.pdf")]
    elif key == "kyte":
        return [("Excel failid", "*.xls *.xlsx")]


def select_file(root, label, btn_text_var, new_text, formats=None):
    """Open file dialog and set label and button text."""
    if formats:
        invoice_formats = formats
    else:
        invoice_type = get_selected_invoice_type(root)
        if invoice_type is None:
            return
        invoice_formats = _get_invoice_file_extension(invoice_type.key)

    path = filedialog.askopenfilename(title="Vali fail", filetypes=invoice_formats)
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

    def apply():
        # Ensure status bar is visible (in case it wasn't packed)
        try:
            parent.status_bar.pack(fill=X, side=BOTTOM)
        except Exception:
            pass

        parent.page_progress.configure(value=pct, mode="determinate")
        parent.status_label.configure(text=f"PDF leht {page_number}/{total_pages} - {fname}")
        parent.update_idletasks()

    parent.after(0, apply)


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


def _process_ocr(invoice_path: str, invoice_type: str, cancel_flag, on_progress):
    """Process the invoice PDF with OCR and return extracted invoices."""
    try:
        print(f"Invoice path is: {invoice_path}")
        if invoice_type == "kommunaal":
            invoices = separate_invoices(
                invoice_path,
                on_progress=on_progress,
                cancel_flag=cancel_flag,
            )
        elif invoice_type == "kyte":
            dest_dir = export_excel_to_pdfs(invoice_path)
            print(f'Excelist eksporditud PDF-id kausta: {dest_dir}')
    except pytesseract.TesseractError as e:
        log_exception(e)
        raise ValidationError(
            f"OCR töötlemine ebaõnnestus. Kontrolli, kas Tesseract on õigesti paigaldatud.\n{e}"
        )
    return invoices


def get_selected_invoice_type(parent):
    key = parent.content_type_var.get()
    return parent.invoice_types.get(key)


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

    # prep the cancel event
    parent.cancel_event.clear()
    parent.btn_cancel.configure(state=NORMAL)

    # Schedule UI work on the Tk main thread
    parent.update_idletasks()

    return invoice_path, clients_path


def on_cancel_ui(parent):
    """Update the UI to reflect cancellation."""
    parent.status_label.config(text="Katkestatud")
    parent.page_progress.config(value=0, mode="determinate")
    parent.btn_cancel.configure(state=DISABLED)
    parent.status_bar.pack_forget()


def extract_person(clients_path: str, cancel_flag):
    """Extract person data from the clients file."""
    persons = extract_person_data(clients_path)  # raise ValidationError on error
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


def _worker_extract_and_process(parent, invoice_type, invoice_path, clients_path, cancel_flag, fname):
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
    invoices = _process_ocr(invoice_path, invoice_type, parent.cancel_event, on_progress)

    if parent.cancel_event.is_set():
        raise _Cancelled()
    return persons, invoices


def _worker_finalize_invoices(parent, invoices, invoice_path):
    """Finalize invoices and create destination folder."""
    # Continue processing (back in main thread)
    parent.after(0, lambda: parent.status_label.configure(text="Töötlen andmeid..."))

    # 4) Create a destination folder
    dest = _create_dest_directory(invoice_path)

    if parent.cancel_event.is_set():
        raise _Cancelled()

    example_invoice = invoices[0]
    subject = f"Arve {example_invoice.period} {example_invoice.year}"
    return dest, subject


def worker(parent, invoice_type, invoice_path, clients_path, template_root, subject, body):
    """Worker thread function to process invoices and open email editor."""
    try:
        # OCR read-through (emits per-page progress)
        fname = os.path.basename(invoice_path)

        parent.after(0, lambda: parent.status_bar.pack(fill=X, side=BOTTOM))
        parent.after(0, lambda: parent.status_label.configure(text="Alustan..."))
        parent.after(0, lambda: parent.page_progress.configure(value=0, mode="determinate"))

        persons, invoices = _worker_extract_and_process(
            parent, invoice_type, invoice_path, clients_path, parent.cancel_event, fname
        )

        dest, subject = _worker_finalize_invoices(parent, invoices, invoice_path)

        # 5) Finalize and open email editor
        parent.after(
            0,
            lambda: finalize_and_open(
                parent, invoices, dest, template_root, persons, subject, body
            ),
        )

    except Exception as e:
        _handle_worker_error(parent, e)
    finally:
        def cleanup():
            parent.page_progress.configure(value=0, mode="determinate")
            parent.btn_cancel.configure(state=DISABLED)
        parent.after(0, cleanup)


def start_processing_thread(target, *args):
    """Start a worker thread to process invoices."""
    threading.Thread(target=lambda: target(*args), daemon=True).start()


def get_data_ready(
    parent,
    invoice_var: str,
    clients_var: str,
    template_root,
    content_type_var,
):
    """Validate inputs and start processing thread."""
    invoice_type = get_selected_invoice_type(parent)

    if invoice_type is None:
        return  # Shouldn't happen because UI disabled buttons

    subject = invoice_type.subject
    body = invoice_type.body

    # Make a difference here based on the invoice type

    try:
        invoice_path, clients_path = validate_and_prepare_ui(
            parent, invoice_var, clients_var
        )
    except ValidationError as ve:
        messagebox.showerror("Viga", str(ve))
        return

    # Start worker
    parent.after(
        10,
        lambda: start_processing_thread(
            worker, parent, invoice_type.key, invoice_path, clients_path, template_root, subject, body
        ),
    )


def open_outlook(persons, invoices_dir, subject, body):
    """Open Outlook email editor with prepared emails."""
    # Compose emails and send them
    ensure_outlook_ready()
    try:
        validate_persons_vs_invoices(persons, invoices_dir)
    except ValidationError as e:
        messagebox.showerror("Viga", str(e))
    save_emails_with_invoices(persons, invoices_dir, subject, body)


def _create_email_subject_section(parent, subject):
    """Create the email subject entry section."""
    subject_var = tb.StringVar(value=subject)

    # Row with label + reparation line
    row = tb.Frame(parent)
    row.pack(fill=X, pady=(0, 8))

    tb.Label(
        row, text="Meili teema:", bootstyle=INFO, font=("Segoe UI", 14, "bold")
    ).pack(side=LEFT)
    tb.Separator(row, orient=HORIZONTAL).pack(
        side=LEFT, fill=X, expand=True, padx=(16, 0)
    )

    subject_entry = tb.Entry(parent, textvariable=subject_var, font=("Segoe UI", 13))
    subject_entry.pack(fill=X, pady=(0, 8), ipady=6)

    tb.Label(
        parent,
        text="See kuvatakse meili pealkirjana",
        bootstyle="secondary",
        font=("Segoe UI", 10),
    ).pack(anchor=W, pady=(0, 20))
    return subject_var, subject_entry


def _create_email_body_section(parent, body):
    """Create the email body text section."""

    row = tb.Frame(parent)
    row.pack(fill=X, pady=(0, 8))

    tb.Label(
        row, text="Meili sisu:", bootstyle=INFO, font=("Segoe UI", 14, "bold")
    ).pack(side=LEFT)
    tb.Separator(row, orient=HORIZONTAL).pack(
        side=LEFT, fill=X, expand=True, padx=(16, 0)
    )

    # Text + scrollbar in a frame that stretches
    body_frame = tb.Frame(parent)
    body_frame.pack(fill=BOTH, expand=True, pady=(0, 10))

    body_text = tk.Text(
        body_frame,
        wrap=tk.WORD,
        font=("Segoe UI", 13),
        padx=10,
        pady=10,
        bd=0,
        highlightthickness=1,  # gives a subtle border
        height=12,
    )
    body_text.pack(side=LEFT, fill=BOTH, expand=True)

    yscroll = tb.Scrollbar(
        body_frame,
        orient=VERTICAL,
        command=body_text.yview,
        bootstyle="secondary-round",
    )
    yscroll.pack(side=RIGHT, fill=Y)
    body_text.configure(yscrollcommand=yscroll.set)

    body_text.insert("1.0", body)

    tb.Label(
        parent,
        text="Seda malli kasutatakse automaatselt kõigi valitud arvete saatmisel.",
        bootstyle="secondary",
        font=("Segoe UI", 10),
    ).pack(anchor=W, pady=(4, 0))

    return body_text


def _validate_email_inputs(top, subject_var, body_text):
    subject = subject_var.get().strip()
    body = body_text.get("1.0", "end-1c").strip()

    if not subject:
        tb.dialogs.Messagebox.show_warning(
            "Palun sisesta meili teema.", title="Puuduv teema", parent=top
        )
        log_exception("Meili teema on puudu.")
        return
    if not body:
        tb.dialogs.Messagebox.show_warning(
            "Palun sisesta meili sisu.", title="Puuduv sisu", parent=top
        )
        log_exception("Meili sisu on puudu.")
        return
    
    return subject, body


def _close_email_editor(top):
    try:
        top.grab_release()
    except Exception:
        pass
    top.destroy()


def _show_email_saving_ui(parent):
    def ui():
        try:
            parent.status_bar.pack(fill=X, side=BOTTOM)
            parent.status_label.configure(text="Koostan mustandeid...")
            parent.page_progress.configure(mode="indeterminate")
            parent.page_progress.start(10)
        except Exception:
            pass
        parent.update_idletasks()
    parent.after(0, ui)


def _run_outlook_job_async(parent, persons, invoices_dir, subject, body):
    def job():
        pythoncom.CoInitialize()
        try:
            open_outlook(persons, invoices_dir, subject, body)

            parent.after(0, parent.on_emails_saved)
            parent.after(0, lambda: parent.status_label.configure(text="Mustandid loodud"))

        except Exception as e:
            log_exception(f'Viga mustandite loomisel: {e}')
            traceback_str = traceback.format_exc()
            parent.after(
                0,
                lambda: messagebox.showerror("Viga", f"{e}]\n\n{traceback_str}")
            )
        finally:
            pythoncom.CoUninitialize()
            def cleanup():
                try:
                    parent.page_progress.stop()
                    parent.page_progress.configure(mode="determinate", value=0)
                except Exception:
                    pass
            parent.after(0, cleanup)

    threading.Thread(target=job, daemon=True).start()


def save_and_close(parent, top, subject_var, body_text, persons, invoices_dir):
    # Basic validation
    result = _validate_email_inputs(top, subject_var, body_text)

    if not result:
        return

    subject, body = result
    _close_email_editor(top)
    _show_email_saving_ui(parent)

    _run_outlook_job_async(parent, persons, invoices_dir, subject, body)


def _cancel_email_editor(top, parent):
    top.destroy()

    # Re-enable cancel button if needed
    try:
        parent.btn_cancel.configure(state=NORMAL)
    except Exception as e:
        log_exception(e)


def _create_email_buttons_section(
    top, container, parent, persons, invoices_dir, subject_var, body_text
):
    """Create the email buttons section."""
    btns_frame = tb.Frame(container)
    btns_frame.pack(pady=(18, 0), fill=X, side=BOTTOM)

    tb.Button(
        btns_frame,
        text="Salvesta",
        bootstyle=SUCCESS,
        width=12,
        command=lambda: save_and_close(
            parent, top, subject_var, body_text, persons, invoices_dir
        ),
    ).pack(side=LEFT, ipady=6)

    tb.Button(
        btns_frame,
        text="Tühista",
        bootstyle=SECONDARY,
        command=lambda: _cancel_email_editor(top, parent),
        width=12,
    ).pack(side=LEFT, padx=(0, 12), ipady=6)


def open_email_editor(parent, persons, invoices_dir, subject, body):
    """Open a window to edit email subject and body before sending."""

    parent.btn_cancel.configure(state=DISABLED)

    top = tb.Toplevel(parent)
    top.title("Muuda meili malli")
    top.transient(parent)
    top.grab_set()

    # Size + basic behavior
    top.minsize(760, 520)
    top.geometry("900x620")
    top.resizable()

    # Use one padded container for clean spacing
    container = tb.Frame(top, padding=24)
    container.pack(fill=BOTH, expand=True)

    style = tb.Style()
    style.configure("info.TLabel", font=("Helvetica", 15))

    # Subject
    subject_var, subject_entry = _create_email_subject_section(container, subject)

    # Body
    body_text = _create_email_body_section(container, body)

    # Buttons
    _create_email_buttons_section(
        top, container, parent, persons, invoices_dir, subject_var, body_text
    )

    # Keyboard shortcuts + nicer flow
    subject_entry.bind("<Return>", lambda e: (body_text.focus_set(), "break"))
    top.bind("<Escape>", lambda e: _cancel_email_editor(top, parent))
    top.bind(
        "<Control-s>",
        lambda e: save_and_close(
            parent, top, subject_var, body_text, persons, invoices_dir
        ),
    )
    top.bind(
        "<Control-S>",
        lambda e: save_and_close(
            parent, top, subject_var, body_text, persons, invoices_dir
        ),
    )

    # Focus
    subject_entry.focus_set()
    subject_entry.selection_range(0, END)

    center_window(top, min_w=760, min_h=520, max_w=960)


class _Cancelled(Exception):
    # "Operation cancelled by user."
    pass


def cancel_current_job(root):
    """Set the cancel event to stop the current job."""
    root.cancel_event.set()  # callable from anywhere
