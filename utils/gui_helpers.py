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


def select_file(label, filetypes, btn_text_var, new_text):
    path = filedialog.askopenfilename(title="Vali fail", filetypes=filetypes)
    if path:
        label.set(path)
        btn_text_var.set(new_text)


def get_window_size(win, min_w=800, min_h=600, max_w=900, max_h=None, margin=40):
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

    width, height, screen_w, screen_h = get_window_size(
        win, min_w, min_h, max_w, max_h, margin
    )

    # keep current position, do not recenter
    m = re.match(r"(\d+)x(\d+)\+(\d+)\+(\d+)", win.geometry())
    x, y = (int(m.group(3)), int(m.group(4))) if m else (100, 100)

    win.geometry(f"{width}x{height}+{x}+{y}")


def _on_progress_ui(parent, page_number, total_pages, fname):
    """Update progress bar and status label in the GUI thread."""
    pct = int(page_number / total_pages * 100) if total_pages else 0
    parent.after(
        0,
        lambda: (
            parent.page_progress.config(value=pct),
            parent.status_label.configure(
                text=f"PDF leht {page_number}/{total_pages} - {fname}"
            ),
        ),
    )


def validate_files(invoice_path: str, clients_path: str):
    if not invoice_path or not clients_path:
        messagebox.showerror("Error", "Palun vali nii arvete kui klientide failid.")
        raise ValueError("Missing invoice file")
    if not Path(invoice_path).is_file():
        messagebox.showerror("Error", f"Arvete fail ei eksisteeri: {invoice_path}")
        raise ValueError("Invalid invoice path")
    if not Path(clients_path).is_file():
        messagebox.showerror("Error", f"Klientide fail ei eksisteeri: {clients_path}")
        raise ValueError("Invalid clients path")
    return invoice_path, clients_path


def get_data_ready(
    parent, invoice_var: str, clients_var: str, template_root, subject, body
):
    # global DEFAULT_SUBJECT
    try:
        invoice_path, clients_path = validate_files(
            invoice_var.get(), clients_var.get()
        )
    except ValidationError as e:
        messagebox.showerror("Viga", str(e))
        return

    # Show status bar
    parent.status_bar.pack(fill=X, side=BOTTOM)
    refit_window(parent)

    # Reset status UI
    parent.status_label.config(text="Alustan...")
    parent.page_progress.config(value=0, mode="determinate")

    # prep the cancel event
    parent.cancel_event.clear()
    parent.btn_cancel.configure(state=NORMAL)

    def on_cancel_ui():
        parent.status_label.config(text="Katkestatud")
        parent.page_progress.config(value=0, mode="determinate")
        parent.btn_cancel.configure(state=DISABLED)
        parent.status_bar.pack_forget()

    def worker():
        try:
            # 1) Extract people (fast, stays here)
            persons = extract_person_data(clients_path)
            if parent.cancel_event.is_set():
                print(f"Cancel event set after extracting persons")
                raise _Cancelled()

            # 2) OCR read-through (emits per-page progress)
            fname = os.path.basename(invoice_path)

            def on_progress(page_number, total_pages):
                if parent.cancel_event.is_set():
                    print(f"Cancel event set during OCR processing")
                    raise _Cancelled()
                _on_progress_ui(parent, page_number, total_pages, fname)


            try:
                invoices = separate_invoices(
                    invoice_path,
                    on_progress=on_progress,
                    cancel_flag=parent.cancel_event,
                )
            except pytesseract.TesseractError as e:
                log_exception(e)
                parent.after(
                    0, lambda: messagebox.showerror("Viga", f"Tesseract OCR viga:\n{e}")
                )

            if parent.cancel_event.is_set():
                print(f"Cancel event set after OCR processing")
                raise _Cancelled()

            # 3) Continue processing (back in main thread)
            parent.after(
                0, lambda: parent.status_label.configure(text="Töötlen andmeid...")
            )

            invoice_file_parent = Path(invoice_path).resolve().parent
            dest = invoice_file_parent / "arved"

            # Try to create the directory (with parent, ignore if already exists)
            try:
                dest.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                parent.after(
                    0,
                    lambda: messagebox.showerror(
                        "Viga", f"Kausta loomine ebaõnnestus:\n{dest}\n\n{e}"
                    ),
                )
                # sys.exit(1)
                return

            if parent.cancel_event.is_set():
                print(f"Cancel event set after creating directory")
                raise _Cancelled()

            if not dest.exists() or not dest.is_dir():
                parent.after(
                    0,
                    lambda: messagebox.showerror(
                        "Viga", f"Kausta ei õnnestunud luua:\n{dest}"
                    ),
                )
                # sys.exit(1)
                return

            example_invoice = invoices[0]
            subject = "Arve " + example_invoice.period + " " + example_invoice.year

            def show_message_and_then_open():
                try:
                    parent.page_progress.configure(value=100)
                    parent.status_label.configure(text="Valmis")

                    # This blocks until the user clicks OK
                    messagebox.showinfo("Info", f"Arved salvestatakse kausta: {dest}")

                    if parent.cancel_event.is_set():
                        on_cancel_ui()
                        return

                    # After OK -> open the email editor
                    invoices_dir = save_each_invoice_as_file(
                        invoices, dest
                    )  # returns the full folder path to all individual invoices
                    parent.on_folder_created(str(invoices_dir))
                    open_email_editor(
                        template_root, persons, invoices_dir, subject, body
                    )

                    # Hide status bar again
                    parent.status_bar.pack_forget()
                except Exception as e:
                    log_exception(e)
                    try:
                        call_error(f"Töö ebaõnnestus:\n{e}")
                    except:
                        pass

            parent.after(0, show_message_and_then_open)

        except _Cancelled:
            print("Operation cancelled by user (caught in worker).")
            parent.after(0, on_cancel_ui)
            log_exception(_Cancelled("Operation cancelled by user."))
        except ValidationError as e:
            parent.after(0, lambda err=e: messagebox.showerror("Viga", str(err)))
        except Exception as e:
            parent.after(
                0,
                lambda err=e: messagebox.showerror("Viga", f"Töö ebaõnnestus:\n{err}"),
            )
        finally:
            # parent.status_label.config(text="Valmis")
            parent.page_progress.config(value=0, mode="determinate")

    threading.Thread(target=worker, daemon=True).start()


def open_outlook(persons, invoices_dir, subject, body):
    # Compose emails and send them
    ensure_outlook_ready()
    try:
        validate_persons_vs_invoices(persons, invoices_dir)
    except ValidationError as e:
        messagebox.showerror("Viga", str(e))
    save_emails_with_invoices(persons, invoices_dir, subject, body)


def open_email_editor(parent, persons, invoices_dir, subject, body):
    # global DEFAULT_SUBJECT, DEFAULT_BODY

    parent.btn_cancel.configure(state=DISABLED)

    top = tb.Toplevel(parent)
    top.title("Muuda meili malli")
    top.transient(parent)
    top.grab_set()

    style = tb.Style()
    style.configure("info.TLabel", font=("Helvetica", 15))

    # Subject
    subject_var = tb.StringVar(value=subject)
    tb.Label(top, text="Meili teema:", bootstyle=INFO).pack(
        anchor="w", padx=12, pady=(12, 4)
    )
    tb.Entry(top, textvariable=subject_var, font=("Helvetica", 15)).pack(
        fill=X, padx=12
    )

    # Body
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

    # Buttons
    btns_frame = tb.Frame(top)
    btns_frame.pack(pady=12, padx=12, fill=X, side=BOTTOM)

    def save_and_close(subject_var, subject, body):
        subject_val = subject_var.get()
        body_val = body_text.get("1.0", tk.END).strip()
        body = body_val
        subject = subject_val
        top.destroy()
        open_outlook(persons, invoices_dir, subject_val, body_val)
        parent.on_emails_saved()

    tb.Button(
        btns_frame,
        text="Salvesta",
        bootstyle=SUCCESS,
        command=lambda: save_and_close(subject_var, subject, body),
    ).pack(side=RIGHT, padx=6)
    tb.Button(
        btns_frame, text="Tühista", bootstyle=SECONDARY, command=top.destroy
    ).pack(side=RIGHT, padx=6)

    center_window(top, min_w=500, min_h=450, max_w=750)


def call_error(text):
    messagebox.ERROR(text)


class _Cancelled(Exception):
    # "Operation cancelled by user."
    pass


def cancel_current_job(root):
    root.cancel_event.set()  # callable from anywhere
