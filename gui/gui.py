import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from pdf_extractor import ocr_pdf_all_pages
import sys, threading, os
import shutil
import traceback
import pytesseract

from ttkbootstrap.style import Bootstyle
from xls_extractor import extract_person_data, ValidationError
from pdf_extractor import separate_invoices, save_each_invoice_as_file, check_ocr_environment
from email_sender import send_emails_with_invoices, ensure_outlook_ready, clear_outlook_cache, validate_persons_vs_invoices


def get_log_path():
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "error.log")

def log_exc_triple(exc_type, exc_value, exc_tb):
    # Write errors to a log file next to the exe, even in PyInstaller
    log_path = get_log_path()
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write("=== Exception ===\n")
            traceback.print_exception(exc_type, exc_value, exc_tb, file=f)
    except Exception:
        pass

sys.excepthook = log_exc_triple

def _thread_excepthook(args):
    log_exc_triple(args.exc_type, args.exc_value, args.exc_traceback)


def log_exception(e: Exception):
    exc_type = type(e)
    exc_value = e
    exc_tb = e.__traceback__
    log_exc_triple(exc_type, exc_value, exc_tb)


def delete_old_error_log():
    log_path = get_log_path()
    if os.path.exists(log_path):
        try:
            os.remove(log_path)
        except Exception:
            # If deletion fails, overwrite instead of crash
            open(log_path, "w").close()


threading.excepthook = _thread_excepthook

def call_error(text):
    messagebox.ERROR(text)


class _Cancelled(Exception):
    # "Operation cancelled by user."
    pass

def cancel_current_job(root):
    root.cancel_event.set() # callable from anywhere


def select_file(label, filetypes, btn_text_var, new_text):
    path = filedialog.askopenfilename(title="Vali fail", filetypes=filetypes)
    if path:
        label.set(path)
        btn_text_var.set(new_text)


def center_window(win, width, height):
    """Center a window on the screen with given width/height."""
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()

    win.minsize(width, height) # user can't shrink below this

    # Enforce minimum dimensions, but also cap at screen size
    w = max(width, min(sw, width))
    h = max(height, min(sh, height))

    x = (sw - w) // 2
    y = (sh - h) // 2

    win.geometry(f"{w}x{h}+{x}+{y}")


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


def get_data_ready(parent, invoice_var: str, clients_var: str, template_root, subject, body):
    # global DEFAULT_SUBJECT
    try:
        invoice_path, clients_path = validate_files(invoice_var.get(), clients_var.get())
    except ValidationError as e:
        messagebox.showerror("Viga", str(e))
        return

    # Show status bar
    parent.status_bar.pack(fill=X, side=BOTTOM)

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
                print(f'Cancel event set after extracting persons')
                raise _Cancelled()

            # 2) OCR read-through (emits per-page progress)
            fname = os.path.basename(invoice_path)

            def on_progress(page_number, total_pages):
                if parent.cancel_event.is_set():
                    print(f'Cancel event set during OCR processing')
                    raise _Cancelled()
                pct = int(page_number / total_pages * 100) if total_pages else 0
                parent.after(0, lambda: (
                    parent.page_progress.config(value=pct),
                    parent.status_label.configure(text=f"PDF leht {page_number}/{total_pages} - {fname}")
                ))

            try:
                invoices = separate_invoices(invoice_path, on_progress=on_progress, cancel_flag=parent.cancel_event)      
            except pytesseract.TesseractError as e:
                log_exception(e)
                parent.after(0, lambda: messagebox.showerror("Viga", f"Tesseract OCR viga:\n{e}"))     

            if parent.cancel_event.is_set():
                print(f'Cancel event set after OCR processing')
                raise _Cancelled() 
            
            # 3) Continue processing (back in main thread)
            parent.after(0, lambda: parent.status_label.configure(text="Töötlen andmeid..."))

            invoice_file_parent = Path(invoice_path).resolve().parent
            dest = invoice_file_parent / "arved"

            # Try to create the directory (with parent, ignore if already exists)
            try:
                dest.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                parent.after(0, lambda: messagebox.showerror("Viga", f"Kausta loomine ebaõnnestus:\n{dest}\n\n{e}"))
                # sys.exit(1)
                return

            if parent.cancel_event.is_set():
                print(f'Cancel event set after creating directory')
                raise _Cancelled()

            if not dest.exists() or not dest.is_dir():
                parent.after(0, lambda: messagebox.showerror("Viga", f"Kausta ei õnnestunud luua:\n{dest}"))
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
                    invoices_dir = save_each_invoice_as_file(invoices, dest) # returns the full folder path to all individual invoices
                    parent.on_folder_created(str(invoices_dir))
                    open_email_editor(template_root, persons, invoices_dir, subject, body)

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
            parent.after(0, lambda err=e: messagebox.showerror("Viga", f"Töö ebaõnnestus:\n{err}"))
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
    send_emails_with_invoices(persons, invoices_dir, subject, body)
    

def open_email_editor(parent, persons, invoices_dir, subject, body):
    # global DEFAULT_SUBJECT, DEFAULT_BODY

    parent.btn_cancel.configure(state=DISABLED)

    top = tb.Toplevel(parent)
    top.title("Muuda meili malli")
    top.transient(parent)
    top.grab_set()

    center_window(top, 600, 450)

    style = tb.Style()
    style.configure("info.TLabel", font=("Helvetica", 15))

    # Subject
    subject_var = tb.StringVar(value=subject)
    tb.Label(top, text="Meili teema:", bootstyle=INFO).pack(anchor="w", padx=12, pady=(12, 4))
    tb.Entry(top, textvariable=subject_var, font=("Helvetica", 15)).pack(fill=X, padx=12)

    # Body
    tb.Label(top, text="Meili sisu:", bootstyle=INFO).pack(anchor="w", padx=12, pady=(12, 4))
    
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

    tb.Button(btns_frame, text="Salvesta", bootstyle=SUCCESS, command=lambda: save_and_close(subject_var, subject, body)).pack(side=RIGHT, padx=6)
    tb.Button(btns_frame, text="Tühista", bootstyle=SECONDARY, command=top.destroy).pack(side=RIGHT, padx=6)


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


def main():
    # Clean Outlook cache on startup
    clear_outlook_cache()

    # Set up default subject, body
    subject = "Arve"
    body = ("Lugupeetud KÜ korteri omanik. Kü edastab järjekordse korteri " 
                        "kuu kulude arve. See on automaatteavitus, palume mitte vastata.")

    # # --- Window setup ---
    root = tb.Window(themename="superhero")

    delete_old_error_log()

    if not check_ocr_environment():
        messagebox.showerror(
            "Tesseract puudub",
            "Tesseract OCR ei ole selles arvutis paigaldatud või ei leitud teekonda.\n\n"
            "Palun paigalda Tesseract OCR ja proovi uuesti."
        )
        root.destroy()

    def tk_report_callback_exception(exc_type, exc_value, exc_tb):
        # Log it
        log_exc_triple(exc_type, exc_value, exc_tb)
        # Also show a more user-friendly error
        messagebox.showerror(
            "Viga",
            f"Kasutajaliidese viga:\n{exc_type.__name__}: {exc_value}"
        )

    root.report_callback_exception = tk_report_callback_exception

    root.title("Invoice Sender")
    root.resizable(True, True)
    center_window(root, 800, 600)
    root.update_idletasks()
    root.deiconify()
    root.lift()
    root.focus_force()

    invoice_var = tb.StringVar()
    clients_var = tb.StringVar()
    invoices_dir_var = tb.StringVar()

    style = tb.Style()

    # Define a custom font and size
    style.configure("TButton", font=("Helvetica", 18))
    style.configure("success.TButton", font=("Helvetica", 18))
    style.configure("TLabel", font=("Helvetica", 15))
    style.configure("info.TLabel", font=("Helvetica", 15))

    # shared state
    root.cancel_event = threading.Event()
    root.current_worker = None

    # --- Bottom bar with Next (right corner) ---
    bottom_bar = tb.Frame(root)
    bottom_bar.pack(fill=X, side=BOTTOM)
    tb.Button(bottom_bar, text="Koosta meilid", bootstyle="success", command=lambda: get_data_ready(root, invoice_var, clients_var, root, subject, body)).pack(side=RIGHT, padx=12, pady=12)
    
    # Cancel button
    root.btn_cancel = tb.Button(bottom_bar, text="Katkesta", bootstyle="danger", command=lambda: cancel_current_job(root))
    root.btn_cancel.pack(side=RIGHT, padx=0, pady=12)
    root.btn_cancel.configure(state=DISABLED)

    # --- Status bar (bottom, above the button row) ---
    status_bar = tb.Frame(root)
    status_bar.pack(fill=X, side=BOTTOM)

    root.status_label = tb.Label(status_bar, text="Valmis", bootstyle=INFO)
    root.status_label.pack(side=LEFT, padx=10, pady=8)

    root.page_progress = tb.Progressbar(status_bar, orient="horizontal", mode="determinate", 
        maximum=100, bootstyle=INFO)

    root.page_progress.pack(side=LEFT, fill=X, expand=True, padx=(10, 12), pady=8)

    # Hide status bar initially
    status_bar.pack_forget()
    root.status_bar = status_bar # store reference on root so we can show/hide later 

    root.invoices_dir_var = tb.StringVar(value="")
    btn_delete_invoices = tb.Button(bottom_bar, text="Kustuta arvekaust", bootstyle=DANGER, command=lambda: delete_folder(root, root.invoices_dir_var.get()))

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

    # --- Center content ---
    content = tb.Frame(root)
    content.pack(expand=True)

    # Invoice
    btn_text_invoice = tk.StringVar(value="Vali arvete fail")
    btn_text_clients = tk.StringVar(value="Vali klientide fail")

    btn_invoice = tb.Button(content, textvariable=btn_text_invoice, bootstyle=INFO, command=lambda: select_file(invoice_var, [("PDF files", "*.pdf")], btn_text_invoice, "Muuda arvete faili"))
    btn_invoice.grid(row=0, column=0, padx=22, pady=22)
    lbl_invoice = tb.Label(content, textvariable=invoice_var, wraplength=680, foreground="#9aa0a6")
    lbl_invoice.grid(row=1, column=0, padx=12, pady=12)


    # Clients
    btn_clients = tb.Button(content, textvariable=btn_text_clients, bootstyle=INFO, command=lambda: select_file(clients_var, [("XLS files", "*.xls"), ("XLSX files", "*.xlsx")], btn_text_clients, "Muuda klientide faili"))
    btn_clients.grid(row=2, column=0, padx=12, pady=12)
    lbl_clients = tb.Label(content, textvariable=clients_var, wraplength=680, foreground="#9aa0a6")
    lbl_clients.grid(row=3, column=0, padx=12, pady=12)

    # Center the column
    content.grid_columnconfigure(0, weight=1)

    root.mainloop()