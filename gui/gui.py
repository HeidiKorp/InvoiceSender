import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from pathlib import Path
from xls_extractor import extract_person_data
from pdf_extractor import separate_invoices, save_each_invoice_as_file
from email_sender import send_emails_with_invoices, ensure_outlook_ready

DEFAULT_SUBJECT = "Arve"
DEFAULT_BODY = ("Lugupeetud KÜ korteri omanik. Kü edastab järjekordse korteri " 
                        "kuu kulude arve. See on automaatteavitus, palume mitte vastata.")


def select_file(label):
    path = filedialog.askopenfilename(title="Vali arvete fail", filetypes=[("All files", "*.*"), ("PDF files", "*.pdf"), ("XLS files", "*.xls"), ("XLSX files", "*.xlsx")])
    if path:
        label.set(path)


def center_window(win, width, height):
    """Center a window on the screen with given width/height."""
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - width) // 2
    y = (sh - height) // 2
    win.geometry(f"{width}x{height}+{x}+{y}")

# TODO: Validate that each file has the correct extension and columns exist

def validate_files():
    invoice_path = invoice_var.get()
    clients_path = clients_var.get()
    if not invoice_path or not clients_path:
        messagebox.showerror("Error", "Palun vali nii arvete kui klientide failid.")
        return False
    if not Path(invoice_path).is_file():
        messagebox.showerror("Error", f"Arvete fail ei eksisteeri: {invoice_path}")
        return False
    if not Path(clients_path).is_file():
        messagebox.showerror("Error", f"Klientide fail ei eksisteeri: {clients_path}")
        return False
    return True


def get_data_ready(subject, body):
    validate_files()
    persons = extract_person_data(clients_var.get())
    invoices = separate_invoices(invoice_var.get())
    print(f"Extracted {len(persons)} persons from the clients file.")
    print(f"Extracted {len(invoices)} invoices from the PDF file.")

    invoice_file_parent = Path(invoice_var.get()).resolve().parent
    dest = invoice_file_parent / "arved"
    messagebox.showinfo("Info", f"Arved salvestatakse kausta: {dest}")
    invoices_dir = save_each_invoice_as_file(invoices, dest) # returns the full folder path to all individual invoices

    # Compose emails and send them
    ensure_outlook_ready()
    send_emails_with_invoices(persons, invoices_dir, subject, body)


def open_email_editor(parent):
    global DEFAULT_SUBJECT, DEFAULT_BODY
    top = tb.Toplevel(parent)
    top.title("Muuda meili malli")
    top.transient(parent)
    top.grab_set()
    center_window(top, 500, 350)

    # Subject
    subject_var = tb.StringVar(value=DEFAULT_SUBJECT)
    tb.Label(top, text="Meili teema:", bootstyle=INFO).pack(anchor="w", padx=12, pady=(12, 4))
    tb.Entry(top, textvariable=subject_var).pack(fill=X, padx=12)

    # Body
    tb.Label(top, text="Meili sisu:", bootstyle=INFO).pack(anchor="w", padx=12, pady=(12, 4))
    body_text = tk.Text(top, wrap=tk.WORD, height=10)
    body_text.pack(fill=BOTH, padx=12, pady=(0, 8), expand=True)
    body_text.insert(tk.END, DEFAULT_BODY)

    # Buttons
    btns_frame = tb.Frame(top)
    btns_frame.pack(pady=12, padx=12, fill=X, side=BOTTOM)

    def save_and_close():
        nonlocal subject_var, body_text
        subject = subject_var.get()
        body = body_text.get("1.0", tk.END).strip()
        top.destroy()
        get_data_ready(subject, body)

    tb.Button(btns_frame, text="Salvesta", bootstyle=SUCCESS, command=save_and_close).pack(side=RIGHT, padx=6)
    tb.Button(btns_frame, text="Tühista", bootstyle=SECONDARY, command=top.destroy).pack(side=RIGHT, padx=6)
    

# # --- Window setup ---
root = tb.Window(themename="superhero")
root.title("Invoice Sender")
root.resizable(True, True)
center_window(root, 600, 400)

# --- Bottom bar with Next (right corner) ---
bottom_bar = tb.Frame(root)
bottom_bar.pack(fill=X, side=BOTTOM)
tb.Button(bottom_bar, text="Koosta meilid", bootstyle="success", command=lambda: open_email_editor(root)).pack(side=RIGHT, padx=12, pady=12)

# --- Center content ---
content = tb.Frame(root)
content.pack(expand=True)

invoice_var = tb.StringVar()
clients_var = tb.StringVar()

# Incoice
btn_invoice = tb.Button(content, text="Vali arvete fail", bootstyle=INFO, command=lambda: select_file(invoice_var))
btn_invoice.grid(row=0, column=0, padx=12, pady=12)
lbl_invoice = tb.Label(content, textvariable=invoice_var, wraplength=480, foreground="#9aa0a6")
lbl_invoice.grid(row=1, column=0, padx=12, pady=12)


# Clients
btn_clients = tb.Button(content, text="Vali kliendi info fail", bootstyle=INFO, command=lambda: select_file(clients_var))
btn_clients.grid(row=2, column=0, padx=12, pady=12)
lbl_clients = tb.Label(content, textvariable=clients_var, wraplength=480, foreground="#9aa0a6")
lbl_clients.grid(row=3, column=0, padx=12, pady=12)

# Center the column
content.grid_columnconfigure(0, weight=1)

root.mainloop()