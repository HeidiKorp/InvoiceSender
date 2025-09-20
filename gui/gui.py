import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from pathlib import Path
from xls_extractor import extract_person_data
from pdf_extractor import separate_invoices, save_each_invoice_as_file
from email_sender import send_emails_with_invoices, ensure_outlook_ready


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

# def get_data_ready():



# # --- Window setup ---
root = tb.Window(themename="superhero")
root.title("Invoice Sender")
root.resizable(True, True)
center_window(root, 600, 400)

# --- Bottom bar with Next (right corner) ---
bottom_bar = tb.Frame(root)
bottom_bar.pack(fill=X, side=BOTTOM)
tb.Button(bottom_bar, text="Järgmine", bootstyle="success", command=lambda: validate_files()).pack(side=RIGHT, padx=12, pady=12)

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