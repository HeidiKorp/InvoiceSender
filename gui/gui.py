import ttkbootstrap as tb
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import sys
import shutil

from ttkbootstrap.style import Bootstyle
from xls_extractor import extract_person_data, ValidationError
from pdf_extractor import separate_invoices, save_each_invoice_as_file
from email_sender import send_emails_with_invoices, ensure_outlook_ready


DEFAULT_SUBJECT = "Arve"
DEFAULT_BODY = ("Lugupeetud KÜ korteri omanik. Kü edastab järjekordse korteri " 
                        "kuu kulude arve. See on automaatteavitus, palume mitte vastata.")

def call_error(text):
    messagebox.ERROR(text)


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

# TODO: Validate that each file has the correct extension and columns exist

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


def get_data_ready(parent, invoice_var: str, clients_var: str, template_root):
    global DEFAULT_SUBJECT
    invoice_path, clients_path = validate_files(invoice_var.get(), clients_var.get())
    try:
        persons = extract_person_data(clients_path)
        invoices = separate_invoices(invoice_path)
    except ValidationError as e:
        messagebox.showerror("Viga", str(e))
        return

    print(f"Extracted {len(persons)} persons from the clients file.")
    print(f"Extracted {len(invoices)} invoices from the PDF file.")

    invoice_file_parent = Path(invoice_path).resolve().parent
    dest = invoice_file_parent / "arved"

    # Try to create the directory (with parent, ignore if already exists)
    try:
        dest.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        messagebox.showerror("Viga", f"Kausta loomine ebaõnnestus:\n{dest}\n\n{e}")
        sys.exit(1)

    if not dest.exists() or not dest.is_dir():
        messagebox.showerror("Viga", f"Kausta ei õnnestunud luua:\n{dest}")
        sys.exit(1)
    messagebox.showinfo("Info", f"Arved salvestatakse kausta: {dest}")
    invoices_dir = save_each_invoice_as_file(invoices, dest) # returns the full folder path to all individual invoices
    parent.on_folder_created(str(invoices_dir))

    example_invoice = invoices[0]
    DEFAULT_SUBJECT = "Arve " + example_invoice.period + " " + example_invoice.year
    open_email_editor(template_root, persons, invoices_dir)


def open_outlook(persons, invoices_dir, subject, body):
    # Compose emails and send them
    ensure_outlook_ready()
    try:
        send_emails_with_invoices(persons, invoices_dir, subject, body)
    except ValidationError as e:
        messagebox.showerror("Viga", str(e))
    

def open_email_editor(parent, persons, invoices_dir):
    global DEFAULT_SUBJECT, DEFAULT_BODY

    top = tb.Toplevel(parent)
    top.title("Muuda meili malli")
    top.transient(parent)
    top.grab_set()

    center_window(top, 600, 450)

    style = tb.Style()
    style.configure("info.TLabel", font=("Helvetica", 15))

    # Subject
    subject_var = tb.StringVar(value=DEFAULT_SUBJECT)
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

    body_text.insert(tk.END, DEFAULT_BODY)

    # Buttons
    btns_frame = tb.Frame(top)
    btns_frame.pack(pady=12, padx=12, fill=X, side=BOTTOM)

    def save_and_close(subject_var):
        subject = subject_var.get()
        body = body_text.get("1.0", tk.END).strip()
        DEFAULT_BODY = body
        DEFAULT_SUBJECT = subject
        top.destroy()
        open_outlook(persons, invoices_dir, subject, body)

    tb.Button(btns_frame, text="Salvesta", bootstyle=SUCCESS, command=lambda: save_and_close(subject_var)).pack(side=RIGHT, padx=6)
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


def main ():
    # # --- Window setup ---
    root = tb.Window(themename="superhero")
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

    # --- Bottom bar with Next (right corner) ---
    bottom_bar = tb.Frame(root)
    bottom_bar.pack(fill=X, side=BOTTOM)
    tb.Button(bottom_bar, text="Koosta meilid", bootstyle="success", command=lambda: get_data_ready(root, invoice_var, clients_var, root)).pack(side=RIGHT, padx=12, pady=12)
    
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