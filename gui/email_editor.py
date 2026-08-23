import ttkbootstrap as tb
import tkinter as tk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import pythoncom
import threading
import traceback

from utils.logging_helper import log_exception
from utils.config_helpers import load_template_config, save_template_config
from utils.gui_helpers import center_window
from src.invoice_batch import InvoiceBatch


class EmailEditor:
    def __init__(self, parent, batch: InvoiceBatch):
        self.parent = parent
        self.batch = batch

        parent.btn_cancel.configure(state=DISABLED)

        self.window = tb.Toplevel(parent)
        self.window.title("Muuda meili malli")
        self.window.transient(parent)
        self.window.grab_set()
        self.window.minsize(760, 520)
        self.window.geometry("900x620")

        container = tb.Frame(self.window, padding=24)
        container.pack(fill=BOTH, expand=True)

        style = tb.Style()
        style.configure("info.TLabel", font=("Helvetica", 15))

        self.subject_var, subject_entry = self._create_subject_section(
            container, batch.subject
        )
        self.body_text = self._create_body_section(container, batch.body)
        self._create_buttons(container)

        subject_entry.bind("<Return>", lambda e: (self.body_text.focus_set(), "break"))
        self.window.bind("<Escape>", lambda e: self._cancel())
        self.window.bind("<Control-s>", lambda e: self._save_and_close())
        self.window.bind("<Control-S>", lambda e: self._save_and_close())

        subject_entry.focus_set()
        subject_entry.selection_range(0, END)
        center_window(self.window, min_w=760, min_h=520, max_w=960)

    def _save_and_close(self) -> None:
        result = self._validate()
        if not result:
            return
        subject, body = result
        self.batch.subject = subject
        self.batch.body = body
        self._close()
        self._show_saving_ui()
        self._run_outlook_job()

    def _save_template(self) -> None:
        subject, body = self._current_values()
        filename = filedialog.asksaveasfilename(
            parent=self.window,
            title="Salvesta meili mall (.cfg)",
            defaultextension=".cfg",
            filetypes=[("Template config", "*.cfg"), ("All files", "*.*")],
        )
        if not filename:
            return
        try:
            save_template_config(filename, subject, body)
            messagebox.showinfo("Salvestatud", f"Meili mall salvestatud:\n{filename}")
        except Exception as e:
            log_exception(e, operation="save_email_template")
            messagebox.showerror("Viga", f"Meili malli salvestamine ebaõnnestus:\n{e}")

    def _load_template(self) -> None:
        filename = filedialog.askopenfilename(
            parent=self.window,
            title="Laadi meili mall (.cfg)",
            defaultextension=".cfg",
            filetypes=[("Template config", "*.cfg"), ("All files", "*.*")],
        )
        if not filename:
            return
        try:
            subject, body = load_template_config(filename)
            self.subject_var.set(subject)
            prev_state = str(self.body_text.cget("state"))
            if prev_state != "normal":
                self.body_text.config(state="normal")
            self.body_text.delete("1.0", "end")
            self.body_text.insert("1.0", body)
            self.body_text.mark_set("insert", "1.0")
            self.body_text.see("1.0")
            if prev_state != "normal":
                self.body_text.config(state=prev_state)
        except Exception as e:
            log_exception(e, operation="load_email_template")
            messagebox.showerror("Viga", f"Meili malli laadimine ebaõnnestus:\n{e}")

    def _cancel(self) -> None:
        self._close()
        try:
            self.parent.btn_cancel.configure(state=NORMAL)
        except Exception as e:
            log_exception(e, operation="email_editor_cancel")

    def _validate(self) -> tuple[str, str] | None:
        subject, body = self._current_values()
        if not subject:
            tb.dialogs.Messagebox.show_warning(
                "Palun sisesta meili teema.",
                title="Puuduv teema",
                parent=self.window,
            )
            log_exception("Meili teema on puudu.", operation="email_editor")
            return None
        if not body:
            tb.dialogs.Messagebox.show_warning(
                "Palun sisesta meili sisu.",
                title="Puuduv sisu",
                parent=self.window,
            )
            log_exception("Meili sisu on puudu.", operation="email_editor")
            return None
        return subject, body

    def _current_values(self) -> tuple[str, str]:
        return (
            self.subject_var.get().strip(),
            self.body_text.get("1.0", "end-1c").strip(),
        )

    def _show_saving_ui(self) -> None:
        def ui():
            try:
                self.parent.status_bar.pack(fill=X, side=BOTTOM)
                self.parent.status_label.configure(text="Koostan mustandeid...")
                self.parent.page_progress.configure(mode="indeterminate")
                self.parent.page_progress.start(10)
            except Exception:
                pass
            self.parent.update_idletasks()

        self.parent.after(0, ui)

    def _run_outlook_job(self) -> None:
        parent = self.parent
        batch = self.batch

        def job():
            pythoncom.CoInitialize()
            try:
                batch.create_drafts()
                parent.after(0, parent.on_emails_saved)
                parent.after(
                    0, lambda: parent.status_label.configure(text="Mustandid loodud")
                )
            except Exception as e:
                log_exception(e, operation="outlook_drafts")
                traceback_str = traceback.format_exc()
                parent.after(
                    0,
                    lambda: messagebox.showerror(
                        "Viga", f"{e}\n\n{traceback_str}"
                    ),
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

    def _close(self) -> None:
        try:
            self.window.grab_release()
        except Exception:
            pass
        self.window.destroy()

    def _create_subject_section(self, parent, subject):
        subject_var = tb.StringVar(value=subject)
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

    def _create_body_section(self, parent, body):
        row = tb.Frame(parent)
        row.pack(fill=X, pady=(0, 8))
        tb.Label(
            row, text="Meili sisu:", bootstyle=INFO, font=("Segoe UI", 14, "bold")
        ).pack(side=LEFT)
        tb.Separator(row, orient=HORIZONTAL).pack(
            side=LEFT, fill=X, expand=True, padx=(16, 0)
        )
        body_frame = tb.Frame(parent)
        body_frame.pack(fill=BOTH, expand=True, pady=(0, 10))
        body_text = tk.Text(
            body_frame,
            wrap=tk.WORD,
            font=("Segoe UI", 13),
            padx=10,
            pady=10,
            bd=0,
            highlightthickness=1,
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

    def _create_buttons(self, container):
        btns_frame = tb.Frame(container)
        btns_frame.pack(pady=(18, 0), fill=X, side=BOTTOM)
        tb.Button(
            btns_frame,
            text="Salvesta",
            bootstyle=SUCCESS,
            width=12,
            command=self._save_and_close,
        ).pack(side=LEFT, ipady=6)
        tb.Button(
            btns_frame,
            text="Tühista",
            bootstyle=SECONDARY,
            command=self._cancel,
            width=12,
        ).pack(side=LEFT, padx=(0, 12), ipady=6)
        tb.Button(
            btns_frame,
            text="Laadi mall",
            bootstyle="info-outline",
            command=self._load_template,
            width=14,
        ).pack(side=RIGHT, ipady=6)
        tb.Button(
            btns_frame,
            text="Salvesta mall",
            bootstyle="primary-outline",
            width=16,
            command=self._save_template,
        ).pack(side=RIGHT, padx=(0, 12), ipady=6)
