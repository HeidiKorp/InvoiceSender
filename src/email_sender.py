import time
import pythoncom
import win32com.client as win32
from pywintypes import com_error
import shutil, os
import winreg

from utils.logging_helper import log_exception

OUTLOOK_MAIL_ITEM = 0
OUTLOOK_FOLDER_DRAFTS = 16
DRAFT_CATEGORY = "ArveteSaatja"


class OutlookMailer:
    def save_drafts(self, batch) -> None:
        self.ensure_ready()
        outlook = self._connect()
        session = outlook.Session
        invoices_dir = batch.dest_dir
        for person in batch.persons:
            invoice_path = person.invoice_pdf_path(invoices_dir)
            if not invoice_path:
                continue
            for email in person.emails:
                self._create_draft(
                    outlook, invoice_path, email, batch.subject, batch.body
                )
        self._show_drafts_only(outlook, session)

    def send_categorized_drafts(self) -> int:
        outlook = self._connect()
        session = outlook.Session
        drafts_folder = session.GetDefaultFolder(OUTLOOK_FOLDER_DRAFTS)
        messages = drafts_folder.Items
        to_send_ids = []

        try:
            explorer = outlook.ActiveExplorer()
            explorer.ClearSelection()
        except Exception:
            pass

        for i in range(1, messages.Count + 1):
            message = messages.Item(i)
            if DRAFT_CATEGORY in (message.Categories or ""):
                to_send_ids.append((message.EntryID, drafts_folder.StoreID))

        sent_count = 0
        for entry_id, store_id in to_send_ids:
            try:
                message = session.GetItemFromID(entry_id, store_id)
                message.Send()
                sent_count += 1
            except Exception as e:
                log_exception(e, operation="outlook_send_draft")
        return sent_count

    def ensure_ready(self, timeout=15) -> bool:
        start = time.time()
        last_err = None
        while time.time() - start < timeout:
            try:
                app = self._connect()
                session = app.GetNamespace("MAPI")
                session.Logon("", "", False, False)
                _ = session.Accounts
                return True
            except com_error as e:
                last_err = e
                time.sleep(0.5)
            except Exception as e:
                last_err = e
                time.sleep(0.5)
            try:
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass
        raise RuntimeError(
            "Outlook ei ole valmis. Ava Outlook käsitsi, vali profiil (kui küsib) ja proovi uuesti."
        ) from last_err

    def clear_cache(self) -> None:
        win32.gencache.is_readonly = False
        gen_py = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Temp", "gen_py")
        if os.path.isdir(gen_py):
            try:
                shutil.rmtree(gen_py)
            except Exception as e:
                log_exception(e, operation="outlook_cache_clear")
        try:
            win32.gencache.Rebuild()
        except Exception as e:
            log_exception(e, operation="outlook_cache_rebuild")

    def _connect(self):
        try:
            return win32.GetActiveObject("Outlook.Application")
        except Exception:
            return win32.DispatchEx("Outlook.Application")

    def _create_draft(self, outlook, invoice_path, to_email, subject, body):
        mail = outlook.CreateItem(OUTLOOK_MAIL_ITEM)
        if invoice_path:
            mail.Attachments.Add(invoice_path)
        mail.To = to_email
        mail.Subject = subject
        mail.Body = body
        mail.Categories = DRAFT_CATEGORY
        mail.Save()
        return mail

    def _show_drafts_only(self, outlook, session) -> None:
        drafts_folder = session.GetDefaultFolder(OUTLOOK_FOLDER_DRAFTS)
        drafts_folder.Display()
        drafts_id = drafts_folder.EntryID
        explorers = outlook.Explorers
        for index in range(explorers.Count, 0, -1):
            explorer = explorers.Item(index)
            try:
                current = explorer.CurrentFolder
                if current is None or current.EntryID != drafts_id:
                    explorer.Close()
            except Exception as e:
                log_exception(e, operation="outlook_close_explorer")


def send_drafts(parent):
    sent_count = OutlookMailer().send_categorized_drafts()
    parent.hide_send_drafts_button()
    return sent_count


def clear_outlook_cache():
    OutlookMailer().clear_cache()


def get_outlook_path():
    try:
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
        )
        value, _ = winreg.QueryValueEx(key, "")
        return value
    except FileNotFoundError:
        return None
