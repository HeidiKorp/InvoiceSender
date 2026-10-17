import threading
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
OUTLOOK_DIALOG_MESSAGE = (
    "Outlookis on dialoog avatud. Sulge see (salvestamise küsimus, parool või "
    "turvahoiatus) ja ava kaust Mustandid."
)
_DIALOG_OPEN_MARKERS = (
    "dialog box is open",
    "dialogue box is open",
    "dialoogiboks on avatud",
)
_opened_window_handles: list[int] = []
_opened_explorers_lock = threading.Lock()


class OutlookDialogOpen(RuntimeError):
    pass


class OutlookDraftsSavedDialogOpen(OutlookDialogOpen):
    """Drafts are saved; Outlook could not show Drafts because a dialog is open."""


class OutlookMailer:
    def save_drafts(self, batch) -> None:
        self.ensure_ready()
        if batch.cancel_event.is_set():
            return
        outlook = self._connect()
        session = outlook.Session
        invoices_dir = batch.dest_dir
        for person in batch.persons:
            if batch.cancel_event.is_set():
                return
            invoice_path = person.invoice_pdf_path(invoices_dir)
            if not invoice_path:
                continue
            for email in person.emails:
                if batch.cancel_event.is_set():
                    return
                self._create_draft(
                    outlook, invoice_path, email, batch.subject, batch.body
                )
        if batch.cancel_event.is_set():
            return
        self._show_drafts(outlook, session)

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
            except OutlookDialogOpen:
                raise
            except com_error as e:
                if is_outlook_dialog_open(e):
                    raise OutlookDialogOpen(OUTLOOK_DIALOG_MESSAGE) from e
                last_err = e
                time.sleep(0.5)
            except Exception as e:
                if is_outlook_dialog_open(e):
                    raise OutlookDialogOpen(OUTLOOK_DIALOG_MESSAGE) from e
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
        except Exception as error:
            if is_outlook_dialog_open(error):
                raise OutlookDialogOpen(OUTLOOK_DIALOG_MESSAGE) from error
            try:
                return win32.DispatchEx("Outlook.Application")
            except Exception as dispatch_error:
                if is_outlook_dialog_open(dispatch_error):
                    raise OutlookDialogOpen(OUTLOOK_DIALOG_MESSAGE) from dispatch_error
                raise

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

    def _show_drafts(self, outlook, session) -> None:
        before = set(_explorer_window_handles(outlook))
        drafts_folder = session.GetDefaultFolder(OUTLOOK_FOLDER_DRAFTS)
        try:
            drafts_folder.Display()
        except Exception as error:
            if is_outlook_dialog_open(error):
                raise OutlookDraftsSavedDialogOpen(OUTLOOK_DIALOG_MESSAGE) from error
            raise
        _remember_new_explorers(outlook, before)


def is_outlook_dialog_open(error: BaseException) -> bool:
    seen: set[int] = set()
    current: BaseException | None = error
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        if _exception_text_has_dialog(current):
            return True
        current = current.__cause__ or current.__context__
    return False


def close_opened_outlook_windows() -> None:
    with _opened_explorers_lock:
        window_handles = list(_opened_window_handles)
    if not window_handles:
        return
    pythoncom.CoInitialize()
    try:
        try:
            outlook = win32.GetActiveObject("Outlook.Application")
        except Exception as error:
            log_exception(error, operation="outlook_close_on_exit")
            if is_outlook_dialog_open(error):
                raise OutlookDialogOpen(OUTLOOK_DIALOG_MESSAGE) from error
            return
        still_open: list[int] = []
        dialog_blocked = False
        targets = set(window_handles)
        explorers = outlook.Explorers
        for index in range(explorers.Count, 0, -1):
            try:
                explorer = explorers.Item(index)
                window_handle = int(explorer.HWND)
            except Exception as error:
                log_exception(error, operation="outlook_close_on_exit")
                continue
            if window_handle not in targets:
                continue
            try:
                explorer.Close()
                targets.discard(window_handle)
            except Exception as error:
                log_exception(error, operation="outlook_close_on_exit")
                still_open.append(window_handle)
                if is_outlook_dialog_open(error):
                    dialog_blocked = True
        with _opened_explorers_lock:
            _opened_window_handles.clear()
            _opened_window_handles.extend(still_open)
        if dialog_blocked:
            raise OutlookDialogOpen(OUTLOOK_DIALOG_MESSAGE)
    finally:
        pythoncom.CoUninitialize()


def _explorer_window_handles(outlook) -> list[int]:
    window_handles: list[int] = []
    explorers = outlook.Explorers
    for index in range(1, explorers.Count + 1):
        try:
            window_handles.append(int(explorers.Item(index).HWND))
        except Exception:
            continue
    return window_handles


def _remember_new_explorers(outlook, before: set[int]) -> None:
    opened = [
        window_handle
        for window_handle in _explorer_window_handles(outlook)
        if window_handle not in before
    ]
    if not opened:
        return
    with _opened_explorers_lock:
        for window_handle in opened:
            if window_handle not in _opened_window_handles:
                _opened_window_handles.append(window_handle)


def _exception_text_has_dialog(error: BaseException) -> bool:
    parts = [str(error)]
    for arg in getattr(error, "args", ()):
        parts.append(str(arg))
        if isinstance(arg, tuple):
            parts.extend(str(item) for item in arg)
    text = " ".join(parts).lower()
    return any(marker in text for marker in _DIALOG_OPEN_MARKERS)


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
