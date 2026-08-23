import time
import subprocess
import win32com.client as win32
from pathlib import Path
from pywintypes import com_error
import shutil, os
import winreg
from collections import Counter

from utils.logging_helper import log_exception

OUTLOOK_MAIL_ITEM = 0
OUTLOOK_FOLDER_DRAFTS = 16
INVOICE_FILE_EXTENSIONS = frozenset({".pdf"})


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


def clear_outlook_cache():
    # Make gencache readable/writable
    win32.gencache.is_readonly = False

    # Purge stale/corrupt cache in disk
    gen_py = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Temp", "gen_py")

    if os.path.isdir(gen_py):
        try:
            shutil.rmtree(gen_py)
        except Exception as e:
            log_exception(e)

    # Rebuild cache
    try:
        win32.gencache.Rebuild()
    except Exception as e:
        log_exception(e)


def _try_start_outlook():
    try:
        outlook_path = get_outlook_path()
        if not outlook_path or not os.path.exists(outlook_path):
            return False
        subprocess.Popen(
            [outlook_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        return True
    except Exception:
        return False

def ensure_outlook_ready(timeout=15):
    """
    Non-interactive readiness check. 
    - No profile dialogs
    - Works in a background thread.
    - If outlook isn't configured/ready, raises RuntimeError
    """
    start = time.time()
    app = None

    _try_start_outlook()

    while time.time() - start < timeout:
        try: # Try attach to running Outlook first (often fastest)
            try:
                app = win32.GetActiveObject("Outlook.Application")
            except Exception:
                app = win32.DispatchEx("Outlook.Application")
            
            session = app.GetNamespace("MAPI")

            # IMPORTANT: do NOT show dialogs from background thread
            session.Logon("", "", False, False)

            _ = session.Accounts  # Accessing Accounts to ensure it's fully loaded
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


def _apartment_key(value) -> str:
    return str(value).strip()


def apartments_from_persons(persons):
    return {_apartment_key(p.apartment) for p in persons if _apartment_key(p.apartment)}


def apartments_from_invoices(invoices_dir, exts=None):
    if exts is None:
        exts = INVOICE_FILE_EXTENSIONS
    counts = Counter()
    for p in Path(invoices_dir).iterdir():
        if p.is_file() and p.suffix.lower() in exts:
            apt = p.stem.strip()
            if apt:
                counts[apt] += 1
    return counts


def _check_missing_invoices(person_apts, invoice_apts):
    """ Check that each person has a corresponding invoice file. """
    return sorted(person_apts - invoice_apts, key=str)


def _check_extra_invoices(person_apts, invoice_apts):
    """ Check for invoice files that don't match any person. """
    return sorted(invoice_apts - person_apts, key=str)


def _check_duplicate_invoices(invoice_counts):
    """ Check for duplicate invoice files for the same apartment. """
    return sorted([apt for apt, c in invoice_counts.items() if c > 1], key=str)


def _build_validation_errors(missing, extra, duplicates):
    problems = []
    if missing:
        problems.append(f"Puuduvad arved korteritele: {', '.join(missing)}.\n")
    if extra:
        problems.append(
            f"Arved, millele ei leitud klienti: {', '.join(extra)}.\n"
        )
    if duplicates:
        problems.append(
            f"Duplikaatsed arvefailid korteritele: {', '.join(duplicates)}.\n"
        )
    return problems


def _matched_persons_by_apartment(persons, invoice_apts, excluded_apts):
    """Keep every person whose apartment has a unique invoice. Emails are not used."""
    matched = []
    for person in persons:
        apartment = _apartment_key(person.apartment)
        if apartment in invoice_apts and apartment not in excluded_apts:
            matched.append(person)
    return matched


def _pair_persons_to_invoice_apartments(persons, invoice_counts):
    person_apts = apartments_from_persons(persons)
    invoice_apts = set(invoice_counts.keys())

    missing_for_people = _check_missing_invoices(person_apts, invoice_apts)
    extra_invoices = _check_extra_invoices(person_apts, invoice_apts)
    duplicates = _check_duplicate_invoices(invoice_counts)

    problems = _build_validation_errors(
        missing_for_people, extra_invoices, duplicates
    )
    excluded_apts = set(missing_for_people) | set(duplicates)
    matched_persons = _matched_persons_by_apartment(
        persons, invoice_apts, excluded_apts
    )
    return matched_persons, problems


def validate_persons_vs_invoice_items(persons, invoices):
    """Pair people to this run's invoices by apartment number only.

    Emails are not part of the match. One apartment may have several
    addresses, and the same address may appear on several apartments.
    """
    invoice_counts = Counter(
        _apartment_key(invoice.apartment)
        for invoice in invoices
        if _apartment_key(invoice.apartment)
    )
    return _pair_persons_to_invoice_apartments(persons, invoice_counts)


def validate_persons_vs_invoices(persons, invoices_dir):
    """Report mismatches from invoice files and return only unique matches."""
    return _pair_persons_to_invoice_apartments(
        persons, apartments_from_invoices(invoices_dir)
    )


def _create_email_draft(
    outlook,
    invoice_path: str,
    to_email: str,
    subject: str,
    body: str,
    category: str = "ArveteSaatja",
):
    mail = outlook.CreateItem(OUTLOOK_MAIL_ITEM)

    if invoice_path:
        mail.Attachments.Add(invoice_path)

    mail.To = to_email
    mail.Subject = subject
    mail.Body = body
    mail.Categories = category
    mail.Save()  # Save to Drafts
    return mail


def save_emails_with_invoices(persons, invoices_dir, subject, body):
    """Create email drafts in Outlook for each person with their invoice attached."""
    outlook = win32.Dispatch("outlook.application")
    ns = outlook.Session

    for person in persons:
        invoice_path = get_person_invoice(person.apartment, invoices_dir)
        if not invoice_path:
            continue
        for email in person.emails:
            _create_email_draft(outlook, invoice_path, email, subject, body)
            print(f"Created email draft for {person.apartment} to {email}")

    # Open drafts folder in Outlook after creating all drafts
    drafts_folder = ns.GetDefaultFolder(OUTLOOK_FOLDER_DRAFTS)
    drafts_folder.Display()


def persons_with_invoices(persons, invoices_dir):
    """Return only people who have a unique matching invoice PDF."""
    matched_persons, _problems = validate_persons_vs_invoices(persons, invoices_dir)
    return matched_persons


def send_drafts(parent):
    """Send all email drafts in Outlook categorized with 'ArveteSaatja'."""
    outlook = win32.Dispatch("outlook.application")
    ns = outlook.Session

    drafts_folder = ns.GetDefaultFolder(OUTLOOK_FOLDER_DRAFTS)
    messages = drafts_folder.Items
    to_send_ids = []

    # Clear any existing selection
    try:
        explorer = outlook.ActiveExplorer()
        explorer.ClearSelection()
    except Exception:
        pass

    for i in range(1, messages.Count + 1):
        message = messages.Item(i)
        if "ArveteSaatja" in (message.Categories or ""):
            to_send_ids.append((message.EntryID, drafts_folder.StoreID))
    sent_count = 0

    for entry_id, store_id in to_send_ids:
        try:
            message = ns.GetItemFromID(entry_id, store_id)
            message.Send()
            sent_count += 1
        except Exception as e:
            log_exception(e)
            continue

    parent.hide_send_drafts_button()
    return sent_count


def get_person_invoice(person_apartment, invoices_dir):
    apartment = _apartment_key(person_apartment)
    invoice_path = Path(invoices_dir) / f"{apartment}.pdf"
    if invoice_path.exists():
        return str(invoice_path)
    else:
        print(
            f"Warning: No invoice found for apartment {person_apartment} at {invoice_path}"
        )
        return None
