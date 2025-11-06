import time
import subprocess
import win32com.client as win32
from pathlib import Path
from pywintypes import com_error
import shutil, os
import winreg
from xls_extractor import ValidationError
from collections import Counter

def get_outlook_path():
    try:
        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
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
            print(f"Could not clear Outlook cache at {gen_py}. Error: {e}")

    # Rebuild cache
    try:
        win32.gencache.Rebuild()
    except Exception as e:
        print(f"Could not rebuild Outlook cache. Error: {e}")


def ensure_outlook_ready(timeout=120):
    """
    Ensure that Outlook is running and ready to send emails. If Outlook is not configured, this will
    show the Outlook logon/profile dialog so you can finish setup.
    """
    start = time.time()
    app = None

    try:
        subprocess.Popen([get_outlook_path()], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        print(f'Could not start Outlook. Please ensure it is installed. Error: {e}')
        pass

    while time.time() - start < timeout:
        try:
            app = win32.gencache.EnsureDispatch('Outlook.Application')
            session = app.GetNamespace("MAPI")
            # profile, password, showDoalog, newSession
            session.Logon("", "", True, True)  # This will prompt for profile if not configured
            _ = session.Accounts  # Accessing Accounts to ensure it's fully loaded
            return app
        except com_error:
            time.sleep(2)
        except Exception:
            time.sleep(2)
    raise RuntimeError("Outlook did not become ready in time. "
                        "Open Outlook manually, finish the setup wizard"
                        "then rerun the script.")


def apartments_from_persons(persons):
    return {str(p.apartment).strip() for p in persons if str(p.apartment).strip()}


def apartments_from_invoices(invoices_dir, exts={".pdf"}):
    counts = Counter()
    for p in Path(invoices_dir).iterdir():
        if p.is_file() and p.suffix.lower() in exts:
            apt = p.stem.strip()
            if apt:
                counts[apt] += 1
    return counts


def validate_persons_vs_invoices(persons, invoices_dir):
    person_apts = apartments_from_persons(persons)
    invoice_counts = apartments_from_invoices(invoices_dir)
    invoice_apts = set(invoice_counts.keys())

    # Who's missing an invoice?
    missing_for_people = sorted(person_apts - invoice_apts, key=str)

    # Invoices that don't match any pattern
    extra_invoices = sorted(invoice_apts - person_apts, key=str)

    # Duplicates (same apartment has >1 file)
    duplicates = sorted([apt for apt, c in invoice_counts.items() if c > 1], key=str)

    problems = []
    if missing_for_people:
        problems.append(f"Puuduvad arved korteritele: {', '.join(missing_for_people)}.")
    if extra_invoices:
        problems.append(f"Arved, millele ei leitud klienti: {', '.join(extra_invoices)}.")
        print("Invoice apartments:", {', '.join(invoice_apts)})
    if duplicates:
        problems.append(f"Duplikaatsed arvefailid korteritele: {', '.join(duplicates)}.")

    print(f'Problems: {problems}')
    if problems:
        raise ValidationError(" ".join(problems))


def send_emails_with_invoices(persons, invoices_dir, subject, body):
    # Validate *before* creating drafts
    # validate_persons_vs_invoices(persons, invoices_dir)

    print(f'Getting past error')
    
    olMailItem = 0
    olFolderDrafts = 16

    outlook = win32.Dispatch('outlook.application')
    ns = outlook.Session
    drafts_folder = ns.GetDefaultFolder(olFolderDrafts)

    for person in persons:
        for i in range(len(person.emails)):
            mail = outlook.CreateItem(olMailItem)

            invoice_path = get_person_invoice(person.apartment, invoices_dir)
            if invoice_path:
                mail.Attachments.Add(invoice_path)

            if not invoice_path:
                # Should not happen now, but guard anyway
                raise ValidationError(f"Arvet ei leitud korterile: {person.apartment}")

            mail.To = person.emails[i]  # Send to the first valid email
            mail.Subject = subject # maybe period is needed here
            mail.Body = body
            mail.Save() # Save to Drafts
    
    drafts_folder.Display()



def get_person_invoice(person_apartment, invoices_dir):
    invoice_path = invoices_dir / f'{person_apartment}.pdf'
    if invoice_path.exists():
        return str(invoice_path)
    else:
        print(f"Warning: No invoice found for apartment {person_apartment} at {invoice_path}")
        return None


# TODO: write a function for sending all emails in drafts folder