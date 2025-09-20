import time
import subprocess
import win32com.client as win32
from pathlib import Path
from pywintypes import com_error
import shutil
import winreg

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


def send_emails_with_invoices(persons, invoices_dir, subject, body):
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