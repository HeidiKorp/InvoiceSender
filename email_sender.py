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

    # print(f'Outlook path: {get_outlook_path()}')
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


def get_email_details_and_send():
    # TODO: Get actual email details
    send_email("", "", "", None)


def send_email(to, subject, body, attachment_path):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = "korpheidi@gmail.com"
    mail.Subject = "Test email"
    mail.Body = "Hello, this is a test email from Python!"
    if attachment_path:
        mail.Attachments.Add(attachment_path)
    mail.Display()
    # mail.Send()