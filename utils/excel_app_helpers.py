import os, re, time
import pythoncom
import win32com.client as win32
import ctypes
from ctypes import wintypes
import subprocess
import gc
import threading

from src.data_classes import Cancelled


def excel_open_workbook(path: str, fn, cancel_event=None, shutdown_timeout=5.0):
    """
    Open Excel + workbook, run function fn(workbook), close workbook and Excel.
    """
    absolute_path = os.path.abspath(path)

    pythoncom.CoInitialize()
    excel_app_instance = workbook = None
    excel_pid = None
    watchdog = None
    try:
        excel_app_instance = excel_app()
        excel_pid = get_excel_pid(excel_app_instance)

        try:
            excel_app_instance.DisplayAlerts = False
            excel_app_instance.AskToUpdateLinks = False
            excel_app_instance.AutomationSecurity = 3 # disable macros
        except Exception:
            pass

        workbook = excel_app_instance.Workbooks.Open(
            absolute_path,
            ReadOnly=True,
            UpdateLinks=0,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )
        if cancel_event is not None and cancel_event.is_set():
            raise Cancelled()

        return fn(excel_app_instance, workbook)
    finally:
        cancelled = (cancel_event is not None and cancel_event.is_set())
        
        if cancelled and excel_pid:
            kill_process(excel_pid)
        
        if excel_pid:
            watchdog = threading.Timer(shutdown_timeout, lambda: kill_process(excel_pid))
            watchdog.daemon = True
            watchdog.start()

        try:
            close_workbook(workbook)
            quit_excel(excel_app_instance)
        finally:
            if watchdog is not None:
                try:
                    watchdog.cancel()
                except Exception:
                    pass

            workbook = None
            excel_app_instance = None
            gc.collect()
            pythoncom.CoUninitialize()


def excel_app():
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AskToUpdateLinks = False
    except Exception:
        pass
    return excel


def close_workbook(wb):
    if wb is not None:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass


def quit_excel(excel):
    if excel is not None:
        try:
            excel.Quit()
        except Exception:
            pass


def get_excel_pid(excel) -> int | None:
    try:
        hwnd = excel.Hwnd
    except Exception:
        return None

    pid = wintypes.DWORD()
    ctypes.windll.user32.GetWindowThreadProcessId(
        wintypes.HWND(hwnd), ctypes.byref(pid)
    )
    return int(pid.value) if pid.value else None


def kill_process(pid: int):
    try:
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(pid)],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        pass
