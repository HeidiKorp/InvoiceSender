import os, re, time
import pythoncom
import win32com.client as win32


def excel_open_workbook(path: str, fn):
    """
    Open Excel + workbook, run function fn(workbook), close workbook and Excel.
    """
    pythoncom.CoInitialize()
    excel = wb = None
    try:
        excel = excel_app()
        wb = excel.Workbooks.Open(os.path.abspath(path), ReadOnly=True)
        fn(excel, wb)
    finally:
        close_workbook(wb)
        quit_excel(excel)
        pythoncom.CoUninitialize()


def excel_app():
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
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

        
