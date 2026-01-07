import os, re, time
import pythoncom
import win32com.client as win32


def excel_open_workbook(path: str, fn):
    """
    Open Excel + workbook, run function fn(workbook), close workbook and Excel.
    """
    absolute_path = os.path.abspath(path)

    pythoncom.CoInitialize()
    excel_app_instance = workbook = None
    try:
        excel_app_instance = excel_app()
        workbook = excel_app_instance.Workbooks.Open(absolute_path, ReadOnly=True)
        return fn(excel_app_instance, workbook)
    finally:
        close_workbook(workbook)
        quit_excel(excel_app_instance)
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

        
