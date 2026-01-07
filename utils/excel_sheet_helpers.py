import os, re, time
import pythoncom
import win32com.client as win32

from utils.logging_helper import log_exception
from utils.excel_constants import XL_FORMULAS, XL_PART, XL_BY_ROWS, XL_BY_COLUMNS, XL_PREVIOUS


# --- Trailing empty space trimming ---
def set_printarea_to_last_content(sheet):
    """ Set print area to used range, trimming trailing empty rows/columns. """
    row, col = _last_content_row_col(sheet)
    if not row or not col:
        sheet.PageSetup.PrintArea = "$A$1:$A$1"
        return
    sheet.PageSetup.PrintArea = f"$A$1:${col_letter(col)}${row}"


def _last_content_row_col(sheet):
    """ Uses Excel FInd(*) to find last non-empty cell's row and column. """
    try:
        last_row_cell = sheet.Cells.Find("*", sheet.Cells(1, 1), 
                                          LookIn=XL_FORMULAS, LookAt=XL_PART,
                                          SearchOrder=XL_BY_ROWS, SearchDirection=XL_PREVIOUS, MatchCase=False)
        last_col_cell = sheet.Cells.Find("*", sheet.Cells(1, 1), 
                                          LookIn=XL_FORMULAS, LookAt=XL_PART, SearchOrder=XL_BY_COLUMNS,
                                          SearchDirection=XL_PREVIOUS, MatchCase=False)
        if not last_row_cell or not last_col_cell:
            return None, None
        return int(last_row_cell.Row), int(last_col_cell.Column)
    except Exception as e:
        log_exception(e)
        return None, None


def col_letter(col_idx: int) -> str:
    """ Convert 1-based column index to letter(s), e.g. 1 -> A, 27 -> AA. """
    letters = ""
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


# TODO: Move to utils/file_utils.py
# TODO: Make the excel output_dir include period and address
def make_output_dir(base_dir: str, prefix="Arved"):
    """ Create a new output directory with given prefix in base_dir. """
    timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
    out_dir = os.path.join(base_dir, f"{prefix}_{timestamp}")
    os.makedirs(out_dir, exist_ok=True)
    return out_dir


def safe_filename(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r"[<>:\"/\\|?*\x00-\x1F]", "", name) # Windows forbidden chars
    name = re.sub(r"\s+", "_", name) # Replace whitespace with underscore
    return name