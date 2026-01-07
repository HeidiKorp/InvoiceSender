import os, re, time
import pythoncom
import win32com.client as win32
from datetime import datetime
from pathlib import Path

from utils.logging_helper import log_exception
from utils.excel_app_helpers import excel_open_workbook
from utils.excel_sheet_helpers import set_printarea_to_last_content, make_output_dir, safe_filename, col_letter
from utils.excel_constants import (XL_FORMULAS, XL_PART, XL_BY_ROWS, XL_BY_COLUMNS, XL_PREVIOUS, XL_NEXT, XL_VALUES,
                                   PDF_TYPE, PDF_QUALITY_STANDARD)
from src.data_classes import InvoiceItem
from utils.file_utils import create_invoice_dir

ESTONIAN_MONTHS = {
    1: "jaanuar", 2: "veebruar", 3: "märts", 4: "aprill",
    5: "mai", 6: "juuni", 7: "juuli", 8: "august",
    9: "september", 10: "oktoober", 11: "november", 12: "detsember",
}

def save_excel_invoices_as_pdfs(invoice_batch: "InvoiceBatch", on_progress=None) -> Path:
    parent = invoice_batch.parent
    cancel_event = invoice_batch.cancel_event
    invoices = invoice_batch.invoices

    # invoice_dir = create_invoice_dir(invoice_batch.dest_dir, invoices[0])

    total = len(invoices)
    fname = os.path.basename(invoice_batch.invoice_path)
    
    if on_progress:
        on_progress(0, total, f"Alustan töötlemist...")

    def export_all(_excel, workbook):
        for index, invoice in enumerate(invoices, start=1):
            if cancel_event.is_set():
                log_exception(KeyboardInterrupt("Kasutaja katkestas töö"))
                break

            sheet_name = invoice.excel_sheet_name
            worksheet = workbook.Sheets(sheet_name)

            set_printarea_to_last_content(worksheet)

            pdf_path = invoice_batch.dest_dir / f"{invoice.apartment}.pdf"

            worksheet.ExportAsFixedFormat(
                Type=PDF_TYPE,  # PDF
                Filename=str(pdf_path),
                Quality=PDF_QUALITY_STANDARD,  # Standard
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
            on_progress(index, total, f"Salvestan Exceli lehti {index}/{total} - {fname}")

            
        # return invoice_dir
    return excel_open_workbook(invoice_batch.invoice_path, export_all)


def create_excel_invoices(sheets: list, meta: dict) -> list[InvoiceItem]:
    """ Create ExcelInvoice objects from given sheets and shared metadata. """
    invoices = []
    for sheet in sheets:
        invoice = InvoiceItem(
            excel_sheet_name=sheet,
            address=meta.get("address", ""),
            period=meta.get("period", ""),
            apartment=meta.get("apartment", ""),
            year=meta.get("year", ""),
        )
        invoices.append(invoice)
    return invoices


# --- Sheet selection and export ---

def get_korter_sheet_names(wb) -> list[str]:
    """ Return list of sheets named "Korter X" where X is a number. """
    pattern = re.compile(r"^Korter\s+\d+$", re.IGNORECASE)
    return [ws.Name for ws in wb.Sheets if pattern.match(str(ws.Name))]


def _export_sheet_to_pdf(sheet, output_dir: str):
    pdf_path = os.path.join(output_dir, f"{_safe_filename(sheet.Name)}.pdf")
    sheet.ExportAsFixedFormat(
        Type=PDF_TYPE,  # PDF
        Filename=pdf_path,
        Quality=PDF_QUALITY_STANDARD,  # Standard
        IncludeDocProperties=True,
        IgnorePrintAreas=False,
        OpenAfterPublish=False,
    )


# --- Metadata extraction ---
def read_invoice_meta_col_a(sheet, max_rows=50):
    """ Read invoice metadata from given sheet. """
    period_text = _find_right_cell_value(sheet, "Periood", max_rows)
    address_text = _find_right_cell_value(sheet, "Aadress", max_rows)
    print(f'Address text: "{address_text}"')
    print(f'Period text: "{period_text}"')
    return {
        "period": _extract_period(period_text),
        "address": _extract_address(address_text),
        "year": _extract_year(period_text),
    }

def _extract_address(text: str) -> str:
    """ Extract address from text by removing apartment number if present. """
    return text.split(",")[0].strip()


def extract_apartment(text: str) -> str:
    """ Extract apartment number from text, if present. """
    return text.split(" ")[-1].strip()


def _extract_year(text: str) -> str:
    """ Extract year from text. """
    return text.split(".")[-1].strip()


def _extract_period(text: str) -> str:
    """ Extract period from text. """
    match = re.search(r"(\d{1,2}\.\d{1,2}\.\d{4})\s*$", text.strip())
    if not match:
        return ""
    parsed_date = datetime.strptime(match.group(1), "%d.%m.%Y")
    return ESTONIAN_MONTHS[parsed_date.month]


def _find_right_cell_value(sheet, label: str, max_rows=50) -> str:
    """ Find cell with given label and return value of the cell to its right. """
    target_label = normalize_label(label)

    for row_idx in range(1, max_rows + 1):
        cell = sheet.Cells(row_idx, 1)  # Column A
        cell_text = normalize_label(get_cell_text(cell))
        if cell_text != target_label:
            continue

        value_cell = sheet.Cells(row_idx, 2)  # Column B
        return get_cell_text(value_cell).strip()
    return ""


def get_cell_text(cell) -> str:
    """ Safely get cell text, handling merged cells. """
    try:
        displayed_text = cell.Text
        if displayed_text is not None and str(displayed_text).strip():
            return str(displayed_text)
    except Exception as e:
        log_exception(e)
        pass

    try:
        raw_value = cell.Value
        return "" if raw_value is None else str(raw_value)
    except Exception as e:
        log_exception(e)
        return ""


def normalize_label(label: str) -> str:
    """ Normalize label for comparison: lowercase, strip whitespace and trailing colon. """
    norm = "" if label is None else str(label).strip().lower()
    norm = norm[:1].strip() if norm.endswith(":") else norm
    return " ".join(norm.split())


def debug_print_range(ws, nrows=40, ncols=8, start_row=1, start_col=1):
    """
    Prints a rectangular block from the worksheet.
    Good for seeing what values Excel COM actually returns.
    """
    rng = ws.Range(ws.Cells(start_row, start_col), ws.Cells(start_row + nrows - 1, start_col + ncols - 1))
    values = rng.Value  # tuple of tuples (rows)

    # Header line with column letters
    col_letters = [col_letter(start_col + i) for i in range(ncols)]
    print("      " + " | ".join(f"{c:>12}" for c in col_letters))
    print("      " + "-" * (15 * ncols))

    for r_idx, row in enumerate(values, start=start_row):
        cells = []
        for v in row:
            s = "" if v is None else str(v)
            s = s.replace("\n", " ").strip()
            if len(s) > 12:
                s = s[:11] + "…"
            cells.append(f"{s:>12}")
        print(f"{r_idx:>4}: " + " | ".join(cells))