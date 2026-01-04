import os, re, time
import pythoncom
import win32com.client as win32

from utils.logging_helper import log_exception
from utils.excel_app_helpers import excel_open_workbook
from utils.excel_sheet_helpers import set_printarea_to_last_content, make_output_dir, safe_filename


# --- Public entrypoint ---
def export_excel_to_pdfs(parent, excel_path: str, output_dir: str, cancel_event) -> str:
    """
    Worker thread function.
    Exports sheets named "Korter X" to PDFs, preserving formatting/images.
    Returns output folder path.
    """
    out_dir = make_output_dir(output_dir, prefix="Arved_Excel")
    _ui(parent, lambda: _status(parent, "Alustan Excel -> PDF...", 0, 1))

    def work(excel, wb):
        sheets = _get_korter_sheets(wb)
        if not sheets:
            raise ValueError("Excelis pole lehti nimega 'Korter ...'")
        
        total = len(sheets)
        for idx, sheet in enumerate(sheets, start=1):
            if cancel_event.is_set():
                log_exception(KeyboardInterrupt("Kasutaja katkestas töö"))
                return

            set_printarea_to_last_content(sheet) # trims only trailing space
            pdf_path = os.path.join(out_dir, f"{safe_filename(sheet.Name)}.pdf")
            _export_sheet_to_pdf(sheet, out_dir)

            parent.after(
                0,
                lambda: _status(
                    parent,
                    f"Ekspordin lehte '{sheet.Name}' PDF-iks... ({idx}/{total})",
                    idx - 1,
                    total,
                ),
            )
    excel_open_workbook(excel_path, work)
    return out_dir



# --- Sheet selection and export ---

def _get_korter_sheets(wb):
    """ Return list of sheets named "Korter X" where X is a number. """
    pattern = re.compile(r"^Korter\s+\d+$", re.IGNORECASE)
    return [ws for ws in wb.Sheets if pattern.match(ws.Name)]


def _export_sheet_to_pdf(sheet, output_dir: str):
    pdf_path = os.path.join(output_dir, f"{_safe_filename(sheet.Name)}.pdf")
    sheet.ExportAsFixedFormat(
        Type=0,  # PDF
        Filename=pdf_path,
        Quality=0,  # Standard
        IncludeDocProperties=True,
        IgnorePrintAreas=False,
        OpenAfterPublish=False,
    )


# --- UI helpers ---
def _status(parent, text: str, progress: int, total: int):
    pct = int((progress / total) * 100) if total else 0
    parent.status_bar.pack(fill="X", side="bottom")
    parent.page_progress.configure(mode="determinate", value=pct)
    parent.status_label.configure(text=text)
    parent.update_idletasks()