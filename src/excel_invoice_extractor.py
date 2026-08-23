from datetime import datetime
import os, re

from utils.excel_app_helpers import excel_open_workbook, close_workbook, quit_excel
from utils.excel_sheet_helpers import set_printarea_to_last_content
from utils.excel_constants import PDF_TYPE, PDF_QUALITY_STANDARD
from utils.logging_helper import log_exception
from src.data_classes import (
    InvoiceItem,
    Cancelled,
    ValidationError,
    ESTONIAN_MONTHS,
    is_valid_address,
    is_valid_period,
    is_valid_year,
)


class ExcelInvoiceExtractor:
    def load(self, invoice_path, on_progress=None, cancel_event=None) -> list[InvoiceItem]:
        def extract_all(_excel, workbook):
            sheet_names = self._korter_sheet_names(workbook)
            if not sheet_names:
                raise ValidationError("Excelis pole lehti nimega 'Korter ...'")

            total = len(sheet_names)
            if on_progress:
                on_progress(0, total)

            meta = self._read_shared_invoice_meta(workbook, sheet_names)

            invoices = []
            for index, sheet_name in enumerate(sheet_names, start=1):
                if cancel_event is not None and cancel_event.is_set():
                    raise Cancelled

                invoices.append(
                    InvoiceItem(
                        apartment=self._extract_apartment(sheet_name),
                        excel_sheet_name=sheet_name,
                        address=meta.get("address"),
                        period=meta.get("period"),
                        year=meta.get("year"),
                    )
                )
                if on_progress:
                    on_progress(index, total)
            return invoices

        return excel_open_workbook(
            invoice_path, extract_all, cancel_event=cancel_event
        )

    def save(self, invoice_batch, on_progress=None):
        cancel_event = invoice_batch.cancel_event
        invoices = invoice_batch.invoices
        total = len(invoices)
        fname = os.path.basename(invoice_batch.invoice_path)

        if on_progress:
            on_progress(0, total, f"Alustan töötlemist...")

        def export_all(_excel, workbook):
            for index, invoice in enumerate(invoices, start=1):
                if cancel_event.is_set():
                    close_workbook(workbook)
                    quit_excel(_excel)
                    raise Cancelled

                sheet_name = invoice.excel_sheet_name
                worksheet = workbook.Sheets(sheet_name)

                set_printarea_to_last_content(worksheet)
                self._remove_forbidden_trailing_rows(
                    worksheet,
                    forbidden_labels=["Radiaator 13", "Radiaator 14"],
                    column_index=1,
                )
                set_printarea_to_last_content(worksheet)

                pdf_path = invoice_batch.dest_dir / f"{invoice.apartment}.pdf"
                worksheet.ExportAsFixedFormat(
                    Type=PDF_TYPE,
                    Filename=str(pdf_path),
                    Quality=PDF_QUALITY_STANDARD,
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False,
                )
                on_progress(index, total, f"Salvestan Exceli lehti {index}/{total} - {fname}")

        return excel_open_workbook(
            invoice_batch.invoice_path, export_all, cancel_event=cancel_event
        )

    def _korter_sheet_names(self, workbook) -> list[str]:
        pattern = re.compile(r"^Korter\s+\d+$", re.IGNORECASE)
        return [ws.Name for ws in workbook.Sheets if pattern.match(str(ws.Name))]

    def _read_shared_invoice_meta(self, workbook, sheet_names: list[str]) -> dict:
        combined = {"period": "", "address": "", "year": ""}
        for sheet_name in sheet_names:
            meta = self._read_invoice_meta(workbook.Sheets(sheet_name))
            if not is_valid_period(combined["period"]) and is_valid_period(meta.get("period")):
                combined["period"] = meta["period"]
            if not is_valid_address(combined["address"]) and is_valid_address(
                meta.get("address")
            ):
                combined["address"] = meta["address"]
            if not is_valid_year(combined["year"]) and is_valid_year(meta.get("year")):
                combined["year"] = meta["year"]
            if (
                is_valid_period(combined["period"])
                and is_valid_address(combined["address"])
                and is_valid_year(combined["year"])
            ):
                break
        return combined

    def _read_invoice_meta(self, sheet, max_rows=50):
        period_text = self._find_right_cell_value(sheet, "Periood", max_rows)
        address_text = self._find_right_cell_value(sheet, "Aadress", max_rows)
        return {
            "period": self._extract_period(period_text),
            "address": self._extract_address(address_text),
            "year": self._extract_year(period_text),
        }

    def _extract_apartment(self, text: str) -> str:
        return text.split(" ")[-1].strip()

    def _extract_address(self, text: str) -> str:
        return text.split(",")[0].strip()

    def _extract_year(self, text: str) -> str:
        year = text.split(".")[-1].strip()
        return year if is_valid_year(year) else ""

    def _extract_period(self, text: str) -> str:
        match = re.search(r"(\d{1,2}\.\d{1,2}\.\d{4})\s*$", text.strip())
        if not match:
            return ""
        parsed_date = datetime.strptime(match.group(1), "%d.%m.%Y")
        return ESTONIAN_MONTHS[parsed_date.month]

    def _find_right_cell_value(self, sheet, label: str, max_rows=50) -> str:
        target_label = self._normalize_label(label)
        for row_idx in range(1, max_rows + 1):
            cell = sheet.Cells(row_idx, 1)
            cell_text = self._normalize_label(self._get_cell_text(cell))
            if cell_text != target_label:
                continue
            value_cell = sheet.Cells(row_idx, 2)
            return self._get_cell_text(value_cell).strip()
        return ""

    def _remove_forbidden_trailing_rows(
        self, worksheet, forbidden_labels: list[str], column_index: int = 1
    ):
        forbidden_norm = {self._normalize_label(label) for label in forbidden_labels}
        last_row = self._get_last_used_row(worksheet)
        while last_row >= 1:
            cell_text = self._get_cell_text(worksheet.Cells(last_row, column_index))
            if self._normalize_label(cell_text) in forbidden_norm:
                worksheet.Rows(last_row).Delete()
                last_row -= 1
                continue
            break

    def _get_last_used_row(self, worksheet) -> int:
        used_range = worksheet.UsedRange
        start_row = used_range.Row
        row_count = used_range.Rows.Count
        return int(start_row + row_count - 1)

    def _get_cell_text(self, cell) -> str:
        try:
            displayed_text = cell.Text
            if displayed_text is not None and str(displayed_text).strip():
                return str(displayed_text)
        except Exception as e:
            log_exception(e)

        try:
            raw_value = cell.Value
            return "" if raw_value is None else str(raw_value)
        except Exception as e:
            log_exception(e)
            return ""

    def _normalize_label(self, label: str) -> str:
        norm = "" if label is None else str(label).strip().lower()
        norm = norm.replace("\xa0", " ")
        if norm.endswith(":"):
            norm = norm[:-1].strip()
        return " ".join(norm.split())
