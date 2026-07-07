"""openpyxl-backed adapter for zh/en paired columns in an Excel sheet.

Column layout (which letters hold zh vs en, which columns form a row key) is entirely
driven by ExcelMapping — nothing here assumes a specific customer layout. Use
`docmgr inspect-excel` to discover the right column letters for a new document.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.workbook import Workbook

from doc_update_tool.config import ExcelMapping
from doc_update_tool.models import Segment, SegmentUpdate


class ExcelAdapter:
    """Loads/saves workbooks in place so styles, merged cells, formulas and unrelated
    sheets are never touched — only the configured en_columns cells are written to."""

    def __init__(self, mapping: ExcelMapping):
        self.mapping = mapping

    def load(self, path: Path) -> Workbook:
        return openpyxl.load_workbook(path)

    def _sheet(self, wb: Workbook):
        if self.mapping.sheet_name == "auto":
            return wb.active
        return wb[self.mapping.sheet_name]

    def extract_segments(self, wb: Workbook) -> list[Segment]:
        ws = self._sheet(wb)
        header_row = self.mapping.header_row
        headers = {col: ws[f"{col}{header_row}"].value for col in self.mapping.zh_columns}

        segments: list[Segment] = []
        for row in range(header_row + 1, ws.max_row + 1):
            zh_values = {col: ws[f"{col}{row}"].value for col in self.mapping.zh_columns}

            row_key = None
            if self.mapping.key_columns:
                key_parts = [str(ws[f"{col}{row}"].value or "").strip() for col in self.mapping.key_columns]
                row_key = "|".join(key_parts) if any(key_parts) else None

            row_has_content = any(
                v is not None and str(v).strip() != "" for v in zh_values.values()
            )
            if not row_has_content and row_key is None:
                continue  # fully blank row, likely trailing whitespace in the sheet

            for zh_col, en_col in zip(self.mapping.zh_columns, self.mapping.en_columns):
                zh_text = zh_values[zh_col]
                en_text = ws[f"{en_col}{row}"].value
                segments.append(
                    Segment(
                        location_id=f"{ws.title}!{zh_col}{row}",
                        en_location_id=f"{ws.title}!{en_col}{row}",
                        row_key=row_key,
                        field_name=headers.get(zh_col),
                        zh_text=str(zh_text) if zh_text is not None else "",
                        en_text=str(en_text) if en_text is not None else "",
                    )
                )
        return segments

    def apply_translations(self, wb: Workbook, updates: list[SegmentUpdate]) -> None:
        for update in updates:
            sheet_name, cell_ref = update.en_location_id.split("!", 1)
            wb[sheet_name][cell_ref] = update.en_text

    def save(self, wb: Workbook, path: Path) -> None:
        wb.save(path)


def inspect(path: Path, sheet_name: str = "auto", header_row: int = 1, sample_rows: int = 3) -> list[dict]:
    """Summarize each column's header + sample values, to help pick zh_columns/en_columns
    for a customer config without knowing the layout up front. Flags formula cells since
    the adapter assumes zh/en columns hold plain text."""
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb.active if sheet_name == "auto" else wb[sheet_name]

    result = []
    for col_idx in range(1, ws.max_column + 1):
        col_letter = ws.cell(row=1, column=col_idx).column_letter
        header = ws.cell(row=header_row, column=col_idx).value
        samples = []
        has_formula = False
        for row in range(header_row + 1, min(header_row + 1 + sample_rows, ws.max_row + 1)):
            cell = ws.cell(row=row, column=col_idx)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                has_formula = True
            samples.append(cell.value)
        result.append(
            {
                "column": col_letter,
                "header": header,
                "samples": samples,
                "has_formula": has_formula,
            }
        )
    return result
