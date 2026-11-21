import re
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

from config import CLAIM_SHEET_NAME, ATTACH_SHEET_NAME
from core.so_utils import clean_so


REVIEW_COLORS = {"FF0000", "FFFF00", "D9EAF7"}

class ExcelHandler:
    def __init__(self, template_path, output_path=None):
        self.path = Path(template_path).resolve()
        self.output_path = Path(output_path).resolve() if output_path else self.path
        self.wb = None
        self.ws_claim = None
        self.ws_attach = None

    def load(self):
        """Loads the workbook and sheet references."""
        print(f"Loading workbook: {self.path.name}...")
        self.wb = load_workbook(self.path, data_only=False, keep_vba=True)
        self.ws_claim = self.wb[CLAIM_SHEET_NAME]
        self.ws_attach = self.wb[ATTACH_SHEET_NAME]

    def save(self):
        """Saves the workbook."""
        print(f"Saving workbook to: {self.output_path.name}...")
        self.wb.save(self.output_path)

    @staticmethod
    def _extend_data_sheet(worksheet, last_column):
        last_data_row = max(
            (row for row in range(3, worksheet.max_row + 1) if clean_so(worksheet.cell(row, 2).value)),
            default=2,
        )
        styled_rows = [
            row for row, dimension in worksheet.row_dimensions.items()
            if 3 <= row <= last_data_row
            and dimension.height is not None
            and worksheet.cell(row, 2).has_style
            and worksheet.cell(row, 2)._style.borderId
        ]
        if styled_rows:
            unmarked_rows = [
                row for row in styled_rows
                if not (
                    worksheet.cell(row, 2).fill.fgColor.type == "rgb"
                    and str(worksheet.cell(row, 2).fill.fgColor.rgb)[-6:] in REVIEW_COLORS
                )
            ]
            source_row = max(unmarked_rows or styled_rows)
            source_height = worksheet.row_dimensions[source_row].height

            for row in range(source_row + 1, last_data_row + 1):
                if worksheet.row_dimensions[row].height is not None:
                    continue
                worksheet.row_dimensions[row].height = source_height
                for col in range(1, worksheet.max_column + 1):
                    source = worksheet.cell(source_row, col)
                    target = worksheet.cell(row, col)
                    existing_fill = copy(target.fill) if target.fill.patternType else None
                    existing_format = target.number_format if target.number_format != "General" else None
                    target._style = copy(source._style)
                    if existing_fill:
                        target.fill = existing_fill
                    if existing_format:
                        target.number_format = existing_format
                    if col == 1 and source.data_type == "f" and target.value is None:
                        target.value = Translator(source.value, origin=source.coordinate).translate_formula(
                            target.coordinate
                        )

        match = re.search(r"\$[A-Z]+\$(\d+)$", str(worksheet.print_area or ""))
        print_end_row = int(match.group(1)) if match else 2
        if last_data_row > print_end_row:
            worksheet.print_area = f"A1:{last_column}{last_data_row}"

    def extend_data_sheets(self):
        """Carry template row layout and print areas through the last SO."""
        self._extend_data_sheet(self.ws_claim, "R")
        self._extend_data_sheet(self.ws_attach, "H")

    def close(self):
        """Closes the workbook."""
        if self.wb:
            self.wb.close()
