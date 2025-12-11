from openpyxl import load_workbook
from pathlib import Path
from config import CLAIM_SHEET_NAME, ATTACH_SHEET_NAME

class ExcelHandler:
    def __init__(self, template_path):
        self.path = Path(template_path).resolve()
        self.wb = None
        self.ws_claim = None
        self.ws_attach = None

    def load(self):
        """Loads the workbook and sheet references."""
        print(f"Loading workbook: {self.path.name}...")
        self.wb = load_workbook(self.path, data_only=False)
        self.ws_claim = self.wb[CLAIM_SHEET_NAME]
        self.ws_attach = self.wb[ATTACH_SHEET_NAME]

    def save(self):
        """Saves the workbook."""
        print(f"Saving workbook...")
        self.wb.save(self.path)

    def close(self):
        """Closes the workbook."""
        if self.wb:
            self.wb.close()
