import tempfile
import unittest
import zipfile
from pathlib import Path

from openpyxl.styles import PatternFill

from core.excel_handler import ExcelHandler


TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "LKS Template (M).xlsm"


class ExcelHandlerTests(unittest.TestCase):
    def setUp(self):
        self.handler = ExcelHandler(TEMPLATE_PATH)
        self.handler.load()

    def tearDown(self):
        self.handler.close()

    def test_extends_both_sheets_past_template_limit(self):
        for row in (1003, 1004):
            self.handler.ws_claim.cell(row, 2).value = str(20000000 + row)
            self.handler.ws_attach.cell(row, 2).value = str(20000000 + row)

        self.handler.extend_data_sheets()

        claim = self.handler.ws_claim
        attachment = self.handler.ws_attach
        for row in (1003, 1004):
            self.assertEqual(claim.row_dimensions[row].height, 30)
            self.assertEqual(attachment.row_dimensions[row].height, 180)
            self.assertEqual(claim.cell(row, 18)._style.borderId, claim.cell(1002, 18)._style.borderId)
            self.assertEqual(attachment.cell(row, 4)._style.borderId, attachment.cell(1002, 4)._style.borderId)
            self.assertEqual(claim.cell(row, 1).value, f"=1+A{row - 1}")
            self.assertEqual(attachment.cell(row, 1).value, f"=1+A{row - 1}")

        self.assertTrue(str(claim.print_area).endswith("$R$1004"))
        self.assertTrue(str(attachment.print_area).endswith("$H$1004"))
        self.assertEqual(claim.sheet_view.view, "pageBreakPreview")

        self.handler.extend_data_sheets()
        self.assertEqual(claim.cell(1004, 1).value, "=1+A1003")

        with tempfile.TemporaryDirectory() as directory:
            self.handler.output_path = Path(directory) / "expanded.xlsm"
            self.handler.save()
            with zipfile.ZipFile(self.handler.output_path) as archive:
                self.assertIn("xl/vbaProject.bin", archive.namelist())

    def test_repairs_old_unstyled_rows_without_erasing_review_fill(self):
        claim = self.handler.ws_claim
        claim.cell(1002, 2).fill = PatternFill(fill_type="solid", fgColor="FFFF00")
        claim.cell(1003, 2).value = "20001003"
        claim.cell(1003, 2).number_format = "@"
        claim.cell(1003, 2).fill = PatternFill(fill_type="solid", fgColor="FF0000")

        self.handler.extend_data_sheets()

        self.assertEqual(claim.row_dimensions[1003].height, 30)
        self.assertEqual(claim.cell(1003, 2).fill.fgColor.rgb[-6:], "FF0000")
        self.assertEqual(claim.cell(1003, 2).number_format, "@")
        self.assertEqual(claim.cell(1003, 1).value, "=1+A1002")
        self.assertEqual(claim.cell(1003, 18)._style.borderId, claim.cell(1001, 18)._style.borderId)
        self.assertTrue(str(claim.print_area).endswith("$R$1003"))


if __name__ == "__main__":
    unittest.main()
