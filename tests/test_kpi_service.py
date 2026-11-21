import tempfile
import unittest
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from types import SimpleNamespace

from openpyxl import load_workbook

from core.services.kpi_service import KpiService
from core.services.summary_service import SummaryService


TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "LKS Template (M).xlsm"


class KpiServiceTests(unittest.TestCase):
    def setUp(self):
        self.workbook = load_workbook(TEMPLATE_PATH, keep_vba=True)
        self.handler = SimpleNamespace(wb=self.workbook, ws_claim=self.workbook["CLAIM"])
        self.kpi = self.workbook["KPI"]

    def tearDown(self):
        self.workbook.close()

    def _claim(self, row, team, date_value, remarks_2=""):
        claim = self.handler.ws_claim
        claim.cell(row, 2).value = str(2000 + row)
        claim.cell(row, 6).value = "01"
        claim.cell(row, 8).value = team
        claim.cell(row, 9).value = date_value
        claim.cell(row, 15).value = "Hari Biasa"
        claim.cell(row, 18).value = remarks_2

    def test_adds_all_teams_and_uses_excel_formulas(self):
        first_day = datetime(2026, 11, 8, 9)
        self._claim(3, "KMRT0001", first_day)
        self._claim(4, "ZMRT0001", first_day + timedelta(hours=2), "KIV")
        self._claim(5, "KMRT0001", first_day + timedelta(days=1), "APPOINTMENT")
        self._claim(6, "ZMRT0006", first_day)
        self._claim(7, "ZMRT0007", first_day)
        self._claim(8, "ZMRT0008", first_day)

        summary_report = SummaryService.populate(self.handler, {})
        report = KpiService.populate(self.handler, summary_report["total_row"])

        self.assertEqual(report["teams"], 8)
        self.assertEqual(report["unassigned_claims"], 0)
        self.assertEqual(self.kpi["B4"].value, "KMRT0001")
        self.assertEqual(self.kpi["B9"].value, "KMRT0006")
        self.assertEqual(self.kpi["B11"].value, "KMRT0008")
        self.assertEqual(self.kpi["A11"].value, 8)
        self.assertEqual(self.kpi["C12"].value, "=SUM(C4:C11)")
        self.assertIn('CLAIM!$H$3:$H$1002,$B4', self.kpi["C4"].value)
        self.assertIn('REPLACE($B4,1,1', self.kpi["C4"].value)
        self.assertIn("SUMPRODUCT", self.kpi["D4"].value)
        self.assertIn("'Summary (SM INSTALLATION)'!$A$6:$A$20", self.kpi["D4"].value)
        self.assertEqual(self.kpi["E4"].value, '=IF(D4=0,"",C4/D4)')
        self.assertEqual(self.kpi["E4"].number_format, "0.00")
        self.assertTrue(str(self.kpi.print_area).endswith("$F$12"))

        KpiService.populate(self.handler, summary_report["total_row"])
        self.assertEqual(self.kpi["C12"].value, "=SUM(C4:C11)")

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "kpi-test.xlsm"
            self.workbook.save(output)
            with zipfile.ZipFile(output) as archive:
                self.assertIn("xl/vbaProject.bin", archive.namelist())

    def test_preserves_zmrt_display_style(self):
        for row in range(4, 9):
            self.kpi.cell(row, 2).value = f"ZMRT{row - 3:04d}"
        self._claim(3, "KMRT0006", datetime(2026, 11, 8))

        summary_report = SummaryService.populate(self.handler, {})
        KpiService.populate(self.handler, summary_report["total_row"])

        self.assertEqual(self.kpi["B9"].value, "ZMRT0006")
        self.assertEqual(self.kpi["C10"].value, "=SUM(C4:C9)")


if __name__ == "__main__":
    unittest.main()
