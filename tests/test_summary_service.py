import tempfile
import unittest
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from types import SimpleNamespace

from openpyxl import load_workbook

from core.services.summary_service import SummaryService


TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "LKS Template (M).xlsm"


class SummaryServiceTests(unittest.TestCase):
    def setUp(self):
        self.workbook = load_workbook(TEMPLATE_PATH, keep_vba=True)
        self.handler = SimpleNamespace(wb=self.workbook, ws_claim=self.workbook["CLAIM"])
        self.summary = self.workbook["Summary (SM INSTALLATION)"]

    def tearDown(self):
        self.workbook.close()

    def _claim(self, row, date_value, voltage="01", day_type="Hari Biasa", remarks_2=""):
        claim = self.handler.ws_claim
        claim.cell(row, 2).value = str(1000 + row)
        claim.cell(row, 6).value = voltage
        claim.cell(row, 9).value = date_value
        claim.cell(row, 15).value = day_type
        claim.cell(row, 18).value = remarks_2

    def test_fills_date_rows_and_preserves_formula_updates(self):
        first_day = datetime(2026, 11, 8, 12)
        self._claim(3, first_day, voltage="01", day_type="Hujung Minggu")
        self._claim(4, first_day, voltage="02", day_type="Hujung Minggu", remarks_2="kiv")
        self._claim(5, first_day + timedelta(days=1), remarks_2="appointment")
        self._claim(6, first_day + timedelta(days=1), remarks_2="KIV / APPOINTMENT")

        report = SummaryService.populate(self.handler, {"08 Nov 2026": 2})

        self.assertEqual(report["dates"], 2)
        self.assertEqual(self.summary["A6"].value, first_day.date())
        self.assertEqual(self.summary["A7"].value, (first_day + timedelta(days=1)).date())
        self.assertEqual(self.summary["O6"].value, 2)
        self.assertEqual(self.summary["N6"].value, "=SUM(B6:M6)")
        self.assertIn('"Hujung Minggu"', self.summary["D6"].value)
        self.assertIn('"Weekend"', self.summary["D6"].value)
        self.assertIn('"<>*KIV*"', self.summary["B7"].value)
        self.assertIn('"<>*APPOINTMENT*"', self.summary["B7"].value)
        self.assertIn('"*KIV*"', self.summary["H7"].value)
        self.assertIn('"*APPOINTMENT*"', self.summary["H7"].value)
        self.assertIn("-COUNTIFS(", self.summary["H7"].value)
        self.assertEqual(self.summary["B7"].number_format, "0;-0;;@")
        self.assertEqual(self.summary["B21"].value, "=SUM(B6:B20)")
        self.assertEqual(self.summary["O21"].value, "=SUM(O6:O20)")

    def test_expands_beyond_fifteen_dates_without_overwriting_footer(self):
        first_day = datetime(2026, 11, 1)
        for offset in range(17):
            self._claim(offset + 3, first_day + timedelta(days=offset))

        SummaryService.populate(self.handler, {"17 Nov 2026": 3})

        self.assertEqual(self.summary["A22"].value, datetime(2026, 11, 17).date())
        self.assertEqual(self.summary["A23"].value, "JUMLAH")
        self.assertEqual(self.summary["B23"].value, "=SUM(B6:B22)")
        self.assertEqual(self.summary["O23"].value, "=SUM(O6:O22)")
        self.assertIn("B24:C24", {str(rng) for rng in self.summary.merged_cells.ranges})
        self.assertTrue(str(self.summary.print_area).endswith("$P$46"))

        SummaryService.populate(self.handler, {"17 Nov 2026": 3})
        self.assertEqual(self.summary["A23"].value, "JUMLAH")

        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "summary-test.xlsm"
            self.workbook.save(output)
            with zipfile.ZipFile(output) as archive:
                self.assertIn("xl/vbaProject.bin", archive.namelist())
            reloaded = load_workbook(output, keep_vba=True)
            try:
                self.assertEqual(reloaded["Summary (SM INSTALLATION)"]["A23"].value, "JUMLAH")
            finally:
                reloaded.close()

    def test_reports_unusable_dates_without_losing_tras_count(self):
        self._claim(3, "not a date")

        report = SummaryService.populate(self.handler, {"Unknown date": 4})

        self.assertEqual(report["undated_claims"], 1)
        self.assertEqual(report["undated_tras"], 4)
        self.assertEqual(self.summary["A6"].value, "Unknown date")
        self.assertEqual(self.summary["O6"].value, 4)


if __name__ == "__main__":
    unittest.main()
