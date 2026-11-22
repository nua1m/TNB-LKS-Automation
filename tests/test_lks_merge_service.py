import unittest
from datetime import datetime, timezone

from core.services.lks_merge_service import MergeRow, _date_value, _runs


class LksMergeServiceTests(unittest.TestCase):
    def test_runs_split_on_source_and_noncontiguous_rows(self):
        rows = [
            MergeRow(0, 3, 3, ()),
            MergeRow(0, 4, 4, ()),
            MergeRow(1, 3, 3, ()),
            MergeRow(1, 5, 5, ()),
        ]
        groups = [(start, len(group)) for start, group in _runs(rows)]
        self.assertEqual(groups, [(0, 2), (2, 1), (3, 1)])

    def test_excel_date_is_timezone_naive(self):
        value = datetime(2026, 11, 22, tzinfo=timezone.utc)
        self.assertEqual(_date_value(value), datetime(2026, 11, 22))


if __name__ == "__main__":
    unittest.main()
