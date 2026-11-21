import re
from collections import Counter
from copy import copy
from datetime import date, datetime

from core.services.date_engine import DateEngine
from core.so_utils import clean_so


SUMMARY_SHEET = "Summary (SM INSTALLATION)"
FIRST_DATE_ROW = 6
COUNT_COLUMNS = (
    ("Hari Biasa", "01", False),
    ("Hari Biasa", "02", False),
    ("Hujung Minggu", "01", False),
    ("Hujung Minggu", "02", False),
    ("Cuti Umum", "01", False),
    ("Cuti Umum", "02", False),
    ("Hari Biasa", "01", True),
    ("Hari Biasa", "02", True),
    ("Hujung Minggu", "01", True),
    ("Hujung Minggu", "02", True),
    ("Cuti Umum", "01", True),
    ("Cuti Umum", "02", True),
)


class SummaryService:
    @staticmethod
    def _as_date(value):
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        return DateEngine.parse_date(value)

    @staticmethod
    def _total_row(worksheet):
        for row in range(FIRST_DATE_ROW, worksheet.max_row + 1):
            if str(worksheet.cell(row, 1).value or "").strip().upper() == "JUMLAH":
                return row
        raise ValueError("Could not find the JUMLAH row in the SUMMARY sheet.")

    @staticmethod
    def _extend_dates(worksheet, total_row, extra_rows):
        if not extra_rows:
            return total_row

        moved_merges = []
        for merged in list(worksheet.merged_cells.ranges):
            if merged.min_row >= total_row:
                moved_merges.append(copy(merged))
                worksheet.unmerge_cells(str(merged))

        date_styles = [copy(worksheet.cell(total_row - 1, col)._style) for col in range(1, 17)]
        date_height = worksheet.row_dimensions[total_row - 1].height
        for row in sorted((row for row in worksheet.row_dimensions if row >= total_row), reverse=True):
            dimension = worksheet.row_dimensions.pop(row)
            dimension.index = row + extra_rows
            worksheet.row_dimensions[row + extra_rows] = dimension

        worksheet.insert_rows(total_row, extra_rows)
        for row in range(total_row, total_row + extra_rows):
            worksheet.row_dimensions[row].height = date_height
            for col, style in enumerate(date_styles, 1):
                worksheet.cell(row, col)._style = copy(style)

        for merged in moved_merges:
            merged.shift(row_shift=extra_rows)
            worksheet.merge_cells(str(merged))

        if worksheet.print_area:
            match = re.search(r"\$P\$(\d+)$", str(worksheet.print_area))
            if match:
                worksheet.print_area = f"A1:P{int(match.group(1)) + extra_rows}"
        return total_row + extra_rows

    @staticmethod
    def _count_formula(row, day_type, voltage, is_kiv, claim_end_row):
        day_types = (day_type, "Weekend") if day_type == "Hujung Minggu" else (day_type,)

        def count(label, *remarks_criteria):
            criteria = (
                f"CLAIM!$B$3:$B${claim_end_row},\"<>\","
                f"CLAIM!$I$3:$I${claim_end_row},\">=\"&$A{row},"
                f"CLAIM!$I$3:$I${claim_end_row},\"<\"&($A{row}+1),"
                f"CLAIM!$F$3:$F${claim_end_row},\"{voltage}\","
                f"CLAIM!$O$3:$O${claim_end_row},\"{label}\""
            )
            remarks = "".join(
                f",CLAIM!$R$3:$R${claim_end_row},\"{criterion}\""
                for criterion in remarks_criteria
            )
            return f"COUNTIFS({criteria}{remarks})"

        terms = []
        for label in day_types:
            if is_kiv:
                kiv = count(label, "*KIV*")
                appointment = count(label, "*APPOINTMENT*")
                both = count(label, "*KIV*", "*APPOINTMENT*")
                terms.append(f"({kiv}+{appointment}-{both})")
            else:
                terms.append(count(label, "<>*KIV*", "<>*APPOINTMENT*"))
        return "=" + "+".join(terms)

    @staticmethod
    def populate(handler, tras_by_date):
        worksheet = handler.wb[SUMMARY_SHEET]
        total_row = SummaryService._total_row(worksheet)
        dates = set()
        undated_claims = 0

        for row in range(3, handler.ws_claim.max_row + 1):
            if not clean_so(handler.ws_claim.cell(row, 2).value):
                continue
            claim_date = SummaryService._as_date(handler.ws_claim.cell(row, 9).value)
            if claim_date:
                dates.add(claim_date)
            else:
                undated_claims += 1

        tras_counts = Counter()
        for date_text, count in tras_by_date.items():
            parsed_date = SummaryService._as_date(date_text)
            if parsed_date:
                dates.add(parsed_date)
                tras_counts[parsed_date] += count
            else:
                tras_counts[None] += count

        sorted_dates = sorted(dates)
        if tras_counts[None]:
            sorted_dates.append(None)

        available_rows = total_row - FIRST_DATE_ROW
        total_row = SummaryService._extend_dates(
            worksheet, total_row, max(0, len(sorted_dates) - available_rows)
        )
        claim_end_row = max(handler.ws_claim.max_row, 1002)

        for row in range(FIRST_DATE_ROW, total_row):
            for col in range(1, 16):
                worksheet.cell(row, col).value = None

        for row, value_date in enumerate(sorted_dates, FIRST_DATE_ROW):
            date_cell = worksheet.cell(row, 1)
            date_cell.value = value_date if value_date else "Unknown date"
            if value_date:
                date_cell.number_format = "d-mmm-yy"
                for col, (day_type, voltage, is_kiv) in enumerate(COUNT_COLUMNS, 2):
                    cell = worksheet.cell(row, col)
                    cell.value = SummaryService._count_formula(
                        row, day_type, voltage, is_kiv, claim_end_row
                    )
                    cell.number_format = "0;-0;;@"
            worksheet.cell(row, 14).value = f"=SUM(B{row}:M{row})"
            worksheet.cell(row, 14).number_format = "0;-0;;@"
            worksheet.cell(row, 15).value = tras_counts[value_date] or None

        for col in range(2, 16):
            letter = chr(64 + col)
            worksheet.cell(total_row, col).value = f"=SUM({letter}{FIRST_DATE_ROW}:{letter}{total_row - 1})"

        handler.wb.calculation.fullCalcOnLoad = True
        handler.wb.calculation.forceFullCalc = True
        handler.wb.calculation.calcMode = "auto"

        return {
            "dates": len(sorted_dates),
            "undated_claims": undated_claims,
            "undated_tras": tras_counts[None],
        }
