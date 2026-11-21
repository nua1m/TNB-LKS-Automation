import re
from copy import copy

from core.so_utils import clean_so


KPI_SHEET = "KPI"
FIRST_TEAM_ROW = 4
TEAM_PATTERN = re.compile(r"([KZ])MRT(\d{4})", re.IGNORECASE)


class KpiService:
    @staticmethod
    def _team_suffix(value):
        match = TEAM_PATTERN.fullmatch(str(value or "").strip())
        return match.group(2) if match else None

    @staticmethod
    def _total_row(worksheet):
        for row in range(FIRST_TEAM_ROW, worksheet.max_row + 1):
            cell = worksheet.cell(row, 3)
            if not worksheet.cell(row, 2).value and cell.data_type == "f" and "SUM(" in cell.value.upper():
                return row
        raise ValueError("Could not find the total row in the KPI sheet.")

    @staticmethod
    def _extend_teams(worksheet, total_row, extra_rows):
        if not extra_rows:
            return total_row

        team_styles = [copy(worksheet.cell(total_row - 1, col)._style) for col in range(1, 7)]
        team_height = worksheet.row_dimensions[total_row - 1].height
        for row in sorted((row for row in worksheet.row_dimensions if row >= total_row), reverse=True):
            dimension = worksheet.row_dimensions.pop(row)
            dimension.index = row + extra_rows
            worksheet.row_dimensions[row + extra_rows] = dimension

        worksheet.insert_rows(total_row, extra_rows)
        for row in range(total_row, total_row + extra_rows):
            worksheet.row_dimensions[row].height = team_height
            for col, style in enumerate(team_styles, 1):
                worksheet.cell(row, col)._style = copy(style)

        if worksheet.print_area:
            match = re.search(r"\$F\$(\d+)$", str(worksheet.print_area))
            if match:
                worksheet.print_area = f"A1:F{int(match.group(1)) + extra_rows}"
        return total_row + extra_rows

    @staticmethod
    def _other_prefix(row):
        return f'REPLACE($B{row},1,1,IF(LEFT($B{row},1)="K","Z","K"))'

    @staticmethod
    def _installed_formula(row, claim_end_row):
        so_range = f"CLAIM!$B$3:$B${claim_end_row}"
        team_range = f"CLAIM!$H$3:$H${claim_end_row}"
        return (
            f'=COUNTIFS({so_range},"<>",{team_range},$B{row})'
            f'+COUNTIFS({so_range},"<>",{team_range},{KpiService._other_prefix(row)})'
        )

    @staticmethod
    def _working_days_formula(row, claim_end_row, summary_total_row):
        dates = f"'Summary (SM INSTALLATION)'!$A$6:$A${summary_total_row - 1}"
        so_range = f"CLAIM!$B$3:$B${claim_end_row}"
        team_range = f"CLAIM!$H$3:$H${claim_end_row}"
        status_range = f"CLAIM!$I$3:$I${claim_end_row}"

        def daily_count(team):
            return (
                f'COUNTIFS({so_range},"<>",{team_range},{team},'
                f'{status_range},">="&{dates},'
                f'{status_range},"<"&({dates}+1))'
            )

        return (
            f"=SUMPRODUCT(--ISNUMBER({dates}),"
            f"--(({daily_count(f'$B{row}')}+{daily_count(KpiService._other_prefix(row))})>0))"
        )

    @staticmethod
    def populate(handler, summary_total_row):
        worksheet = handler.wb[KPI_SHEET]
        total_row = KpiService._total_row(worksheet)
        existing_ids = [worksheet.cell(row, 2).value for row in range(FIRST_TEAM_ROW, total_row)]
        first_id = next((value for value in existing_ids if KpiService._team_suffix(value)), "KMRT0001")
        prefix = str(first_id).strip().upper()[0]
        suffixes = {suffix for value in existing_ids if (suffix := KpiService._team_suffix(value))}
        unassigned_claims = 0

        for row in range(3, handler.ws_claim.max_row + 1):
            if not clean_so(handler.ws_claim.cell(row, 2).value):
                continue
            suffix = KpiService._team_suffix(handler.ws_claim.cell(row, 8).value)
            if suffix:
                suffixes.add(suffix)
            else:
                unassigned_claims += 1

        sorted_suffixes = sorted(suffixes, key=int)
        total_row = KpiService._extend_teams(
            worksheet, total_row, max(0, len(sorted_suffixes) - (total_row - FIRST_TEAM_ROW))
        )
        claim_end_row = max(handler.ws_claim.max_row, 1002)

        for row in range(FIRST_TEAM_ROW, total_row):
            for col in range(1, 6):
                worksheet.cell(row, col).value = None

        for row, suffix in enumerate(sorted_suffixes, FIRST_TEAM_ROW):
            worksheet.cell(row, 1).value = row - FIRST_TEAM_ROW + 1
            worksheet.cell(row, 2).value = f"{prefix}MRT{suffix}"
            worksheet.cell(row, 3).value = KpiService._installed_formula(row, claim_end_row)
            worksheet.cell(row, 4).value = KpiService._working_days_formula(
                row, claim_end_row, summary_total_row
            )
            worksheet.cell(row, 5).value = f'=IF(D{row}=0,"",C{row}/D{row})'
            worksheet.cell(row, 3).number_format = "0;-0;;@"
            worksheet.cell(row, 4).number_format = "0;-0;;@"
            worksheet.cell(row, 5).number_format = "0.00"

        worksheet.cell(total_row, 3).value = f"=SUM(C{FIRST_TEAM_ROW}:C{total_row - 1})"
        handler.wb.calculation.fullCalcOnLoad = True
        handler.wb.calculation.forceFullCalc = True
        handler.wb.calculation.calcMode = "auto"

        return {"teams": len(sorted_suffixes), "unassigned_claims": unassigned_claims}
