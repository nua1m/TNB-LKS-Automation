"""Merge completed LKS workbooks without turning in-cell pictures back into links."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from openpyxl.utils.datetime import from_excel

from config import DEFAULT_TEMPLATE_PATH
from core.excel_handler import ExcelHandler
from core.services.kpi_service import KpiService
from core.services.summary_service import SummaryService
from core.so_utils import clean_so


CLAIM = "CLAIM"
ATTACHMENT = "ATTACHMENT"
FIRST_ROW = 3
XL_PASTE_ALL = -4104
XL_PASTE_VALUES = -4163


@dataclass(frozen=True)
class MergeRow:
    source: int
    claim_row: int
    attachment_row: int
    claim_values: tuple


def _column_values(sheet, column: str) -> tuple:
    last = sheet.UsedRange.Rows.Count + sheet.UsedRange.Row - 1
    if last < FIRST_ROW:
        return ()
    values = sheet.Range(f"{column}{FIRST_ROW}:{column}{last}").Value2
    if last == FIRST_ROW:
        return (values,)
    return tuple(row[0] for row in values)


def _index_rows(sheet) -> dict[str, int]:
    result = {}
    for row, value in enumerate(_column_values(sheet, "B"), FIRST_ROW):
        so = clean_so(value)
        if so:
            if so in result:
                raise ValueError(f"Duplicate SO {so} within an ATTACHMENT sheet.")
            result[so] = row
    return result


def _select_rows(workbooks) -> tuple[list[MergeRow], int]:
    selected = []
    seen = set()
    duplicates = 0
    for source_index, workbook in enumerate(workbooks):
        claim = workbook.Worksheets(CLAIM)
        attachments = _index_rows(workbook.Worksheets(ATTACHMENT))
        claim_sos = set()
        for row, so_value in enumerate(_column_values(claim, "B"), FIRST_ROW):
            so = clean_so(so_value)
            if not so:
                continue
            claim_sos.add(so)
            if so in seen:
                duplicates += 1
                continue
            if so not in attachments:
                raise ValueError(f"SO {so} has no ATTACHMENT row in input {source_index + 1}.")
            values = claim.Range(f"A{row}:R{row}").Value
            selected.append(MergeRow(source_index, row, attachments[so], tuple(values[0])))
            seen.add(so)
        orphaned = attachments.keys() - claim_sos
        if orphaned:
            raise ValueError(
                f"Input {source_index + 1} has {len(orphaned)} ATTACHMENT SOs missing from CLAIM."
            )
    return selected, duplicates


def _date_value(value):
    if isinstance(value, (int, float)):
        return from_excel(value)
    if isinstance(value, datetime) and value.tzinfo is not None:
        return value.replace(tzinfo=None)
    return value


def _make_skeleton(path: Path, rows: list[MergeRow]) -> None:
    handler = ExcelHandler(DEFAULT_TEMPLATE_PATH, path)
    handler.load()
    try:
        for target_row, item in enumerate(rows, FIRST_ROW):
            values = item.claim_values
            for column in (2, 6, 8, 9, 15, 18):
                value = values[column - 1]
                if column == 9:
                    value = _date_value(value)
                handler.ws_claim.cell(target_row, column).value = value
            handler.ws_attach.cell(target_row, 2).value = values[1]
        handler.extend_data_sheets()
        last_row = FIRST_ROW + len(rows) - 1
        handler.ws_claim.print_area = f"A1:R{last_row}"
        handler.ws_attach.print_area = f"A1:H{last_row}"
        summary = SummaryService.populate(handler, {})
        KpiService.populate(handler, summary["total_row"])
        handler.save()
    finally:
        handler.close()


def _runs(rows: list[MergeRow]):
    start = 0
    for index in range(1, len(rows) + 1):
        if index < len(rows) and index - start < 25 and (
            rows[index].source == rows[index - 1].source
            and rows[index].claim_row == rows[index - 1].claim_row + 1
            and rows[index].attachment_row == rows[index - 1].attachment_row + 1
        ):
            continue
        yield start, rows[start:index]
        start = index


def _paste_block(source_sheet, first: int, last: int, target_sheet, target: int, width: str) -> None:
    source_sheet.Range(f"A{first}:{width}{last}").Copy()
    target_sheet.Range(f"A{target}").PasteSpecial(Paste=XL_PASTE_ALL)
    height = source_sheet.Rows(f"{first}:{last}").RowHeight
    if height is not None:
        target_sheet.Rows(f"{target}:{target + last - first}").RowHeight = height
    else:
        for offset, source_row in enumerate(range(first, last + 1)):
            target_sheet.Rows(target + offset).RowHeight = source_sheet.Rows(source_row).RowHeight


def _fill_output(excel, workbooks, path: Path, rows: list[MergeRow]) -> None:
    output = excel.Workbooks.Open(str(path), UpdateLinks=0)
    try:
        claim = output.Worksheets(CLAIM)
        attachment = output.Worksheets(ATTACHMENT)
        for start, run in _runs(rows):
            target_row = FIRST_ROW + start
            source = workbooks[run[0].source]
            _paste_block(
                source.Worksheets(CLAIM), run[0].claim_row, run[-1].claim_row,
                claim, target_row, "R",
            )
            source_attachment = source.Worksheets(ATTACHMENT)
            _paste_block(
                source_attachment, run[0].attachment_row, run[-1].attachment_row,
                attachment, target_row, "H",
            )
            source_attachment.Range(f"D{run[0].attachment_row}:F{run[-1].attachment_row}").Copy()
            attachment.Range(f"D{target_row}").PasteSpecial(Paste=XL_PASTE_VALUES)
            excel.CutCopyMode = False
        output.ForceFullCalculation = True
        excel.CalculateFull()
        output.Save()
    finally:
        output.Close(SaveChanges=False)


def merge_lks(first: Path, second: Path, output_dir: Path, limit: int = 1000, log=print) -> dict:
    """Keep the first occurrence of each SO and split outputs at ``limit``."""
    if limit < 1:
        raise ValueError("The SO limit must be positive.")
    first, second, output_dir = Path(first).resolve(), Path(second).resolve(), Path(output_dir).resolve()
    if first == second:
        raise ValueError("Select two different LKS files.")
    for path in (first, second):
        if not path.is_file() or path.suffix.lower() != ".xlsm":
            raise ValueError(f"Not an LKS .xlsm file: {path}")
    output_dir.mkdir(parents=True, exist_ok=True)

    import win32com.client as win32

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.AutomationSecurity = 3  # Never execute macros from input workbooks.
    workbooks = []
    try:
        for path in (first, second):
            log(f"Opening {path.name} read-only...")
            workbooks.append(excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True))
        excel.Calculation = -4135  # Avoid recalculating reports after every paste.
        rows, duplicates = _select_rows(workbooks)
        if not rows:
            raise ValueError("Neither LKS file contains CLAIM SOs.")
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        paths = []
        for index, start in enumerate(range(0, len(rows), limit), 1):
            batch = rows[start:start + limit]
            path = output_dir / f"LKS Combined {stamp} Part {index}.xlsm"
            log(f"Creating part {index}: {len(batch)} SOs...")
            _make_skeleton(path, batch)
            _fill_output(excel, workbooks, path, batch)
            paths.append(path)
        return {"paths": paths, "unique_sos": len(rows), "duplicates_skipped": duplicates}
    finally:
        for workbook in reversed(workbooks):
            workbook.Close(SaveChanges=False)
        excel.Quit()
