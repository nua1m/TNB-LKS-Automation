import sys
import time
from pathlib import Path
from typing import Callable, Optional

from ui.ascii import show_title
from ui.layout import set_window_size
from ui.components import summary_block, step_progress
from ui.colors import CYAN, GREEN, YELLOW, RED, RESET, DIM

from core.excel_handler import ExcelHandler
from core.so_utils import clean_so
from core.services.claim_service import ClaimService
from core.services.image_injector import ImageInjector
from core.services.quality_control import QualityControl
from core.services.preprocessor import Preprocessor

LogFn = Callable[[str], None]
ConfirmAppendFn = Callable[[int, int], bool]
StatusFn = Callable[[str], None]


def _emit(log_fn: LogFn, message: str = "") -> None:
    log_fn(message)


def run_process(
    data_path,
    template_path,
    log_fn: Optional[LogFn] = None,
    confirm_append_fn: Optional[ConfirmAppendFn] = None,
    status_fn: Optional[StatusFn] = None,
    show_cli_summary: bool = True,
):
    log_fn = log_fn or print
    step_index = 1
    generated_input_path = str(data_path)

    def step(title: str, status_message: str | None = None) -> None:
        nonlocal step_index
        if status_fn and status_message:
            status_fn(status_message)
        _emit(log_fn, f"{CYAN}Step {step_index}: {title}{RESET}")
        step_index += 1

    if data_path.suffix.lower() == ".xls":
        step("Converting legacy .xls file", "Converting legacy .xls file")
        try:
            new_path = Preprocessor.process_legacy_file(data_path)
            generated_input_path = str(new_path)
            _emit(log_fn, f"{GREEN}> Legacy file processed successfully.{RESET}")
            _emit(log_fn, f"{DIM}  New input file: {new_path.name}{RESET}")
            _emit(log_fn, f"{CYAN}> Continuing with the cleaned file...{RESET}")
            return run_process(
                new_path,
                template_path,
                log_fn=log_fn,
                confirm_append_fn=confirm_append_fn,
                status_fn=status_fn,
                show_cli_summary=show_cli_summary,
            )
        except Exception as exc:
            _emit(log_fn, f"{RED}Error processing legacy file: {exc}{RESET}")
            raise

    if status_fn:
        status_fn("Ready to process")

    _emit(log_fn, f"{CYAN}Input file  : {RESET}{data_path}")
    _emit(log_fn, f"{CYAN}Template    : {RESET}{template_path}")

    output_name = f"LKS ({data_path.stem}).xlsm"
    output_path = data_path.parent / output_name
    _emit(log_fn, f"{CYAN}Result file : {RESET}{output_path}")

    start_time = time.time()

    handler = ExcelHandler(template_path, output_path=output_path)
    handler.load()

    source_path = data_path
    source_sheet = None

    step("Reading input data", "Reading input data")
    claim_rows, stats = ClaimService.build_rows(source_path, sheet_name=source_sheet)

    _emit(log_fn, f"{DIM}  - SOs after TRAS removal : {RESET}{GREEN}{stats['sos_after_tras']}{RESET}")
    _emit(log_fn, f"{DIM}  - Duplicate SOs skipped : {RESET}{GREEN}{stats['duplicates_skipped']}{RESET}")
    _emit(log_fn, f"{DIM}  - Rows skipped for TRAS : {RESET}{GREEN}{stats['tras_removed']}{RESET}")
    if stats.get("duplicate_groups"):
        _emit(log_fn, f"{DIM}  - SOs with duplicates    : {RESET}{GREEN}{stats['duplicate_groups']}{RESET}")
        _emit(log_fn, f"{DIM}  - Duplicate SO list:{RESET}")
        duplicate_items = list(stats.get("duplicate_counts", {}).items())
        for so_value, row_count in duplicate_items[:10]:
            _emit(log_fn, f"{DIM}      {so_value}: {RESET}{GREEN}{row_count} rows{RESET}")
        if len(duplicate_items) > 10:
            _emit(log_fn, f"{DIM}      ... and {len(duplicate_items) - 10} more SOs{RESET}")
    if stats.get("tras_by_date"):
        _emit(log_fn, f"{DIM}  - TRAS by date:{RESET}")
        for tras_date, tras_count in stats["tras_by_date"].items():
            _emit(log_fn, f"{DIM}      {tras_date}: {RESET}{GREEN}{tras_count}{RESET}")

    step("Checking existing template data", "Checking existing template data")

    ws_claim = handler.ws_claim
    existing_sos = set()
    for row_index in range(3, ws_claim.max_row + 1):
        so = clean_so(ws_claim.cell(row_index, 2).value)
        if so:
            existing_sos.add(so)

    new_rows = [row for row in claim_rows if clean_so(row["Service Order"]) not in existing_sos]

    if existing_sos:
        should_continue = True
        if confirm_append_fn:
            should_continue = confirm_append_fn(len(existing_sos), len(new_rows))
        else:
            _emit(
                log_fn,
                f"{YELLOW}Template already has {len(existing_sos)} SOs. {len(new_rows)} new SOs will be added. Continue? (y/n){RESET}",
            )
            try:
                should_continue = input(">> ").strip().lower() == "y"
            except EOFError:
                should_continue = True

        if not should_continue:
            _emit(log_fn, f"{RED}Aborted. No changes were saved.{RESET}")
            handler.close()
            return {"aborted": True, "output_path": str(output_path)}

    if not new_rows:
        _emit(log_fn, f"{YELLOW}All SOs already exist in the template. Nothing new was added.{RESET}")
        handler.close()
        return {
            "aborted": False,
            "output_path": str(output_path),
            "new_rows": 0,
            "existing_rows": len(existing_sos),
            "missing_count": 0,
            "counts": {"old": 0, "card": 0, "new": 0},
            "elapsed": time.time() - start_time,
            "generated_input_path": generated_input_path,
        }

    step("Writing rows into the template", "Writing rows into template")

    def get_next_empty(worksheet, col=2):
        for row_index in range(3, worksheet.max_row + 2):
            if worksheet.cell(row_index, col).value in (None, "", " "):
                return row_index
        return worksheet.max_row + 1

    start_claim = get_next_empty(handler.ws_claim)
    start_attach = get_next_empty(handler.ws_attach)
    ClaimService.write_data(handler, new_rows, start_claim, start_attach)

    step("Checking image links", "Checking image links")
    _emit(log_fn, f"{DIM}  Reviewing OLD meter, CARD, and NEW meter image links.{RESET}")

    total_imgs = len(new_rows)
    img_counter = 0

    def img_progress(message):
        nonlocal img_counter
        img_counter += 1
        if status_fn:
            status_fn(f"Checking images ({img_counter}/{total_imgs})")
        if show_cli_summary:
            step_progress("IMAGES", img_counter, total_imgs, extra=message, spinner_i=img_counter)

    ImageInjector.run(handler, source_path, progress_cb=img_progress, sheet_name=source_sheet)
    if show_cli_summary:
        _emit(log_fn, "")

    step("Reviewing rows that need attention", "Reviewing rows that need attention")

    missing, counts = QualityControl.analyze_missing(handler)
    if missing:
        _emit(log_fn, f"{YELLOW}Rows needing review because one or more images are missing:{RESET}")
        max_show = 10
        for index, (so, slots) in enumerate(missing.items()):
            if index < max_show:
                _emit(log_fn, f"  - SO {so} -> {', '.join(slots)}")
        if len(missing) > max_show:
            _emit(log_fn, f"  ... and {len(missing) - max_show} more.")
    else:
        _emit(log_fn, f"{GREEN}All SOs have complete images.{RESET}")

    QualityControl.mark_defective(handler, missing)
    QualityControl.format_all(handler)

    step("Saving the result workbook", "Saving result workbook")
    handler.save()
    handler.close()

    try:
        import win32com.client as win32

        step("Finalizing the workbook", "Finalizing workbook")

        try:
            excel = win32.gencache.EnsureDispatch("Excel.Application")
        except AttributeError:
            excel = win32.Dispatch("Excel.Application")

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        workbook = excel.Workbooks.Open(str(output_path), UpdateLinks=0)
        workbook.ForceFullCalculation = True
        workbook.Save()
        workbook.Close()
        excel.Quit()
        _emit(log_fn, f"{GREEN}Workbook refresh completed.{RESET}")
    except Exception as exc:
        _emit(
            log_fn,
            f"{YELLOW}File saved successfully, but Excel could not auto-refresh ({exc}). If Excel asks, click 'Enable Content'.{RESET}",
        )
        try:
            excel.Quit()
        except Exception:
            pass

    elapsed = time.time() - start_time
    summary = {
        "Processed SOs": stats["sos_after_tras"],
        "Added to template": len(new_rows),
        "Duplicate SOs skipped": stats["duplicates_skipped"],
        "Rows skipped for TRAS": stats["tras_removed"],
        "SOs with duplicates": stats.get("duplicate_groups", 0),
        "Rows needing review": len(missing),
        "Missing OLD meter": counts["old"],
        "Missing CARD": counts["card"],
        "Missing NEW meter": counts["new"],
        "Execution time": f"{elapsed:.2f}s",
    }

    _emit(log_fn, "")
    _emit(log_fn, f"{GREEN}Run complete.{RESET}")
    _emit(log_fn, f"{DIM}Next step: open the saved workbook and review any rows needing attention.{RESET}")

    if show_cli_summary:
        summary_block(summary, str(output_path))

    return {
        "aborted": False,
        "output_path": str(output_path),
        "generated_input_path": generated_input_path,
        "new_rows": len(new_rows),
        "existing_rows": len(existing_sos),
        "missing_count": len(missing),
        "counts": counts,
        "elapsed": elapsed,
        "summary": summary,
        "tras_by_date": stats.get("tras_by_date", {}),
        "duplicates_skipped": stats["duplicates_skipped"],
        "duplicate_groups": stats.get("duplicate_groups", 0),
        "duplicate_counts": stats.get("duplicate_counts", {}),
    }


def main():
    if len(sys.argv) < 2:
        print("Usage: python main.py <data.xlsx> [template.xlsx]")
        sys.exit(1)

    data_path = Path(sys.argv[1]).resolve()

    if len(sys.argv) >= 3:
        template_path = Path(sys.argv[2]).resolve()
    else:
        from config import DEFAULT_TEMPLATE_PATH

        template_path = Path(DEFAULT_TEMPLATE_PATH).resolve()

        if not template_path.exists():
            if Path(DEFAULT_TEMPLATE_PATH).name in [path.name for path in Path.cwd().glob("*")]:
                template_path = Path(Path(DEFAULT_TEMPLATE_PATH).name).resolve()

    if not template_path.exists():
        print(f"{RED}Error: Template file not found: {template_path}{RESET}")
        sys.exit(1)

    set_window_size(110, 40)
    show_title()
    run_process(data_path, template_path, log_fn=print, show_cli_summary=True)


if __name__ == "__main__":
    main()
