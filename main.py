import sys
import time
from pathlib import Path

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

def run_process(data_path, template_path):
    # -----------------------------------------------------
    # LEGACY FILE CHECK (.xls)
    # -----------------------------------------------------
    if data_path.suffix.lower() == ".xls":
        print(f"{CYAN}› Detected Legacy Raw File (.xls). Converting & Cleaning...{RESET}")
        try:
            # Overhaul Workflow: Process & Rewrite
            new_path = Preprocessor.process_legacy_file(data_path)
            
            print(f"{GREEN}› Legacy File Processed successfully.{RESET}")
            print(f"{DIM}  New Input File: {new_path.name}{RESET}")
            print(f"{CYAN}› Relaunching workflow with the clean file...{RESET}\n")
            
            # Recursive Relaunch
            run_process(new_path, template_path)
            return
            
        except Exception as e:
            print(f"{RED}Error processing legacy file: {e}{RESET}")
            import traceback
            traceback.print_exc()
            sys.exit(1)

    # -----------------------------------------------------
    # STANDARD WORKFLOW (.xlsx)
    # -----------------------------------------------------
    # If we are here, data_path is .xlsx (either original or converted)
    
    print(f"{CYAN}RAW       : {RESET}{data_path}")
    print(f"{CYAN}TEMPLATE  : {RESET}{template_path}\n")

    output_name = f"LKS ({data_path.stem}).xlsm"
    output_path = data_path.parent / output_name

    print(f"{CYAN}OUTPUT    : {RESET}{output_path}\n")

    start_time = time.time()

    # -----------------------------------------------------
    # INIT HANDLER
    # -----------------------------------------------------
    handler = ExcelHandler(template_path, output_path=output_path)
    handler.load() 
    
    source_path = data_path
    source_sheet = None # Default first sheet

    # -----------------------------------------------------
    # STEP 1 — PROCESS DATA
    # -----------------------------------------------------
    print(f"{CYAN}› Reading RAW Data...{RESET}")
    # Read from source_path 
    claim_rows, stats = ClaimService.build_rows(source_path, sheet_name=source_sheet)

    print(f"{DIM}  • SOs after TRAS removal     : {RESET}{GREEN}{stats['sos_after_tras']}{RESET}")
    print(f"{DIM}  • Duplicates skipped in RAW  : {RESET}{GREEN}{stats['duplicates_skipped']}{RESET}")
    print(f"{DIM}  • TRAS removed               : {RESET}{GREEN}{stats['tras_removed']}{RESET}\n")

    # -----------------------------------------------------
    # STEP 2 — FILTER NEW ROWS
    # -----------------------------------------------------
    print(f"{CYAN}› Checking existing data...{RESET}")
    wsC = handler.ws_claim
    existing_sos = set()
    for r in range(3, wsC.max_row + 1):
        so = clean_so(wsC.cell(r, 2).value)
        if so: existing_sos.add(so)

    new_rows = [r for r in claim_rows if clean_so(r["Service Order"]) not in existing_sos]

    if existing_sos:
        # Prompt only if potentially interactive? Or auto-append?
        # User requested Append Mode before but reverted.
        # Assuming we stick to "Prompt" logic from main.
        print(f"{YELLOW}The LKS already contains data. Continue adding new SOs? (y/n){RESET}")
        try:
           if input(">> ").strip().lower() != "y":
               print(f"{RED}Aborted.{RESET}")
               handler.close()
               return
        except EOFError:
           pass # Non-interactive
    
    if not new_rows:
        print(f"{YELLOW}All SOs already exist in TEMPLATE.{RESET}")
        handler.close()
        return

    # -----------------------------------------------------
    # STEP 3 — WRITE CLAIM & ATTACHMENT
    # -----------------------------------------------------
    print(f"{CYAN}› Writing CLAIM & ATTACHMENT...{RESET}")
    
    def get_next_empty(ws, col=2):
        for r in range(3, ws.max_row + 2):
            if ws.cell(r, col).value in (None, "", " "): return r
        return ws.max_row + 1

    start_claim = get_next_empty(handler.ws_claim)
    start_attach = get_next_empty(handler.ws_attach)

    ClaimService.write_data(handler, new_rows, start_claim, start_attach)
    print()

    # -----------------------------------------------------
    # STEP 4 — IMAGE PIPELINE
    # -----------------------------------------------------
    print("› Processing IMAGES...\n")
    
    total_imgs = len(new_rows)
    img_counter = 0

    def img_progress(msg):
        nonlocal img_counter
        img_counter += 1
        step_progress("IMAGES", img_counter, total_imgs, extra=msg, spinner_i=img_counter)

    ImageInjector.run(handler, source_path, progress_cb=img_progress, sheet_name=source_sheet)
    print("\n\n")

    # -----------------------------------------------------
    # STEP 5 & 6 — QC (MISSING IMAGES & HIGHLIGHTING)
    # -----------------------------------------------------
    print(f"{RED}› Analyzing & Highlighting defective rows...{RESET}")
    
    missing, counts = QualityControl.analyze_missing(handler)
    
    if missing:
        print(f"{YELLOW}Missing images detected:{RESET}\n")
        max_show = 10
        for i, (so, slots) in enumerate(missing.items()):
            if i < max_show:
                print(f"  • SO {so} → {', '.join(slots)}")
        if len(missing) > max_show:
            print(f"  ... and {len(missing) - max_show} more.")
    else:
        print(f"{GREEN}All SOs have complete images!{RESET}")

    QualityControl.mark_defective(handler, missing)
    QualityControl.format_all(handler)

    print("› Saving Workbook...")
    handler.save()
    handler.close()

    # -----------------------------------------------------
    # STEP 6b — ENABLE EXTERNAL CONTENT (Excel COM)
    # -----------------------------------------------------
    try:
        import win32com.client as win32
        print(f"{DIM}› Activating external content...{RESET}")
        
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
        except AttributeError:
             excel = win32.Dispatch('Excel.Application')

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        
        wb = excel.Workbooks.Open(str(output_path), UpdateLinks=0)
        wb.ForceFullCalculation = True
        wb.Save()
        wb.Close()
        excel.Quit()
        print(f"{GREEN}  External content activated!{RESET}")
    except Exception as e:
        print(f"{YELLOW}  Note: Could not auto-refresh ({e}). Open file in Excel and click 'Enable Content'.{RESET}")
        try: excel.Quit()
        except: pass

    # -----------------------------------------------------
    # STEP 7 — SUMMARY
    # -----------------------------------------------------
    elapsed = time.time() - start_time
    summary_block(
        {
            "Total SOs in RAW": stats["sos_after_tras"],
            "New SOs appended": len(new_rows),
            "Missing OLD meter": counts["old"],
            "Missing CARD": counts["card"],
            "Missing NEW meter": counts["new"],
            "Total defective rows": len(missing),
            "Execution time": f"{elapsed:.2f}s",
        },
        str(output_path),
    )


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
             if Path(DEFAULT_TEMPLATE_PATH).name in [p.name for p in Path.cwd().glob("*")]:
                 template_path = Path(Path(DEFAULT_TEMPLATE_PATH).name).resolve()
                 
    if not template_path.exists():
        print(f"{RED}Error: Template file not found: {template_path}{RESET}")
        sys.exit(1)

    set_window_size(110, 40)
    show_title()
    
    run_process(data_path, template_path)

if __name__ == "__main__":
    main()
