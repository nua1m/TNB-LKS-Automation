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
from core.services.summary_service import SummaryService

def run_process(data_path, target_path=None):
    """
    Core automation logic. 
    data_path: Path to raw data (str or Path)
    target_path: Path to existing master file (str or Path), optional.
    """
    data_path = Path(data_path).resolve()
    
    # OUTPUT NAME LOGIC
    if target_path:
        # User specified target file (e.g. "LKS Jan.xlsm")
        output_path = Path(target_path).resolve()
    else:
        # Default: "LKS ({DataName}).xlsm" in same folder as Data
        output_name = f"LKS ({data_path.stem}).xlsm"
        output_path = data_path.parent / output_name

    # TEMPLATE LOGIC
    # Always try default template first since we rarely change it via CLI now
    from config import DEFAULT_TEMPLATE_PATH
    template_path = Path(DEFAULT_TEMPLATE_PATH).resolve()
    
    if not template_path.exists():
         # Fallback: check if it's in the current dir directly
         if Path(DEFAULT_TEMPLATE_PATH).name in [p.name for p in Path.cwd().glob("*")]:
             template_path = Path(Path(DEFAULT_TEMPLATE_PATH).name).resolve()
             
    if not template_path.exists():
        print(f"{RED}Error: Template file not found: {template_path}{RESET}")
        return # Return instead of exit for GUI

    set_window_size(110, 40)
    show_title()

    print(f"{CYAN}RAW       : {RESET}{data_path}")
    if output_path.exists():
        print(f"{CYAN}TARGET    : {RESET}{output_path} (Append Mode)")
    else:
        print(f"{CYAN}TARGET    : {RESET}{output_path} (New File)")
    print(f"{CYAN}TEMPLATE  : {RESET}{template_path}\n")

    start_time = time.time()

    # -----------------------------------------------------
    # INIT HANDLER
    # -----------------------------------------------------
    # Logic: If output_path exists, we LOAD it (Append Mode).
    # If not, we load TEMPLATE and save to output_path (Create Mode).
    
    append_mode = False
    
    if output_path.exists():
        print(f"{YELLOW}› Output file exists. Switching to APPEND MODE.{RESET}")
        print(f"{DIM}  Loading: {output_path.name}{RESET}")
        handler = ExcelHandler(output_path, output_path=output_path)
        handler.load()
        append_mode = True
    else:
        print(f"{GREEN}› Creating NEW Report.{RESET}")
        handler = ExcelHandler(template_path, output_path=output_path)
        handler.load() # Opens Workbook from template_path
    
    # -----------------------------------------------------
    # PREPROCESS RAW DATA (IF .XLS)
    # -----------------------------------------------------
    source_path = data_path
    source_sheet = None
    
    # Check for legacy excel extension (.xls)
    if data_path.suffix.lower() == ".xls":
        print(f"{CYAN}› Detected Legacy Raw File (.xls). cleaning...{RESET}")
        try:
            clean_df = Preprocessor.clean_raw_data(data_path)
            raw_sheet_name = Preprocessor.insert_clean_data(handler, clean_df)
            
            # Save intermediate to allow Pandas to read the new sheet
            print(f"{DIM}  Saving intermediate cleaned data...{RESET}")
            handler.save() 
            
            # RELOAD HANDLER to avoid "I/O operation on closed file" error with images
            # The first save might close image file handles from the template.
            # Since we just saved to output_path, we load from there now.
            handler = ExcelHandler(output_path)
            handler.load()
            
            # Update pointers
            source_path = output_path
            source_sheet = raw_sheet_name
            print(f"{GREEN}  Cleaned data saved to sheet '{raw_sheet_name}' in output.{RESET}\n")
            
        except ImportError as e:
            print(f"{RED}Error: {e}{RESET}")
            return
        except Exception as e:
             print(f"{RED}Preprocessor Error: {e}{RESET}")
             return

    # -----------------------------------------------------
    # STEP 1 — PROCESS DATA
    # -----------------------------------------------------
    print(f"{CYAN}› Reading RAW Data...{RESET}")
    # Read from source_path (either original or the intermediate cleaned one)
    claim_rows, tras_rows, stats = ClaimService.build_rows(source_path, sheet_name=source_sheet)

    print(f"{DIM}  • SOs after TRAS removal     : {RESET}{GREEN}{stats['sos_after_tras']}{RESET}")
    print(f"{DIM}  • Duplicates skipped in RAW  : {RESET}{GREEN}{stats['duplicates_skipped']}{RESET}")
    print(f"{DIM}  • TRAS removed               : {RESET}{GREEN}{stats['tras_removed']}{RESET}\n")

    # EXPORT TRAS (If any)
    # Define Helper early
    def get_next_empty(ws, col=2):
        for r in range(3, ws.max_row + 2):
            if ws.cell(r, col).value in (None, "", " "): return r
        return ws.max_row + 1

    # EXPORT TRAS (If any)
    if tras_rows:
        tras_output_name = f"TRAS ({data_path.stem}).xlsm"
        tras_output_path = data_path.parent / tras_output_name
        print(f"{CYAN}› Exporting TRAS Report...{RESET}")
        
        # TRAS Handler (Append if exists)
        if tras_output_path.exists():
            print(f"{DIM}  Appended to existing: {tras_output_path.name}{RESET}")
            tras_handler = ExcelHandler(tras_output_path, output_path=tras_output_path)
            tras_handler.load()
            tras_start_claim = get_next_empty(tras_handler.ws_claim)
            tras_start_attach = get_next_empty(tras_handler.ws_attach)
            
            # Filter duplicates for TRAS?
            # Assuming we want to avoid duplicates same as main
            t_existing = set()
            for r in range(3, tras_handler.ws_claim.max_row + 1):
                so = clean_so(tras_handler.ws_claim.cell(r, 2).value)
                if so: t_existing.add(so)
            tras_rows_final = [r for r in tras_rows if clean_so(r["Service Order"]) not in t_existing]
        else:
            print(f"{DIM}  Creating New: {tras_output_path.name}{RESET}")
            tras_handler = ExcelHandler(template_path, output_path=tras_output_path)
            tras_handler.load()
            tras_start_claim = 3
            tras_start_attach = 3
            tras_rows_final = tras_rows

        if tras_rows_final:
            ClaimService.write_data(tras_handler, tras_rows_final, tras_start_claim, tras_start_attach)
            
            # Inject Images (Same Source)
            print(f"{DIM}  Injecting Images for TRAS...{RESET}")
            ImageInjector.run(tras_handler, source_path, sheet_name=source_sheet)
            
            tras_handler.save()
            tras_handler.close()
            print(f"{GREEN}  TRAS Report saved!{RESET}\n")
        else:
            print(f"{YELLOW}  TRAS Data duplicated/empty. Skipped save.{RESET}\n")
            if tras_output_path.exists(): tras_handler.close()


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
        # For CLI we might prompt but for library usage we skip prompt if seemingly automated or append mode valid
        pass
    
    if not new_rows:
        print(f"{YELLOW}All SOs already exist in TEMPLATE.{RESET}")
        handler.close()
        return

    # -----------------------------------------------------
    # STEP 3 — WRITE CLAIM & ATTACHMENT
    # -----------------------------------------------------
    print(f"{CYAN}› Writing CLAIM & ATTACHMENT...{RESET}")
    
    # Determine start rows


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
        # Use counter % total or simple increment logic depending on how many calls are made
        # ImageInjector calls this once per row
        step_progress("IMAGES", img_counter, total_imgs, extra=msg, spinner_i=img_counter)

    # Note: ImageInjector currently processes the 'new_rows' theoretically, 
    # but the implementation scans the WHOLE sheet range for safety/robustness.
    # To match old behavior, pass callback.
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
    
    # -----------------------------------------------------
    # STEP 6a — UPDATE SUMMARY SHEET
    # -----------------------------------------------------
    # We update the summary based on the FINAL state of Claim sheet
    print(f"{CYAN}› Updating Summary Sheet...{RESET}")
    SummaryService.update_summary(handler)

    print("› Saving Workbook...")
    handler.save()
    handler.close()

    # -----------------------------------------------------
    # STEP 6b — ENABLE EXTERNAL CONTENT (Excel COM)
    # This opens the file in Excel and refreshes to activate IMAGE formulas
    # -----------------------------------------------------
    try:
        import win32com.client as win32
        print(f"{DIM}› Activating external content...{RESET}")
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(output_path), UpdateLinks=3)  # 3 = always update
        wb.RefreshAll()
        wb.Save()
        wb.Close()
        print(f"{GREEN}  External content activated!{RESET}")
    except Exception as e:
        print(f"{YELLOW}  Note: Could not auto-refresh. Open file in Excel and click 'Enable Content'.{RESET}")

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
        print("Usage: python main.py <data.xlsx> [optional output file]")
        sys.exit(1)
    
    data_file = sys.argv[1]
    target_file = sys.argv[2] if len(sys.argv) >= 3 else None
    
    run_process(data_file, target_file)

if __name__ == "__main__":
    main()
