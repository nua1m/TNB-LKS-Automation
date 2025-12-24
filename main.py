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
from core.services.claim_service import ClaimService
from core.services.image_injector import ImageInjector
from core.services.quality_control import QualityControl
from core.services.preprocessor import Preprocessor

def main():
    if len(sys.argv) < 2:
        print("Usage: python main.py <data.xlsx> [template.xlsx]")
        sys.exit(1)

    data_path = Path(sys.argv[1]).resolve()
    
    if len(sys.argv) >= 3:
        template_path = Path(sys.argv[2]).resolve()
    else:
        # Use default template from config
        from config import DEFAULT_TEMPLATE_PATH
        # Try to resolve relative to current dir, then relative to project root if needed
        # Assuming run from root
        template_path = Path(DEFAULT_TEMPLATE_PATH).resolve()
        
        if not template_path.exists():
             # Fallback: check if it's in the current dir directly
             if Path(DEFAULT_TEMPLATE_PATH).name in [p.name for p in Path.cwd().glob("*")]:
                 template_path = Path(Path(DEFAULT_TEMPLATE_PATH).name).resolve()
                 
    if not template_path.exists():
        print(f"{RED}Error: Template file not found: {template_path}{RESET}")
        sys.exit(1)

    set_window_size(110, 40)
    show_title()

    print(f"{CYAN}RAW       : {RESET}{data_path}")
    print(f"{CYAN}TEMPLATE  : {RESET}{template_path}\n")

    start_time = time.time()

    # If using default template, we should NOT overwrite it.
    # Instead, we create a new output file name based on the Data file name.
    # E.g. "Data_Dec12.xlsx" -> "LKS_Dec12.xlsm"
    
    # Logic:
    # 1. If user provided 2nd arg explicitly, we assume they MIGHT want to overwrite it (or we can still treat it as "Template to Load"). 
    #    Actually current logic was: arg2 IS the file effectively modified.
    #    The user said: "just write python main.py data.xlsx and it generates a new LKS file".
    
    # So:
    # Template Input = template_path (from arg or default)
    # Output File = ?
    
    # Let's derive output name from Data file name, in the same folder as Data file.
    # "LKS Final <DataName>.xlsm"
    
    output_name = f"LKS ({data_path.stem}).xlsm"
    output_path = data_path.parent / output_name

    print(f"{CYAN}OUTPUT    : {RESET}{output_path}\n")

    start_time = time.time()

    # -----------------------------------------------------
    # INIT HANDLER
    # -----------------------------------------------------
    # We load the template, but save to output_path
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
            sys.exit(1)
        except Exception as e:
             print(f"{RED}Preprocessor Error: {e}{RESET}")
             sys.exit(1)

    # -----------------------------------------------------
    # STEP 1 — PROCESS DATA
    # -----------------------------------------------------
    print(f"{CYAN}› Reading RAW Data...{RESET}")
    # Read from source_path (either original or the intermediate cleaned one)
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
        print(f"{YELLOW}The LKS already contains data. Continue adding new SOs? (y/n){RESET}")
        if input(">> ").strip().lower() != "y":
            print(f"{RED}Aborted.{RESET}")
            handler.close()
            return
    
    if not new_rows:
        print(f"{YELLOW}All SOs already exist in TEMPLATE.{RESET}")
        handler.close()
        return

    # -----------------------------------------------------
    # STEP 3 — WRITE CLAIM & ATTACHMENT
    # -----------------------------------------------------
    print(f"{CYAN}› Writing CLAIM & ATTACHMENT...{RESET}")
    
    # Determine start rows
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

    print("› Saving Workbook...")
    handler.save()
    handler.close()

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

if __name__ == "__main__":
    main()
