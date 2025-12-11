import shutil
import sys
from pathlib import Path
from bs4 import BeautifulSoup
import xlwings as xw

# ==========================================
# CONFIGURATION
# ==========================================
HTML_DUMP_DIR_NAME = "html_dump_temp"
CACHE_DIR_NAME = "cache_images"

# HTML Column Indices (0-based)
# A=0, B=1 ... R=17, S=18
IDX_SO = 0   # Column A: SO Number
IDX_URL = 17 # Column R: URL Text
IDX_IMG = 18 # Column S: The Image

def detect_type(text: str) -> str | None:
    """Decides if image is old, ticket, or new based on URL text."""
    if not text: return None
    t = text.lower()
    if "old_read" in t: return "old"
    if "card" in t: return "ticket"
    if "new_meter" in t: return "new"
    return None

def extract_images(data_file: str):
    data_path = Path(data_file).resolve()
    base_dir = data_path.parent

    # 1. Setup Folders
    html_dir = base_dir / HTML_DUMP_DIR_NAME
    cache_dir = base_dir / CACHE_DIR_NAME

    if html_dir.exists(): shutil.rmtree(html_dir)
    if cache_dir.exists(): shutil.rmtree(cache_dir)
    
    html_dir.mkdir(parents=True, exist_ok=True)
    cache_dir.mkdir(parents=True, exist_ok=True)

    # 2. Export Excel to HTML (This will take 30-60 minutes for large files)
    print(f"⏳ Opening {data_path.name}...")
    print(f"   ⚠️  This may take 30-60 minutes for large files. Please be patient.")
    app = xw.App(visible=False)
    try:
        wb = app.books.open(str(data_path))
        dump_html = html_dir / "dump.htm"
        
        print(f"⏳ Exporting to HTML (this is the slow part)...")
        wb.api.SaveAs(str(dump_html), FileFormat=44) 
        
        wb.close()
        
    except Exception as e:
        print(f"❌ Excel export failed: {e}")
        try:
            wb_src.close()
        except: pass
        return
    finally:
        app.quit()

    # 3. LOCATE THE CORRECT SHEET FILE (CRITICAL FIX)
    # Excel creates 'dump.htm' (frameset) and 'dump_files/sheet001.htm' (data).
    # We MUST find 'sheet001.htm' or similar.
    
    dump_files_dir = html_dir / "dump_files"
    if not dump_files_dir.exists():
        # Sometimes (single sheet) it puts everything in the root
        dump_files_dir = html_dir

    # Find all 'sheet*.htm' files
    sheet_candidates = list(dump_files_dir.glob("sheet*.htm"))
    
    sheet_file = None
    if sheet_candidates:
        # Sort to pick sheet001.htm first
        sheet_candidates.sort(key=lambda x: x.name)
        sheet_file = sheet_candidates[0]
    else:
        # Fallback: maybe the main file is the data file (if single sheet)
        if (html_dir / "dump.htm").exists():
            sheet_file = html_dir / "dump.htm"

    if not sheet_file:
        print(f"❌ Could not find any HTML data file (sheet*.htm) in {html_dir}")
        return

    print(f"🔍 Parsing HTML Data: {sheet_file.name} ...")
    
    with open(sheet_file, "r", encoding="utf-8", errors="ignore") as f:
        soup = BeautifulSoup(f, "html.parser")

    # 4. Iterate Rows and Map
    table = soup.find("table")
    if not table:
        print(f"❌ No <table> found in {sheet_file.name}. Is the sheet empty?")
        return

    saved_count = 0
    last_so = None 

    rows = table.find_all("tr")
    print(f"   Scanning {len(rows)} HTML rows...")
    
    from ui.components import step_progress

    total_rows = len(rows)
    for i, tr in enumerate(rows):
        # Update progress bar every 10 rows or last row
        if i % 10 == 0 or i == total_rows - 1:
            step_progress("EXTRACT", i + 1, total_rows, extra=f"Saved: {saved_count}", spinner_i=i)

        tds = tr.find_all("td")
        
        # Ensure row is long enough to have an image column
        if len(tds) <= IDX_IMG:
            continue

        # --- A. GET SO NUMBER (Fill Down Logic) ---
        raw_so = tds[IDX_SO].get_text(strip=True)
        if raw_so and len(raw_so) > 4 and "SO" not in raw_so: 
            last_so = raw_so
        
        if not last_so: 
            continue

        # --- B. DETECT IMAGE TYPE FROM URL ---
        url_text = tds[IDX_URL].get_text(strip=True)
        img_type = detect_type(url_text)

        # --- C. CHECK FOR IMAGE ---
        img_tag = tds[IDX_IMG].find("img")
        
        if img_tag and img_type:
            src_filename = Path(img_tag.get("src")).name
            
            # The image might be in 'dump_files' or the same dir as the sheet
            src_path = sheet_file.parent / src_filename
            
            # Fallback check if it's in the root of dump_files
            if not src_path.exists():
                src_path = dump_files_dir / src_filename

            if src_path.exists():
                new_name = f"{last_so}_{img_type}.png"
                dst_path = cache_dir / new_name
                
                shutil.copy2(src_path, dst_path)
                saved_count += 1
                # print(f"   Saved {new_name}")

    # 5. Cleanup
    # try:
    #     shutil.rmtree(html_dir)
    # except:
    #     pass

    print("\n" + "="*40)
    print(f"✅ DONE! Extracted {saved_count} mapped images.")
    print(f"📂 Images are in: {cache_dir}")
    print("="*40)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python extract_images.py <Data.xlsx>")
    else:
        extract_images(sys.argv[1])
