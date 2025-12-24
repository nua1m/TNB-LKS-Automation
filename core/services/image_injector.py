import pandas as pd
from core.so_utils import clean_so
from config import (
    DATA_START_ROW, SERVICE_ORDER_COL_IDX, COL_3MS_SO, COL_ATTACH_URL
)

class ImageInjector:
    @staticmethod
    def detect_type(url: str) -> str | None:
        if not url: return None
        u = url.lower()
        if "old_read" in u: return "old"
        if "card" in u: return "card"
        if "new_meter" in u: return "new"
        return None

    @staticmethod
    def build_url_map(data_path, sheet_name=None):
        """Reads raw data and maps SO -> {old, card, new} URLs."""
        # Need to fix circular import or just move import? Config is fine.
        from config import DATA_SHEET_NAME 
        target_sheet = sheet_name if sheet_name else (DATA_SHEET_NAME if DATA_SHEET_NAME else 0)

        df = pd.read_excel(data_path, sheet_name=target_sheet, dtype=str).fillna("")
        
        # Normalize Columns (Handle Preprocessor Uppercase output)
        df = df.rename(columns={
            "3MS SO NO.": COL_3MS_SO, "3MS SO NO": COL_3MS_SO,
            "ATTACHMENTS URL": COL_ATTACH_URL, "ATTACHMENTS URL": COL_ATTACH_URL,
            "ATTACHMENT URL": COL_ATTACH_URL,
        })
        
        # Forward fill SO numbers as per original logic
        df[COL_3MS_SO] = df[COL_3MS_SO].replace("", pd.NA).ffill()
        
        url_map = {}
        for _, row in df.iterrows():
            so = clean_so(row[COL_3MS_SO])
            url = str(row[COL_ATTACH_URL]).strip()
            if not so or not url: continue

            if so not in url_map:
                url_map[so] = {"old": None, "card": None, "new": None, "first": None}
            
            if url_map[so]["first"] is None:
                url_map[so]["first"] = url
            
            t = ImageInjector.detect_type(url)
            if t and url_map[so][t] is None:
                url_map[so][t] = url
        return url_map

    @staticmethod
    def img_formula(url: str) -> str | None:
        if not url: return None
        return f'=_xlfn.IMAGE("{url}",,1)'

    @staticmethod
    def set_formula(cell, formula):
        if not formula:
            cell.value = ""
            return
        cell.value = formula
        cell.data_type = "f"

    @staticmethod
    def run(handler, data_path, progress_cb=None, sheet_name=None):
        """Injects image formulas into Attachment sheet."""
        data_path = str(data_path) # pandas needs string
        url_map = ImageInjector.build_url_map(data_path, sheet_name=sheet_name)
        
        wsA = handler.ws_attach
        last_row = wsA.max_row
        if last_row < DATA_START_ROW: last_row = DATA_START_ROW
        
        col_old, col_card, col_new = 4, 5, 6
        idx = 0
        total = last_row - DATA_START_ROW + 1

        for r in range(DATA_START_ROW, last_row + 1):
            so = clean_so(wsA.cell(r, SERVICE_ORDER_COL_IDX).value)
            if not so: continue

            idx += 1
            imgs = url_map.get(so, {})
            
            old_url = imgs.get("old") or imgs.get("first")
            ImageInjector.set_formula(wsA.cell(r, col_old), ImageInjector.img_formula(old_url))
            ImageInjector.set_formula(wsA.cell(r, col_card), ImageInjector.img_formula(imgs.get("card")))
            ImageInjector.set_formula(wsA.cell(r, col_new), ImageInjector.img_formula(imgs.get("new")))

            if progress_cb:
                progress_cb(f"Processing SO {so} ({idx}/{total})")
