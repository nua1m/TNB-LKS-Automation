import re
from collections import Counter
from datetime import datetime, timedelta, timezone

import pandas as pd
from core.so_utils import clean_so
from core.services.date_engine import DateEngine
from config import (
    DATA_START_ROW, SERVICE_ORDER_COL_IDX, COL_3MS_SO, COL_ATTACH_URL
)

MALAYSIA_TZ = timezone(timedelta(hours=8))

class ImageInjector:
    @staticmethod
    def _normalize_text(text: str) -> str:
        return re.sub(r"[^a-z0-9]+", " ", str(text or "").lower()).strip()

    @staticmethod
    def _contains_any(text: str, needles: list[str]) -> bool:
        return any(needle in text for needle in needles)

    @staticmethod
    def choose_diskon_url(img_info):
        for image_type in ("old", "new", "card", "first"):
            url = img_info.get(image_type)
            if url:
                return url, None
        return None, "No image timestamp found"

    @staticmethod
    def detect_type(*texts: str) -> str | None:
        url_text = str(texts[0] or "") if texts else ""
        attachment_name = str(texts[1] or "") if len(texts) > 1 else ""
        combined = f"{url_text} {attachment_name}".strip().lower()
        if not combined:
            return None

        filename = url_text.lower().split("/")[-1] if url_text else combined
        normalized_filename = ImageInjector._normalize_text(filename)
        normalized_name = ImageInjector._normalize_text(attachment_name)
        normalized_combined = ImageInjector._normalize_text(combined)

        # OLD READ / OLD METER
        if ImageInjector._contains_any(
            normalized_combined,
            [
                "old read",
                "oldread",
                "old red",
                "old rea",
                "old meter",
                "oldmeter",
                "meter lama",
                "lama meter",
            ],
        ):
            return "old"

        # CARD
        if ImageInjector._contains_any(
            normalized_name,
            [
                "card",
                "meter card",
                "kad",
                "meter kad",
                "kad meter",
            ],
        ):
            return "card"
        if any(x in normalized_filename for x in ["card", "cad", "crd", "car", "ard"]):
            return "card"

        # NEW METER
        if ImageInjector._contains_any(
            normalized_combined,
            [
                "new meter",
                "newmeter",
                "neww meter",
                "newwmeter",
                "newmemeter",
                "nee meter",
                "nee meter",
                "new metwr",
                "new meer",
                "mew meter",
                "mewmeter",
                "ew meter",
                "ewmeter",
                "nw meter",
                "new mter",
                "new metr",
                "new mete",
                "new kete",
                "new mtr",
                "newmeterr",
                "new eter",
                "meter baru",
                "baru meter",
            ],
        ):
            return "new"

        # Fallback fuzzy for new meter
        if (
            ("new m" in normalized_combined or "mew m" in normalized_combined or "ew m" in normalized_combined)
            and ("eer" in normalized_combined or "ter" in normalized_combined or "etr" in normalized_combined or "meter" in normalized_combined)
        ):
            return "new"

        return None

    @staticmethod
    def _find_attachment_name_column(columns: list[str], url_header: str) -> str | None:
        explicit_names = {
            "ATTACHMENT NAME",
            "ATTACHMENTS NAME",
            "FILE NAME",
            "ATTACHMENT FILENAME",
            "ATTACHMENT FILE NAME",
            "DOCUMENT NAME",
            "IMAGE NAME",
        }
        for column in columns:
            if column in explicit_names:
                return column
            if "ATTACH" in column and "NAME" in column:
                return column
            if "FILE" in column and "NAME" in column:
                return column

        try:
            url_idx = columns.index(url_header)
        except ValueError:
            return None

        if url_idx > 0:
            candidate = columns[url_idx - 1]
            if "NAME" in candidate:
                return candidate
        return None

    @staticmethod
    def build_url_map(data_path, sheet_name=None):
        """Reads raw data and maps SO -> image URLs plus old-image timestamp metadata."""
        from config import DATA_SHEET_NAME 
        target_sheet = sheet_name if sheet_name else (DATA_SHEET_NAME if DATA_SHEET_NAME else 0)

        df = pd.read_excel(data_path, sheet_name=target_sheet, dtype=str).fillna("")
        
        # Normalize Columns: Strip and Upper
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # Mapping: keys must be UPPERCASE versions of expected headers
        target_so = COL_3MS_SO.upper()
        target_url = COL_ATTACH_URL.upper()
        
        rename_map = {}
        attachment_name_col = ImageInjector._find_attachment_name_column(df.columns.tolist(), target_url)
        for c in df.columns:
            if c in ["3MS SO NO.", "3MS SO NO", "SO NUMBER", target_so]:
                rename_map[c] = COL_3MS_SO
            elif c in ["ATTACHMENTS URL", "ATTACHMENT URL", "ATTACHMENT_URL", target_url]:
                rename_map[c] = COL_ATTACH_URL
                
        df = df.rename(columns=rename_map)
        
        # Debugging check
        if COL_3MS_SO not in df.columns or COL_ATTACH_URL not in df.columns:
            print(f"Warning: Columns not found. Available: {df.columns.tolist()}")
        
        # Forward fill SO numbers as per original logic
        if COL_3MS_SO in df.columns:
            df[COL_3MS_SO] = df[COL_3MS_SO].replace("", pd.NA).ffill()
        else:
            return {} 
        
        url_map = {}
        # ... rest of loop logic is fine if columns exist
        for _, row in df.iterrows():
            if COL_3MS_SO not in row or COL_ATTACH_URL not in row: continue
            
            so = clean_so(row[COL_3MS_SO])
            url = str(row[COL_ATTACH_URL]).strip()
            attachment_name = str(row.get(attachment_name_col, "")).strip() if attachment_name_col else ""
            if not so or not url: continue

            if so not in url_map:
                url_map[so] = {
                    "old": None,
                    "card": None,
                    "new": None,
                    "first": None,
                }
            
            if url_map[so]["first"] is None:
                url_map[so]["first"] = url
            
            t = ImageInjector.detect_type(url, attachment_name)
            if t and url_map[so][t] is None:
                url_map[so][t] = url
        return url_map

    @staticmethod
    def extract_url_timestamp(url: str):
        if not url:
            return None
        match = re.search(r"(\d{10,17})(?=\.[A-Za-z0-9]+(?:$|\?))", str(url))
        if not match:
            return None

        raw_value = match.group(1)
        try:
            timestamp_value = int(raw_value)
        except ValueError:
            return None

        if len(raw_value) >= 13:
            timestamp_value = timestamp_value / 1000

        try:
            return datetime.fromtimestamp(timestamp_value, tz=timezone.utc).astimezone(MALAYSIA_TZ)
        except (OverflowError, OSError, ValueError):
            return None

    @staticmethod
    def format_diskon_date(date_value):
        if not date_value:
            return ""
        return f"{date_value.day} {date_value.strftime('%B').upper()}"

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
        """Injects image formulas into Attachment sheet and returns image metadata."""
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

        return url_map

    @staticmethod
    def apply_diskon_checks(handler, url_map, target_sos=None):
        ws_claim = handler.ws_claim
        ws_attach = handler.ws_attach
        target_so_set = set(target_sos or [])

        claim_rows = {}
        for row_idx in range(DATA_START_ROW, ws_claim.max_row + 1):
            so = clean_so(ws_claim.cell(row_idx, SERVICE_ORDER_COL_IDX).value)
            if so:
                claim_rows[so] = row_idx

        review_rows = {}
        diskon_by_date = Counter()
        diskon_count = 0
        diskon_rows = {}

        for attach_row in range(DATA_START_ROW, ws_attach.max_row + 1):
            so = clean_so(ws_attach.cell(attach_row, SERVICE_ORDER_COL_IDX).value)
            if not so:
                continue
            if target_so_set and so not in target_so_set:
                continue

            claim_row = claim_rows.get(so)
            if not claim_row:
                continue

            attach_date_cell = ws_attach.cell(attach_row, 8)
            claim_remark_cell = ws_claim.cell(claim_row, 17)
            claim_review_cell = ws_claim.cell(claim_row, 18)
            attach_date_cell.value = ""
            claim_remark_cell.value = ""
            claim_review_cell.value = ""

            if so not in url_map:
                continue
            img_info = url_map.get(so, {})

            review_reason = None

            chosen_url, review_reason = ImageInjector.choose_diskon_url(img_info)
            if chosen_url and not review_reason:
                chosen_datetime = ImageInjector.extract_url_timestamp(chosen_url)
                if not chosen_datetime:
                    review_reason = "Selected image has no timestamp"
                else:
                    chosen_date = chosen_datetime.date()
                    claim_datetime = ws_claim.cell(claim_row, 9).value
                    if isinstance(claim_datetime, datetime):
                        claim_date = claim_datetime.date()
                    else:
                        claim_date = DateEngine.parse_date(claim_datetime)

                    if not claim_date:
                        review_reason = "Missing CLAIM status date"
                    elif chosen_date != claim_date:
                        formatted_date = ImageInjector.format_diskon_date(chosen_date)
                        attach_date_cell.value = formatted_date
                        claim_remark_cell.value = f"TECO LEWAT SEBAB DISKON ({formatted_date})"
                        diskon_by_date[formatted_date] += 1
                        diskon_count += 1
                        diskon_rows[so] = formatted_date

            if review_reason:
                review_rows[so] = review_reason
                claim_review_cell.value = review_reason

        return {
            "diskon_count": diskon_count,
            "diskon_by_date": dict(
                sorted(
                    diskon_by_date.items(),
                    key=lambda item: datetime.strptime(item[0], "%d %B").replace(year=2000),
                )
            ),
            "diskon_rows": diskon_rows,
            "review_rows": review_rows,
        }
