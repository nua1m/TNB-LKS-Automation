import pandas as pd
from datetime import datetime
from pathlib import Path

from core.so_utils import clean_so
from core.services.date_engine import DateEngine
from config import (
    DATA_SHEET_NAME, HEADER_ROW,
    COL_3MS_SO, COL_CONTRACT, COL_SO_STATUS, COL_USER_STATUS,
    COL_ADDRESS, COL_VOLTAGE, COL_SO_TYPE, COL_SO_DESC,
    COL_TECHNICIAN, COL_STATUS_DATE, COL_SITE_ID,
    COL_OLD_METER, COL_NEW_METER, COL_NEW_COMM
)

# Helpers


def get_business_area(site_id):
    s = str(site_id).strip()
    if s == "6340": return "Johor Bahru"
    if s == "6346": return "Johor Jaya"
    return ""

class ClaimService:
    @staticmethod
    def build_rows(data_path, sheet_name=None):
        """Reads RAW data and returns processed rows + stats."""
        data_path = Path(data_path)
        
        # Load Data
        target_sheet = sheet_name if sheet_name else DATA_SHEET_NAME
        
        df = pd.read_excel(
            data_path, 
            sheet_name=target_sheet, 
            header=HEADER_ROW - 1, 
            dtype=str
        )
        
        # Handle case where sheet_name=None returns a dict of all sheets
        if isinstance(df, dict):
            # Use the first sheet found
            df = list(df.values())[0]

        # Normalize Columns
        # Note: Preprocessor converts headers to UPPERCASE.
        df = df.rename(columns={
            "3MS SO No.": COL_3MS_SO, "3MS SO No": COL_3MS_SO, "SO Number": COL_3MS_SO, "3MS SO NO.": COL_3MS_SO,
            "Contract Account": COL_CONTRACT, "CONTRACT ACCOUNT": COL_CONTRACT,
            "SO Status": COL_SO_STATUS, "SO STATUS": COL_SO_STATUS,
            "User Status": COL_USER_STATUS, "USER STATUS": COL_USER_STATUS,
            "Address": COL_ADDRESS, "ADDRESS": COL_ADDRESS,
            "Voltage": COL_VOLTAGE, "VOLTAGE": COL_VOLTAGE,
            "SO Type": COL_SO_TYPE, "SO TYPE": COL_SO_TYPE,
            "SO Description": COL_SO_DESC, "SO DESCRIPTION": COL_SO_DESC,
            "Technician": COL_TECHNICIAN, "TECHNICIAN": COL_TECHNICIAN,
            "Status Date": COL_STATUS_DATE, "STATUS DATE": COL_STATUS_DATE,
            "Site ID": COL_SITE_ID, "SITE ID": COL_SITE_ID,
            "Old Meter no": COL_OLD_METER, "Old Meter No": COL_OLD_METER, "OLD METER NO": COL_OLD_METER,
            "New Meter no": COL_NEW_METER, "New Meter No": COL_NEW_METER, "NEW METER NO": COL_NEW_METER,
            "New Comm Module": COL_NEW_COMM, "NEW COMM MODULE": COL_NEW_COMM,
        })

        if COL_3MS_SO not in df.columns:
            # Debugging helper: Identify close matches or print all
            raise KeyError(f"Missing '{COL_3MS_SO}' in RAW DATA. Available columns: {list(df.columns)}")

        df = df[df[COL_3MS_SO].astype(str).str.strip() != ""]
        if df.empty: raise ValueError("RAW DATA has no valid SO rows.")

        stats = {
            "total_sos_raw": 0, "tras_removed": 0, "duplicates_skipped": 0,
            "sos_after_tras": 0, "invalid_dates": 0, "missing_address": 0
        }

        so_groups = []
        tras_rows = []
        seen = set()

        for so, subdf in df.groupby(COL_3MS_SO, sort=False):
            so_clean = clean_so(so)
            if not so_clean: continue
            stats["total_sos_raw"] += 1

            if so_clean in seen:
                stats["duplicates_skipped"] += 1
                continue
            seen.add(so_clean)

            if COL_USER_STATUS in subdf.columns:
                if subdf[COL_USER_STATUS].astype(str).str.upper().str.contains("TRAS").any():
                    stats["tras_removed"] += 1
                    # Capture first row of TRAS group (simplest approach)
                    tras_rows.append(subdf.iloc[0])
                    continue
            
            # Use first row of the group
            row0 = subdf.iloc[0]
            
            raw_status = row0.get(COL_STATUS_DATE, "") or ""
            # Parse date using DateEngine
            # We want to keep the full datetime object for writing to Excel later
            date_obj = DateEngine.parse_datetime(raw_status)
            status_str = str(raw_status) # Keep original just in case, or for debug
            
            # If we successfully parsed a datetime, we can format it for display in logs if needed
            # but for the "Status Date" field in the row, we should store the object.
            
            # date_obj is now a datetime object (or None)
            
            if not date_obj and str(raw_status).strip():
                stats["invalid_dates"] += 1
            if not str(row0.get(COL_ADDRESS, "") or "").strip():
                stats["missing_address"] += 1

            so_groups.append({
                "so": so_clean,
                "row0": row0,
                "status_val": date_obj if date_obj else raw_status, # Store object or raw string
                "date_obj": date_obj,
                "site_id": row0.get(COL_SITE_ID, "") or ""
            })

        # Helper to create Record
        def create_record(idx, r_dict, stat_val_obj):
            site_id = r_dict.get(COL_SITE_ID, "") or ""
            # DateEngine likely expects a string or handles object? 
            # If status_val is datetime, convert to string formatted roughly?
            # Let's ensure we pass what worked before.
            # Before: status_val was PASSED "date_obj if date_obj else raw_status"
            # It seems DateEngine might have been updated to handle objects or it failed silently.
            # Safest: Convert obj to string if DateEngine expects string.
            # Assuming DateEngine.calculate is robust.
            
            logic = DateEngine.calculate(stat_val_obj, ocr_date_str=None)
            
            # Status Date: if obj, format it. If string, use as is.
            disp_date = stat_val_obj
            if isinstance(stat_val_obj, datetime):
                disp_date = stat_val_obj.strftime("%d.%m.%Y")
            
            return {
                "Qty": idx,
                "Service Order": r_dict.get(COL_3MS_SO, "") or "",
                "Account Number": r_dict.get(COL_CONTRACT, "") or "",
                "Status": r_dict.get(COL_SO_STATUS, "") or "",
                "Address": r_dict.get(COL_ADDRESS, "") or "",
                "Voltage": r_dict.get(COL_VOLTAGE, "") or "",
                "SO Description": r_dict.get(COL_SO_TYPE, "") or r_dict.get(COL_SO_DESC, "") or "",
                "Labor": r_dict.get(COL_TECHNICIAN, "") or "",
                "Status Date": disp_date,
                "Site": site_id,
                "Business Area": get_business_area(site_id),
                "Old Device No": r_dict.get(COL_OLD_METER, "") or "",
                "New Device No": r_dict.get(COL_NEW_METER, "") or "",
                "Comm Module No": r_dict.get(COL_NEW_COMM, "") or "",
                
                # Derived Fields
                "Hari Field": logic["hari"],
                "Jenis Kerja": "KERJA BIASA",
                "Remarks 1": logic["remarks_1"],
                "Remarks 2": logic["remarks_2"],
            }

        # ---------------------------------------------------------------------
        # 1. Process Main SO Groups
        # ---------------------------------------------------------------------
        # Sort by datetime (date + time) from oldest to newest
        so_groups.sort(key=lambda g: g["date_obj"] if g["date_obj"] else datetime.min)
        stats["sos_after_tras"] = len(so_groups)

        rows = []
        for i, g in enumerate(so_groups, 1):
            rows.append(create_record(i, g["row0"], g["status_val"]))

        # ---------------------------------------------------------------------
        # 2. Process TRAS Rows
        # ---------------------------------------------------------------------
        tras_formatted = []
        
        # Helper to get date for sorting TRAS
        def get_tras_date_obj(series):
            raw = str(series.get(COL_STATUS_DATE, "")).strip()
            return DateEngine.parse_datetime(raw) # Helper from valid engine? or just use pd
        
        # Sort TRAS by date
        # Use simple pandas parse for sorting locally
        tras_rows.sort(key=lambda r: pd.to_datetime(str(r.get(COL_STATUS_DATE,"")), dayfirst=True, errors='coerce') or datetime.min)
        
        for i, series in enumerate(tras_rows, 1):
            r_dict = series.to_dict()
            raw_d = str(r_dict.get(COL_STATUS_DATE, "")).strip()
            # Try to reproduce the 'status_val' logic (Object or Raw String)
            try:
                dt = pd.to_datetime(raw_d, dayfirst=True)
                stat_val = dt
            except:
                stat_val = raw_d
                
            tras_formatted.append(create_record(i, r_dict, stat_val))

        return rows, tras_formatted, stats
        
        return rows, tras_rows, stats

    @staticmethod
    def export_tras(tras_rows, output_path):
        """Writes TRAS rows to a separate Excel file."""
        if not tras_rows: return
        
        # We export essentially the raw columns for TRAS rows
        df = pd.DataFrame(tras_rows)
        # Clean up: If they are pandas Series, DataFrame constructor works fine.
        
        print(f"  Writing {len(tras_rows)} TRAS rows to {output_path.name}...")
        df.to_excel(output_path, index=False)

    @staticmethod
    def write_data(handler, rows, start_claim=3, start_attach=3):
        """Writes rows to Claim and Attachment sheets using ExcelHandler."""
        wsC = handler.ws_claim
        wsA = handler.ws_attach
        
        # CLAIM SHEET MAP
        col_map = {
            "Service Order": 2, "Account Number": 3, "Status": 4, "Address": 5,
            "Voltage": 6, "SO Description": 7, "Labor": 8, "Status Date": 9,
            "Site": 10, "Business Area": 11, "Old Device No": 12, "New Device No": 13,
            "Comm Module No": 14, "Hari Field": 15, "Jenis Kerja": 16,
            "Remarks 1": 17, "Remarks 2": 18,
        }
        text_cols = {"Service Order", "Account Number", "Site", "Old Device No", "New Device No", "Comm Module No"}

        for i, row in enumerate(rows):
            # Write Claim
            rC = start_claim + i
            for field, col in col_map.items():
                val = row.get(field)
                cell = wsC.cell(row=rC, column=col)
                if field in text_cols:
                    cell.value = "" if val is None else str(val)
                    cell.number_format = "@"
                elif field == "Status Date" and isinstance(val, (datetime, pd.Timestamp)):
                     cell.value = val
                     cell.number_format = "d mmm, yyyy, h:mm AM/PM"
                else:
                    cell.value = val

            # Write Attachment (SO + Old Meter)
            rA = start_attach + i
            so = clean_so(row.get("Service Order", ""))
            old_dev = str(row.get("Old Device No", "")).strip()
            
            cSO = wsA.cell(row=rA, column=2)
            cSO.value = so
            cSO.number_format = "@"

            cOld = wsA.cell(row=rA, column=3)
            cOld.value = old_dev
            cOld.number_format = "@"

        print(f"Written {len(rows)} rows to Claim & Attachment.")
