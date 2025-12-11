# core/text.py — RAW → CLAIM rows + write CLAIM & ATTACHMENT (v3.0)

from pathlib import Path
from datetime import datetime

import pandas as pd

from core.so_utils import clean_so
from config import (
    DATA_SHEET_NAME,
    HEADER_ROW,
    COL_3MS_SO, COL_CONTRACT, COL_SO_STATUS, COL_USER_STATUS,
    COL_ADDRESS, COL_VOLTAGE, COL_SO_TYPE, COL_SO_DESC,
    COL_TECHNICIAN, COL_STATUS_DATE, COL_SITE_ID,
    COL_OLD_METER, COL_OLD_COMM, COL_NEW_METER, COL_NEW_COMM,
)


# ============================================================
#   DATE / BUSINESS HELPERS
# ============================================================

def normalize_status_date_string(raw):
    """Convert RAW date to 'Jan 20, 2025, 10:43 PM' format."""
    if raw is None:
        return ""
    s = str(raw).strip()
    if not s:
        return ""

    # Many 3MS exports use these patterns
    for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%b %d, %Y, %I:%M %p")
        except ValueError:
            pass

    # Fallback for pure YYYY-MM-DD
    try:
        dt = datetime.strptime(s, "%Y-%m-%d")
        return dt.strftime("%b %d, %Y, %I:%M %p")
    except ValueError:
        return s


def extract_date_only(raw_text):
    """Extract 'Jan 02, 2025' → date() object."""
    if not raw_text:
        return None

    s = str(raw_text).strip()
    if "," in s:
        parts = s.split(",")
        if len(parts) >= 2:
            s = parts[0].strip() + ", " + parts[1].strip()

    try:
        dt = datetime.strptime(s, "%b %d, %Y")
        return dt.date()
    except ValueError:
        return None


def hari_biasa_weekend_from_raw(raw_text):
    d = extract_date_only(raw_text)
    if not d:
        return ""
    return "Hujung Minggu" if d.weekday() == 6 else "Hari Biasa"


def business_area_from_site(site_id):
    if site_id is None:
        return ""
    s = str(site_id).strip()
    if s == "6340":
        return "Johor Bahru"
    if s == "6346":
        return "Johor Jaya"
    return ""


# ============================================================
#   BUILD CLAIM ROWS (RAW → CLEAN DICT LIST)
# ============================================================

def build_claim_rows(data_path: Path | str):
    """
    Loads RAW DATA and returns:
      - claim_rows: list of dict
      - stats: {
            total_sos_raw,
            tras_removed,
            duplicates_skipped,
            sos_after_tras,
            invalid_dates,
            missing_address,
            missing_technician,
            same_old_new_meter
        }
    Soft validation: warn via stats, only hard-fail on missing SO column
    or completely empty RAW.
    """
    data_path = Path(data_path)

    # Read sheet
    if DATA_SHEET_NAME:
        df = pd.read_excel(data_path, sheet_name=DATA_SHEET_NAME,
                           header=HEADER_ROW - 1, dtype=str)
    else:
        df = pd.read_excel(data_path, header=HEADER_ROW - 1, dtype=str)

    # Normalize col names
    df = df.rename(columns={
        "3MS SO No.": COL_3MS_SO,
        "3MS SO No": COL_3MS_SO,
        "SO Number": COL_3MS_SO,
        "Contract Account": COL_CONTRACT,
        "SO Status": COL_SO_STATUS,
        "User Status": COL_USER_STATUS,
        "Address": COL_ADDRESS,
        "Voltage": COL_VOLTAGE,
        "SO Type": COL_SO_TYPE,
        "SO Description": COL_SO_DESC,
        "Technician": COL_TECHNICIAN,
        "Status Date": COL_STATUS_DATE,
        "Site ID": COL_SITE_ID,
        "Old Meter no": COL_OLD_METER,
        "Old Meter No": COL_OLD_METER,
        "Old Comm Module": COL_OLD_COMM,
        "New Meter no": COL_NEW_METER,
        "New Meter No": COL_NEW_METER,
        "New Comm Module": COL_NEW_COMM,
    })

    if COL_3MS_SO not in df.columns:
        raise KeyError(f"Missing '{COL_3MS_SO}' in RAW DATA.")

    # Soft check: required-but-not-fatal cols
    required_soft = [
        COL_STATUS_DATE,
        COL_ADDRESS,
        COL_TECHNICIAN,
        COL_SO_TYPE,
        COL_OLD_METER,
        COL_NEW_METER,
    ]
    # (we won't crash if missing, but stats may show effects)
    # Remove blank rows
    df = df[df[COL_3MS_SO].astype(str).str.strip() != ""]
    if df.empty:
        raise ValueError("RAW DATA has no valid SO rows.")

    total_sos_raw = 0
    tras_removed = 0
    duplicates_skipped = 0
    invalid_dates = 0
    missing_address = 0
    missing_technician = 0
    same_old_new_meter = 0

    seen = set()
    so_groups = []

    # Group by SO order in Excel
    for so, subdf in df.groupby(COL_3MS_SO, sort=False):
        so_clean = clean_so(so)
        if not so_clean:
            continue

        total_sos_raw += 1

        # Duplicate SO?
        if so_clean in seen:
            duplicates_skipped += 1
            continue
        seen.add(so_clean)

        # Remove TRAS rows
        if COL_USER_STATUS in subdf.columns:
            if subdf[COL_USER_STATUS].astype(str).str.upper().str.contains("TRAS").any():
                tras_removed += 1
                continue

        row0 = subdf.iloc[0]

        raw_status = row0.get(COL_STATUS_DATE, "") or ""
        status_str = normalize_status_date_string(raw_status)
        date_obj = extract_date_only(status_str)
        if date_obj is None and str(raw_status).strip():
            invalid_dates += 1

        # Data-quality checks
        addr = (row0.get(COL_ADDRESS, "") or "").strip()
        tech = (row0.get(COL_TECHNICIAN, "") or "").strip()
        old_dev = (row0.get(COL_OLD_METER, "") or "").strip()
        new_dev = (row0.get(COL_NEW_METER, "") or "").strip()

        if not addr:
            missing_address += 1
        if not tech:
            missing_technician += 1
        if old_dev and new_dev and old_dev == new_dev:
            same_old_new_meter += 1

        so_groups.append({
            "so": so_clean,
            "row0": row0,
            "status_str": status_str,
            "date_obj": date_obj,
            "site_id": row0.get(COL_SITE_ID, "") or "",
        })

    if not so_groups:
        raise ValueError("After removing TRAS and invalid rows, no SOs remain.")

    # Sort SOs by date (unknown dates go first)
    so_groups.sort(key=lambda g: g["date_obj"] or datetime.min.date())

    # Build claim rows dicts
    claim_rows = []
    qty = 1

    for g in so_groups:
        row0 = g["row0"]

        ba = business_area_from_site(g["site_id"])
        hari = hari_biasa_weekend_from_raw(g["status_str"])

        claim_rows.append({
            "Qty": qty,
            "Service Order": g["so"],
            "Account Number": row0.get(COL_CONTRACT, "") or "",
            "Status": row0.get(COL_SO_STATUS, "") or "",
            "Address": row0.get(COL_ADDRESS, "") or "",
            "Voltage": row0.get(COL_VOLTAGE, "") or "",
            "SO Description": row0.get(COL_SO_TYPE, "") or row0.get(COL_SO_DESC, "") or "",
            "Labor": row0.get(COL_TECHNICIAN, "") or "",
            "Status Date": g["status_str"],
            "Site": g["site_id"],
            "Business Area": ba,
            "Old Device No": row0.get(COL_OLD_METER, "") or "",
            "New Device No": row0.get(COL_NEW_METER, "") or "",
            "Comm Module No": row0.get(COL_NEW_COMM, "") or "",
            "Hari Field": hari,
            "Jenis Kerja": "KERJA BIASA",
            "Remarks 1": "",
            "Remarks 2": "",
        })

        qty += 1

    stats = {
        "total_sos_raw": total_sos_raw,
        "tras_removed": tras_removed,
        "duplicates_skipped": duplicates_skipped,
        "sos_after_tras": len(so_groups),
        "invalid_dates": invalid_dates,
        "missing_address": missing_address,
        "missing_technician": missing_technician,
        "same_old_new_meter": same_old_new_meter,
    }

    return claim_rows, stats


# ============================================================
#   WRITE CLAIM SHEET
# ============================================================

def write_to_claim_sheet(ws, rows, start_row=3):
    """Writes CLAIM data to Excel (no progress bar here)."""

    col_map = {
        "Service Order": 2,
        "Account Number": 3,
        "Status": 4,
        "Address": 5,
        "Voltage": 6,
        "SO Description": 7,
        "Labor": 8,
        "Status Date": 9,
        "Site": 10,
        "Business Area": 11,
        "Old Device No": 12,
        "New Device No": 13,
        "Comm Module No": 14,
        "Hari Field": 15,
        "Jenis Kerja": 16,
        "Remarks 1": 17,
        "Remarks 2": 18,
    }

    text_cols = {
        "Service Order", "Account Number", "Site",
        "Old Device No", "New Device No", "Comm Module No"
    }

    for i, row in enumerate(rows, start=0):
        r = start_row + i
        for field, col in col_map.items():
            val = row.get(field)
            cell = ws.cell(row=r, column=col)
            if field in text_cols:
                cell.value = "" if val is None else str(val)
                cell.number_format = "@"
            else:
                cell.value = val

    print(f"[CLAIM] Finished writing {len(rows)} rows (up to row {start_row + len(rows) - 1}).")


# ============================================================
#   WRITE ATTACHMENT SHEET
# ============================================================

def write_to_attachment_sheet(ws, rows, start_row=3):
    """
    Writes:
      B = Service Order
      C = Old Device No
    """

    for i, row in enumerate(rows, start=0):
        r = start_row + i

        so = clean_so(row.get("Service Order", ""))
        old_dev = str(row.get("Old Device No", "")).strip()

        ws.cell(row=r, column=2).value = so
        ws.cell(row=r, column=2).number_format = "@"

        ws.cell(row=r, column=3).value = old_dev
        ws.cell(row=r, column=3).number_format = "@"

    print(f"[ATTACH] Finished writing {len(rows)} rows (up to row {start_row + len(rows) - 1}).")
