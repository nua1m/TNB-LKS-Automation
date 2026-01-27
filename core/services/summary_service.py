from datetime import datetime
from config import CLAIM_SHEET_NAME, SUMMARY_SHEET_NAME

class SummaryService:
    @staticmethod
    def update_summary(handler):
        """
        Updates the SUMMARY sheet with statistics from the CLAIM sheet.
        Calculates:
        - Total Count
        - Date Range (Min - Max)
        - Breakdown by Business Area (Johor Bahru vs Johor Jaya)
        """
        wb = handler.wb
        ws_claim = handler.ws_claim
        
        # Check/Create Summary Sheet
        if SUMMARY_SHEET_NAME in wb.sheetnames:
            ws_summary = wb[SUMMARY_SHEET_NAME]
        else:
            ws_summary = wb.create_sheet(SUMMARY_SHEET_NAME)

        # 1. Read Data for Stats
        # Business Area is Col 11 (K)
        # Status Date is Col 9 (I)
        
        ba_counts = {}
        dates = []
        
        total_rows = 0
        
        for r in range(3, ws_claim.max_row + 1):
            # Check if row has SO (Col 2)
            so_val = ws_claim.cell(r, 2).value
            if not so_val: continue
            
            total_rows += 1
            
            # Business Area
            ba = str(ws_claim.cell(r, 11).value).strip()
            if ba:
                ba_counts[ba] = ba_counts.get(ba, 0) + 1
                
            # Date
            dt_val = ws_claim.cell(r, 9).value
            if isinstance(dt_val, datetime):
                dates.append(dt_val)

        # 2. Calculate Derived Stats
        min_date = min(dates) if dates else None
        max_date = max(dates) if dates else None
        
        if min_date and max_date:
            date_range_str = f"{min_date.strftime('%d %b %Y')} - {max_date.strftime('%d %b %Y')}"
        else:
            date_range_str = "N/A"

        # 3. Write to Summary Sheet
        # Simple Layout
        
        # Header
        ws_summary['B2'] = "LKS REPORT SUMMARY"
        ws_summary['B2'].font = ws_summary['B2'].font.copy(bold=True, size=14)
        
        # Stats
        ws_summary['B4'] = "Total Claims"
        ws_summary['C4'] = total_rows
        
        ws_summary['B5'] = "Date Range"
        ws_summary['C5'] = date_range_str
        
        # Business Area Breakdown
        row_idx = 7
        ws_summary.cell(row=row_idx, column=2, value="Business Area Breakdown").font = ws_summary['B2'].font.copy(bold=True, size=11)
        row_idx += 1
        
        for ba, count in ba_counts.items():
            ws_summary.cell(row=row_idx, column=2, value=ba)
            ws_summary.cell(row=row_idx, column=3, value=count)
            row_idx += 1
            
        print(f"  Summary Updated: {total_rows} claims. Range: {date_range_str}")
