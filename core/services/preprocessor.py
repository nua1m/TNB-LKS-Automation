import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

class Preprocessor:
    @staticmethod

    def clean_raw_data(file_path):
        """
        Converts .xls (legacy/XML) to .xlsx using Excel COM interface for robustness,
        then reads, skips header, drops Col D.
        Returns: DataFrame ready for insertion.
        """
        import win32com.client as win32
        import os
        from pathlib import Path
        
        file_path = str(Path(file_path).resolve())
        temp_xlsx = file_path + "_temp.xlsx"
        
        # 1. Convert to XLSX using Excel App (Handles XML/Format mismatch gracefully)
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = None
        try:
            print(f"  Converting {Path(file_path).name} to .xlsx format...")
            # CorruptLoad=1 helps with problematic/older files
            wb = excel.Workbooks.Open(file_path, UpdateLinks=0, CorruptLoad=1)
            # FileFormat=51 is xlOpenXMLWorkbook (.xlsx)
            wb.SaveAs(temp_xlsx, FileFormat=51)
            wb.Close(SaveChanges=False)
        except Exception as e:
            print(f"Excel Conversion Failed: {e}")
            if wb: 
                try: wb.Close(SaveChanges=False)
                except: pass
            raise e
        finally:
            try:
                excel.Quit()
            except:
                pass

        # 2. Read the new clean XLSX
        try:
            df = pd.read_excel(temp_xlsx, header=14, dtype=str)
        finally:
            # Cleanup temp file
            if os.path.exists(temp_xlsx):
                os.remove(temp_xlsx)
                
        # 3. Clean Columns
        # Drop Column D (4th column, index 3)
        if len(df.columns) > 3:
            df.drop(df.columns[3], axis=1, inplace=True)
            
        # Remove last 2 rows (unimportant footer)
        if len(df) > 2:
            df = df.iloc[:-2]
            
        return df


    @staticmethod
    def insert_clean_data(handler, df, sheet_name="RAW_CLEANED"):
        """
        Writes the cleaned DataFrame to a new sheet in the workbook.
        Applies formatting: Col S width 33, Row Height 180.
        Adds Formula in Col S pointing to Col R.
        """
        wb = handler.wb
        if sheet_name in wb.sheetnames:
            # If exists, maybe clear it or use it? Let's overwrite/recreate
            del wb[sheet_name]
        
        ws = wb.create_sheet(sheet_name)
        
        from openpyxl.styles import Font, Alignment, PatternFill
        from openpyxl.utils import get_column_letter
        from datetime import datetime
        import pandas as pd
        
        # Write Header & Data
        # Apply Styling: Calibri, 7, Center, Middle, Uppercase, Wrap Text
        # Header Color: Text #333399, White Background (Default)
        base_font = Font(name='Calibri', size=7)
        header_font = Font(name='Calibri', size=7, bold=True, color="333399") 
        # No fill needed if white is desired as default
        
        align_style = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # Identify Status Date Column Index
        # We need to scan headers first. or rely on DF columns.
        status_date_col_idx = None
        for i, col_name in enumerate(df.columns, 1):
            if "Status Date" in str(col_name):
                status_date_col_idx = i
                break
        
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                
                # Header Styling (Row 1)
                if r_idx == 1:
                    cell.value = str(value).upper()
                    cell.font = header_font
                    # cell.fill = header_fill # Removed
                else:
                    # Data Rows
                    # Special Handling for Status Date
                    if c_idx == status_date_col_idx:
                        # Value came from df as string usually. Try to parse it using pandas or dateutil
                        # User wants: Dec 14, 2025, 5:23PM
                        # Input might be ISO: 2025-12-14 17:23:42.062000
                        try:
                            if value and str(value).strip():
                                dt_val = pd.to_datetime(value)
                                cell.value = dt_val
                                cell.number_format = "d mmm, yyyy, h:mm AM/PM"
                            else:
                                cell.value = value
                        except:
                            cell.value = value
                    else:
                         # Convert to uppercase if string
                        if isinstance(value, str):
                            cell.value = value.upper()
                        else:
                            cell.value = value
                            
                    cell.font = base_font
                    
                cell.alignment = align_style
                
        # Apply Logic:
        # User: "now in column r ... links. put image formula in column s"
        # DYNAMIC SEARCH for URL Column to be safe
        link_col_idx = 18 # Default fallback (Col R)
        
        # Scan headers in row 1
        for col in range(1, ws.max_column + 1):
            val = str(ws.cell(1, col).value).upper()
            if "URL" in val and "ATTACH" in val:
                link_col_idx = col
                break
        
        img_col_idx = link_col_idx + 1
        
        # 1. Add "IMAGES" Header
        ws.cell(row=1, column=img_col_idx).value = "IMAGES"
        ws.cell(row=1, column=img_col_idx).font = header_font
        ws.cell(row=1, column=img_col_idx).alignment = align_style
        
        # 2. Add Images & Formatting
        ws.column_dimensions[get_column_letter(img_col_idx)].width = 33
        
        # Iterate rows (skip header row 1)
        for r in range(2, ws.max_row + 1):
            # Set Row Height
            ws.row_dimensions[r].height = 180
            
            # Get Link Value
            link_val = ws.cell(row=r, column=link_col_idx).value
            
            if link_val:
                formula = f'=_xlfn.IMAGE("{str(link_val).strip()}",,1)'
                cell = ws.cell(row=r, column=img_col_idx)
                cell.value = formula
                # Re-apply styles just in case
                cell.font = base_font
                cell.alignment = align_style
                
        # 3. Remove Columns T and U (Indices 20, 21)?
        # Only if strict requirement. Let's make it safer:
        # Hide columns after the Image column.
        
        # 4. Set Dimensions
        # Row 1 Height = 25
        ws.row_dimensions[1].height = 25
        
        # Column Widths (Indices 1 to 19)
        # 12,12,7,7,7,12,7,7,12,12,15,7,12,12,12,12,12,12,33
        widths = [12, 12, 7, 7, 7, 12, 7, 7, 12, 12, 15, 7, 12, 12, 12, 12, 12, 12, 33]
        
        for i, width in enumerate(widths, 1):
            # Safety check if i exists
            col_letter = get_column_letter(i)
            ws.column_dimensions[col_letter].width = width
            
        # Ensure Image Column has width 33 (in case it fell outside 'widths' list)
        ws.column_dimensions[get_column_letter(img_col_idx)].width = 33
            
        # 5. Hide Unused Columns (T onwards) -> T is 20
        # Excel typically goes up to XFD (Column 16384).
        # We can hide 20 to 16384. But iterating is slow.
        # Efficient way: Group columns.
        ws.column_dimensions.group(get_column_letter(20), get_column_letter(16384), hidden=True)
        
        # 6. Hide Unused Rows
        # Current used rows = 1 (header) + len(df)
        last_used_row = 1 + len(df)
        # Hide from next row to ... arbitrary large number? Or just 200?
        # User said "hide all rows that come after it".
        # Hiding to 1048576 is overkill and can bloat file.
        # Usually checking default view is enough. But we can hide the next 1000.
        # Or better: Group from last_used+1 to max.
        
        # NOTE: OpenPyXL grouping for rows:
        # ws.row_dimensions.group(start, end, hidden=True)
        # Max rows in excel is 1048576. This might be safe if done via grouping.
        ws.row_dimensions.group(last_used_row + 1, 1048576, hidden=True)

        return sheet_name
