from datetime import datetime

class DateEngine:
    @staticmethod
    @staticmethod
    def parse_datetime(date_str):
        """Attempts to parse a date string into a datetime object."""
        if not date_str: return None
        if isinstance(date_str, datetime): return date_str

        s = str(date_str).strip()
        # Handle "Jan 01, 2025, 10:00 AM" format
        if "," in s:
            # Try full datetime first: "Nov 12, 2025, 1:24 PM"
            try:
                # Cleaning up potential messiness if needed, but strptime is specific
                # Adjust format to match "Nov 12, 2025, 1:24 PM"
                return datetime.strptime(s, "%b %d, %Y, %I:%M %p")
            except ValueError:
                pass
            
            # Fallback to date only if time fails but comma exists
            parts = s.split(",")
            if len(parts) >= 2:
                s_date = parts[0].strip() + ", " + parts[1].strip()
                try:
                    return datetime.strptime(s_date, "%b %d, %Y")
                except ValueError:
                    pass

        # Handle "4 Dec 2025" (OCR format)
        try:
             return datetime.strptime(s, "%d %b %Y")
        except ValueError:
            pass

        # Handle ISO
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d"):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                pass

        return None

    @staticmethod
    def parse_date(date_str):
        """Attempts to parse a date string into a date object."""
        if isinstance(date_str, datetime):
            return date_str.date()
            
        dt = DateEngine.parse_datetime(date_str)
        if dt:
            return dt.date()
        return None

    @staticmethod
    def calculate(status_date_str, ocr_date_str=None):
        """
        Applies Business Logic:
        1. Effective Date = OCR if present, else Status.
        2. Diskon = (OCR is present AND OCR != Status).
        3. Hari = "Hujung Minggu" if Effective is Sunday.
        4. Remarks 2 = "TASK FORCE" if Effective is Saturday.
        5. Remarks 1 = "TECO LEWAT..." if Diskon.
        """
        status_date = DateEngine.parse_date(status_date_str)
        ocr_date = DateEngine.parse_date(ocr_date_str)

        # Default to Status Date
        effective_date = status_date
        is_diskon = False

        # Apply OCR override if valid
        if ocr_date:
            if status_date != ocr_date:
                effective_date = ocr_date
                is_diskon = True
        
        # Derived Fields
        hari_field = "Hari Biasa"
        remarks_1 = ""
        remarks_2 = ""

        if effective_date:
            # Sunday Rule
            if effective_date.weekday() == 6:
                hari_field = "Hujung Minggu"
            
            # Saturday Rule -> Task Force
            if effective_date.weekday() == 5:
                remarks_2 = "TASK FORCE"

        if is_diskon:
            remarks_1 = f"TECO LEWAT SEBAB DISKON ({effective_date.strftime('%d %b %Y')})"

        return {
            "effective_date": effective_date,
            "is_diskon": is_diskon,
            "hari": hari_field,
            "remarks_1": remarks_1,
            "remarks_2": remarks_2
        }
