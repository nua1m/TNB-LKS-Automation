from datetime import datetime

class DateEngine:
    @staticmethod
    def parse_date(date_str):
        """Attempts to parse a date string into a date object."""
        if not date_str: return None
        s = str(date_str).strip()
        # Handle "Jan 01, 2025, 10:00 AM" format (common in this app)
        if "," in s:
            parts = s.split(",")
            s = parts[0].strip() + ", " + parts[1].strip()
            try:
                return datetime.strptime(s, "%b %d, %Y").date()
            except ValueError:
                pass
        
        # Handle "4 Dec 2025" (OCR format)
        try:
             return datetime.strptime(s, "%d %b %Y").date()
        except ValueError:
            pass

        # Handle ISO
        try:
            return datetime.strptime(s, "%Y-%m-%d").date()
        except ValueError:
            pass

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
