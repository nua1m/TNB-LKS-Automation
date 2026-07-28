from datetime import datetime

class DateEngine:
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

        # Handle "4 Dec 2025" style dates
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
    def get_hari_field(status_date_str):
        """Returns the day-type label from the CLAIM status date."""
        status_date = DateEngine.parse_date(status_date_str)
        if status_date and status_date.weekday() == 6:
            return "Hujung Minggu"
        return "Hari Biasa"
