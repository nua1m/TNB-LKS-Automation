# core/so_utils.py — shared helpers


def clean_so(value):
    """Normalize SO values into a consistent string (no .0, trimmed)."""
    if value is None:
        return ""
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def is_missing(value) -> bool:
    """True if value is effectively empty."""
    if value is None:
        return True
    return str(value).strip() == ""
