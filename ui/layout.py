# ui/layout.py

import os

def set_window_size(columns=110, rows=40):
    """Windows-only: sets terminal size safely."""
    try:
        os.system(f"mode con: cols={columns} lines={rows}")
    except:
        pass

def center_window():
    """Optional: attempt to center window."""
    pass  # Windows terminal doesn't easily support centering programmatically
