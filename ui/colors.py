# ui/colors.py
# Fully-patched VSCode-style color set

import os
import sys

# Detect if terminal supports ANSI
SUPPORTS_ANSI = sys.platform != "win32" or "ANSICON" in os.environ or "WT_SESSION" in os.environ

def c(code: str) -> str:
    """Return ANSI code only if supported."""
    return code if SUPPORTS_ANSI else ""

RESET = c("\033[0m")

# Text styles
BOLD = c("\033[1m")
DIM  = c("\033[2m")

# Classic colors
RED     = c("\033[31m")
GREEN   = c("\033[32m")
YELLOW  = c("\033[33m")
BLUE    = c("\033[34m")
MAGENTA = c("\033[35m")
CYAN    = c("\033[36m")

# Bright colors
BRIGHT_RED     = c("\033[91m")
BRIGHT_GREEN   = c("\033[92m")
BRIGHT_YELLOW  = c("\033[93m")
BRIGHT_BLUE    = c("\033[94m")
BRIGHT_MAGENTA = c("\033[95m")
BRIGHT_CYAN    = c("\033[96m")

# Extra colors you requested
PINK = BRIGHT_MAGENTA   # alias
PURPLE = MAGENTA
