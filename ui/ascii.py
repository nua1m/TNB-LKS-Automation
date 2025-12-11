# ui/ascii.py
from ui.colors import PINK, RESET
import shutil

TITLE = r"""
                                                              
██     ██ ▄█▀ ▄█████   ▄▄ ▄▄ ▄██   ████▄ 
██     ████   ▀▀▀▄▄▄   ██▄██  ██    ▄▄██ 
██████ ██ ▀█▄ █████▀    ▀█▀   ██ ▄ ▄▄▄█▀ 
                                         
"""

def terminal_width():
    try:
        return shutil.get_terminal_size().columns
    except:
        return 80

def show_title():
    cols = terminal_width()
    for line in TITLE.splitlines():
        print(PINK + line.center(cols) + RESET)
    print()
