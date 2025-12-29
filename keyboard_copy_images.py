"""
Keyboard Automation - Copy Ticket & New Meter Images

PURPOSE: Automate copying images from DATA sheet to LKS ATTACHMENT sheet
         for TICKET (card) and NEW METER (new_meter) only.

DATA SHEET LAYOUT:
- Column A: SO Number (with blanks for subsequent rows of same SO)
- Column B: Image
- Column C: URL (contains 'card' or 'new_meter' keywords)

LKS ATTACHMENT LAYOUT:
- Column B: SO Number
- Column E: Ticket image (card)
- Column F: New Meter image (new_meter)

PREREQUISITES:
1. Both sheets in same Excel workbook
2. Cursor on first SO to process in ATTACHMENT sheet (Column B)
3. DATA sheet accessible via Ctrl+PageDown
"""

import pyautogui
import time

# Safety
pyautogui.FAILSAFE = True  # Move mouse to top-left corner to ABORT
pyautogui.PAUSE = 0.03

# Delays - BALANCED MODE
TINY = 0.08
SHORT = 0.2
MEDIUM = 0.3
LONG = 0.5

def wait(sec=SHORT):
    time.sleep(sec)

def key(k, times=1):
    for _ in range(times):
        pyautogui.press(k)
        wait(TINY)

def hotkey(*keys):
    pyautogui.hotkey(*keys)
    wait(SHORT)

def copy_image(image_type):
    """
    image_type: 'ticket' or 'new'
    
    Workflow:
    1. Copy SO from current cell (Column B in ATTACHMENT)
    2. Go to DATA sheet
    3. Ctrl+F, paste SO, Enter to find
    4. Close Find
    5. Ctrl+F again, search for keyword (card or new_meter)
    6. Close Find - now on the right row
    7. Go left to Column B (the image)
    8. Copy image
    9. Go back to ATTACHMENT
    10. Navigate to correct column (E or F)
    11. Paste
    12. Return to Column B
    """
    
    keyword = 'card' if image_type == 'ticket' else 'new_meter'
    target_col_offset = 3 if image_type == 'ticket' else 4  # E=3, F=4 from B
    
    print(f"    Copying {image_type}...")
    
    # 1. Copy SO
    hotkey('ctrl', 'c')
    wait(MEDIUM)
    
    # 2. Go to DATA sheet
    hotkey('ctrl', 'pagedown')
    wait(LONG)
    
    # 3. Find SO
    hotkey('ctrl', 'f')
    wait(MEDIUM)
    hotkey('ctrl', 'v')  # Paste SO
    wait(SHORT)
    key('enter')  # Find
    wait(MEDIUM)
    key('escape')  # Close Find
    wait(SHORT)
    
    # 4. Now search for the keyword (card or new_meter) within this area
    hotkey('ctrl', 'f')
    wait(MEDIUM)
    
    # Type the keyword
    pyautogui.typewrite(keyword, interval=0.03)
    wait(SHORT)
    key('enter')  # Find
    wait(MEDIUM)
    key('escape')  # Close Find
    wait(SHORT)
    
    # 5. We should now be on the URL column (C) with the keyword
    # Go LEFT to Column B where the image is
    key('left', times=1)  # C -> B
    wait(SHORT)
    
    # 6. Copy the image cell
    hotkey('ctrl', 'c')
    wait(MEDIUM)
    
    # 7. Go back to ATTACHMENT sheet
    hotkey('ctrl', 'pageup')
    wait(LONG)
    
    # 8. Navigate to target column (E or F)
    key('right', times=target_col_offset)
    wait(SHORT)
    
    # 9. Paste with Ctrl+V
    hotkey('ctrl', 'v')
    wait(MEDIUM)
    
    # 10. Return to Column B (SO column)
    key('left', times=target_col_offset)
    wait(SHORT)
    
    print(f"    ✓ {image_type} done!")

def process_row():
    """Process ticket only for current row."""
    copy_image('ticket')
    wait(SHORT)
    
    # Move to next row
    key('down')
    wait(SHORT)

def main():
    print("=" * 50)
    print("Image Copy Automation - Ticket & New Meter")
    print("=" * 50)
    print()
    print("SETUP:")
    print("1. Excel open with ATTACHMENT sheet active")
    print("2. Cursor on FIRST SO to process (Column B)")
    print("3. DATA sheet is the NEXT sheet (Ctrl+PageDown)")
    print()
    
    num_rows = int(input("How many rows to process? "))
    
    print()
    print("Press ENTER when ready (then focus Excel IMMEDIATELY)")
    input()
    
    print("Starting in 3 seconds... FOCUS EXCEL NOW!")
    for i in [3, 2, 1]:
        print(f"  {i}...")
        time.sleep(1)
    
    print()
    for row in range(num_rows):
        print(f"Row {row + 1}/{num_rows}")
        try:
            process_row()
        except KeyboardInterrupt:
            print("\n⚠️ Aborted!")
            break
    
    print()
    print("=" * 50)
    print("DONE!")
    print("=" * 50)

if __name__ == "__main__":
    main()
