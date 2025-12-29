"""
Keyboard Automation - Copy NEW METER Images Only

Uses identical logic to the working ticket (card) script.
"""

import pyautogui
import time

# Safety
pyautogui.FAILSAFE = True
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
    Identical to keyboard_copy_images.py logic.
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
    """Process new_meter only for current row."""
    copy_image('new')  # This will use 'new_meter' keyword
    wait(SHORT)
    
    # Move to next row
    key('down')
    wait(SHORT)

def main():
    print("=" * 50)
    print("Image Copy - NEW METER ONLY → Column F")
    print("=" * 50)
    print()
    print("SETUP:")
    print("1. ATTACHMENT sheet active, cursor on Column B")
    print("2. DATA sheet is NEXT sheet (Ctrl+PageDown)")
    print()
    
    num_rows = int(input("How many rows? "))
    
    print()
    print("Press ENTER, then FOCUS EXCEL within 3 seconds!")
    input()
    
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
    
    print("\nDONE!")

if __name__ == "__main__":
    main()
