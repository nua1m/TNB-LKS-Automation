# ui/components.py

import shutil
from ui.colors import CYAN, PURPLE, PINK, BLUE, RESET

# Spinner frames — consistent with your image pipeline
SPIN = ["⠋","⠙","⠹","⠸","⠼","⠴","⠦","⠧","⠇","⠏"]


# ========== Helpers ==========

def terminal_width():
    try:
        return shutil.get_terminal_size().columns
    except:
        return 80


# ========== Progress Bar Renderer ==========

def progress_bar(current, total, width=40, color=CYAN):
    ratio = current / total
    filled = int(width * ratio)
    bar = "█" * filled + "░" * (width - filled)
    return f"{color}{bar}{RESET}"


def step_progress(label, i, total, extra="", spinner_i=0):
    """
    Unified step progress:
    ⠴ [BUILD] ███████████░░░░░ 32/421  30%  Processing Images (32/562)
    """
    spin = SPIN[spinner_i % len(SPIN)]
    bar = progress_bar(i, total, width=42, color=PURPLE)
    percent = int((i / total) * 100)

    print(
        f"\r{spin} [{label}] {bar} {i}/{total}  {percent:3d}%  {extra}",
        end="",
        flush=True
    )


# ========== Summary Block ==========

def summary_block(metrics: dict, out_file: str):
    """Clean display of final results (no tables)."""

    cols = terminal_width()
    line = "─" * cols

    print("\n")
    print(PINK + line + RESET)
    print(PURPLE + "SUMMARY".center(cols) + RESET)
    print(PINK + line + RESET)

    for key, val in metrics.items():
        print(f" • {BLUE}{key:<22}{RESET}: {val}")

    print("\n✔ Output written to:")
    print(f"   {PINK}{out_file}{RESET}")

    print(PINK + line + RESET)
