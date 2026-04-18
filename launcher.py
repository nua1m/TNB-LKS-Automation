import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import sys
import os
from pathlib import Path


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()
VERSION_FILE = APP_DIR / "VERSION"
VERSION = VERSION_FILE.read_text(encoding="utf-8").strip() if VERSION_FILE.exists() else "0.0.0"


def get_processor_command(data_path: str) -> list[str]:
    if getattr(sys, "frozen", False):
        processor_exe = APP_DIR / "processor.exe"
        return [str(processor_exe), data_path]

    main_script = APP_DIR / "main.py"
    return [sys.executable, str(main_script), data_path]


def get_updater_command() -> list[str]:
    if getattr(sys, "frozen", False):
        return [str(APP_DIR / "updater.exe"), "--check-only"]

    return [sys.executable, str(APP_DIR / "updater.py"), "--check-only"]

def select_file():
    file_path = filedialog.askopenfilename(
        title="Select LKS Data File",
        filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
    )
    if file_path:
        entry_var.set(file_path)

def run_script():
    data_path = entry_var.get().strip()
    if not data_path:
        messagebox.showwarning("No File Selected", "Please select a data file first.")
        return

    if not os.path.exists(data_path):
        messagebox.showerror("File Not Found", f"Could not find:\n{data_path}")
        return

    try:
        result = subprocess.run(
            get_processor_command(data_path),
            capture_output=False,
            text=True,
            cwd=str(APP_DIR),
            env={**os.environ, "PYTHONIOENCODING": "utf-8"}
        )
        if result.returncode == 0:
            messagebox.showinfo("Success", "Processing completed!")
        else:
            messagebox.showerror("Error", f"Script exited with code {result.returncode}")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def check_for_updates():
    try:
        subprocess.run(
            get_updater_command(),
            check=False,
            cwd=str(APP_DIR),
        )
    except Exception as exc:
        messagebox.showerror("Update Check Failed", str(exc))


root = tk.Tk()
root.title(f"TNB LKS Automation v{VERSION}")
root.geometry("600x150")
root.resizable(False, False)

tk.Label(root, text=f"Version {VERSION}", fg="#666666").pack(pady=(10, 0))
tk.Label(root, text="Select LKS Data File:").pack(pady=(15, 5))

frame = tk.Frame(root)
frame.pack(pady=(0, 15), padx=15, fill="x")

entry_var = tk.StringVar()
entry = tk.Entry(frame, textvariable=entry_var)
entry.pack(side="left", fill="x", expand=True)

tk.Button(frame, text="Browse...", command=select_file).pack(side="right", padx=(5, 0))

tk.Button(root, text="RUN", command=run_script, bg="#4CAF50", fg="white", width=15).pack(pady=(0, 15))
tk.Button(root, text="Check Updates", command=check_for_updates, width=15).pack(pady=(0, 10))

root.mainloop()
