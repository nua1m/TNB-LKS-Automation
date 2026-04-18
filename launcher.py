import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import sys
import os
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent
VERSION_FILE = APP_DIR / "VERSION"
VERSION = VERSION_FILE.read_text(encoding="utf-8").strip() if VERSION_FILE.exists() else "0.0.0"

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

    script_dir = os.path.dirname(os.path.abspath(__file__))
    main_script = os.path.join(script_dir, "main.py")

    try:
        result = subprocess.run(
            [sys.executable, main_script, data_path],
            capture_output=False,
            text=True,
            cwd=script_dir,
            env={**os.environ, "PYTHONIOENCODING": "utf-8"}
        )
        if result.returncode == 0:
            messagebox.showinfo("Success", "Processing completed!")
        else:
            messagebox.showerror("Error", f"Script exited with code {result.returncode}")
    except Exception as e:
        messagebox.showerror("Error", str(e))


def check_for_updates():
    updater_script = APP_DIR / "updater.py"
    try:
        subprocess.run(
            [sys.executable, str(updater_script), "--check-only"],
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
