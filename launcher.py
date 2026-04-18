import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from config import DEFAULT_TEMPLATE_PATH
from main import run_process

APP_DIR = Path(__file__).resolve().parent
VERSION_FILE = APP_DIR / "VERSION"
VERSION = VERSION_FILE.read_text(encoding="utf-8").strip() if VERSION_FILE.exists() else "0.0.0"


class LauncherApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"TNB LKS Automation v{VERSION}")
        self.root.geometry("760x520")
        self.root.minsize(720, 500)

        self.events: queue.Queue = queue.Queue()
        self.processing = False
        self.last_output_path: str | None = None
        self.result_folder: str = str(APP_DIR)

        self.file_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        self.update_var = tk.StringVar(value="Updates: Check manually")

        self._build_ui()
        self.root.after(100, self._drain_events)

    def _build_ui(self) -> None:
        self.root.configure(bg="#f5f7fb")

        container = tk.Frame(self.root, bg="#f5f7fb", padx=18, pady=18)
        container.pack(fill="both", expand=True)

        header = tk.Frame(container, bg="#f5f7fb")
        header.pack(fill="x", pady=(0, 12))

        tk.Label(
            header,
            text="TNB LKS Automation",
            font=("Segoe UI", 18, "bold"),
            bg="#f5f7fb",
            fg="#17324d",
        ).pack(anchor="w")
        tk.Label(
            header,
            text=f"Version {VERSION}",
            font=("Segoe UI", 10),
            bg="#f5f7fb",
            fg="#66758a",
        ).pack(anchor="w")

        status_card = tk.Frame(container, bg="white", bd=1, relief="solid", padx=14, pady=10)
        status_card.pack(fill="x", pady=(0, 12))
        tk.Label(status_card, text="Status", font=("Segoe UI", 10, "bold"), bg="white", fg="#17324d").pack(anchor="w")
        tk.Label(status_card, textvariable=self.status_var, font=("Segoe UI", 11), bg="white", fg="#1f2937").pack(anchor="w", pady=(4, 0))
        tk.Label(status_card, textvariable=self.update_var, font=("Segoe UI", 10), bg="white", fg="#66758a").pack(anchor="w", pady=(2, 0))

        task_card = tk.Frame(container, bg="white", bd=1, relief="solid", padx=14, pady=14)
        task_card.pack(fill="x", pady=(0, 12))
        tk.Label(task_card, text="Input Excel File", font=("Segoe UI", 10, "bold"), bg="white", fg="#17324d").pack(anchor="w")
        tk.Label(
            task_card,
            text="Choose the technician Excel file (.xls or .xlsx). The app will generate an LKS result workbook in the same folder.",
            font=("Segoe UI", 9),
            bg="white",
            fg="#66758a",
            justify="left",
            wraplength=680,
        ).pack(anchor="w", pady=(4, 10))

        file_row = tk.Frame(task_card, bg="white")
        file_row.pack(fill="x")
        self.file_entry = tk.Entry(file_row, textvariable=self.file_var, font=("Segoe UI", 10))
        self.file_entry.pack(side="left", fill="x", expand=True)
        self.browse_button = tk.Button(file_row, text="Browse", command=self.select_file, width=12)
        self.browse_button.pack(side="left", padx=(8, 0))

        action_row = tk.Frame(task_card, bg="white")
        action_row.pack(fill="x", pady=(14, 0))
        self.process_button = tk.Button(
            action_row,
            text="Process LKS",
            command=self.start_processing,
            bg="#1f8f4e",
            fg="white",
            activebackground="#16713d",
            activeforeground="white",
            relief="flat",
            padx=20,
            pady=8,
        )
        self.process_button.pack(side="left")
        self.update_button = tk.Button(action_row, text="Check Updates", command=self.check_for_updates, width=16)
        self.update_button.pack(side="left", padx=(8, 0))
        self.output_button = tk.Button(action_row, text="Open Result Folder", command=self.open_result_folder, width=18)
        self.output_button.pack(side="left", padx=(8, 0))

        log_card = tk.Frame(container, bg="white", bd=1, relief="solid", padx=14, pady=14)
        log_card.pack(fill="both", expand=True)
        tk.Label(log_card, text="Run Log", font=("Segoe UI", 10, "bold"), bg="white", fg="#17324d").pack(anchor="w")
        tk.Label(
            log_card,
            text="Processing messages and review notes will appear here.",
            font=("Segoe UI", 9),
            bg="white",
            fg="#66758a",
        ).pack(anchor="w", pady=(4, 8))

        self.log_text = ScrolledText(log_card, height=14, font=("Consolas", 9), wrap="word", bg="#fbfcfe", fg="#1f2937")
        self.log_text.pack(fill="both", expand=True)
        self.log_text.configure(state="disabled")

    def append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def set_processing_state(self, active: bool) -> None:
        self.processing = active
        state = "disabled" if active else "normal"
        self.process_button.configure(state=state)
        self.browse_button.configure(state=state)
        self.update_button.configure(state=state)
        self.file_entry.configure(state=state)

    def select_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if file_path:
            self.file_var.set(file_path)

    def check_for_updates(self) -> None:
        updater_script = APP_DIR / "updater.py"
        try:
            subprocess.run([sys.executable, str(updater_script), "--check-only"], check=False, cwd=str(APP_DIR))
        except Exception as exc:
            messagebox.showerror("Update Check Failed", str(exc))

    def open_result_folder(self) -> None:
        target = self.result_folder if os.path.isdir(self.result_folder) else str(APP_DIR)
        try:
            os.startfile(target)
        except Exception as exc:
            messagebox.showerror("Open Folder Failed", str(exc))

    def _queue_log(self, message: str) -> None:
        self.events.put(("log", message))

    def _queue_status(self, message: str) -> None:
        self.events.put(("status", message))

    def _confirm_append(self, existing_count: int, new_count: int) -> bool:
        event = threading.Event()
        result: dict[str, bool] = {}
        self.events.put(("confirm_append", existing_count, new_count, event, result))
        event.wait()
        return result.get("value", False)

    def start_processing(self) -> None:
        data_path = self.file_var.get().strip()
        if not data_path:
            messagebox.showwarning("Input File Required", "Choose the Excel file you want to process first.")
            return

        if not os.path.exists(data_path):
            messagebox.showerror("File Not Found", f"Could not find:\n{data_path}")
            return

        template_path = Path(DEFAULT_TEMPLATE_PATH).resolve()
        if not template_path.exists():
            messagebox.showerror("Template Not Found", f"Could not find the template file:\n{template_path}")
            return

        self.last_output_path = None
        self.result_folder = str(Path(data_path).resolve().parent)
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self.append_log("Starting LKS processing...")
        self.status_var.set("Processing workbook... please wait")
        self.set_processing_state(True)

        worker = threading.Thread(
            target=self._process_worker,
            args=(Path(data_path).resolve(), template_path),
            daemon=True,
        )
        worker.start()

    def _process_worker(self, data_path: Path, template_path: Path) -> None:
        try:
            result = run_process(
                data_path,
                template_path,
                log_fn=self._queue_log,
                confirm_append_fn=self._confirm_append,
                status_fn=self._queue_status,
                show_cli_summary=False,
            )
            self.events.put(("done", result))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def _drain_events(self) -> None:
        while True:
            try:
                item = self.events.get_nowait()
            except queue.Empty:
                break

            kind = item[0]
            if kind == "log":
                self.append_log(item[1])
            elif kind == "status":
                self.status_var.set(item[1])
            elif kind == "confirm_append":
                _, existing_count, new_count, event, result = item
                answer = messagebox.askyesno(
                    "Append New SOs",
                    f"The template already has {existing_count} SOs.\n\n{new_count} new SOs will be added.\n\nChoose Yes to continue. No changes will be saved if you choose No.",
                )
                result["value"] = answer
                event.set()
            elif kind == "done":
                result = item[1]
                self.set_processing_state(False)
                self.last_output_path = result.get("output_path")
                self.result_folder = str(Path(self.last_output_path).parent) if self.last_output_path else self.result_folder
                if result.get("aborted"):
                    self.status_var.set("Processing cancelled")
                    self.append_log("Run cancelled before saving changes.")
                else:
                    self.status_var.set("Completed")
                    if self.last_output_path:
                        self.append_log(f"Saved to: {self.last_output_path}")
                    summary = result.get("summary", {})
                    if summary:
                        self.append_log("")
                        self.append_log("Summary:")
                        for key, value in summary.items():
                            self.append_log(f"- {key}: {value}")
                    messagebox.showinfo(
                        "Processing Complete",
                        f"LKS processing completed successfully.\n\nSaved to:\n{self.last_output_path}",
                    )
            elif kind == "error":
                self.set_processing_state(False)
                self.status_var.set("Failed")
                self.append_log(f"Error: {item[1]}")
                messagebox.showerror(
                    "Processing Failed",
                    "The file could not be processed.\n\nCheck the run log for details.",
                )

        self.root.after(100, self._drain_events)


def main() -> None:
    root = tk.Tk()
    LauncherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
