import customtkinter as ctk
from tkinter import filedialog
import sys
import threading
import os
from pathlib import Path

# Import the core logic
from main import run_process

# Configuration
ctk.set_appearance_mode("System")  # Modes: system (default), light, dark
ctk.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green

class RedirectText:
    """Redirects stdout/stderr to a text widget."""
    def __init__(self, text_widget):
        self.output = text_widget

    def write(self, string):
        self.output.insert("end", string)
        self.output.see("end")

    def flush(self):
        pass

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("TNB LKS Automation v1.4")
        self.geometry("700x550")
        
        # Grid Configuration
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1) # Log area expands

        # HEADER
        self.header = ctk.CTkLabel(self, text="TNB LKS Automation", font=ctk.CTkFont(size=20, weight="bold"))
        self.header.grid(row=0, column=0, padx=20, pady=(20, 10))

        # INPUT FRAME
        self.input_frame = ctk.CTkFrame(self)
        self.input_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        
        # Raw Data Input
        self.lbl_data = ctk.CTkLabel(self.input_frame, text="Raw Data File:")
        self.lbl_data.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.entry_data = ctk.CTkEntry(self.input_frame, width=400, placeholder_text="Select .xlsx or legacy .xls file...")
        self.entry_data.grid(row=0, column=1, padx=10, pady=10)
        
        self.btn_data = ctk.CTkButton(self.input_frame, text="Browse", width=80, command=self.browse_data)
        self.btn_data.grid(row=0, column=2, padx=10, pady=10)

        # Target File Input (Optional)
        self.lbl_target = ctk.CTkLabel(self.input_frame, text="Master File (Opt):")
        self.lbl_target.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        
        self.entry_target = ctk.CTkEntry(self.input_frame, width=400, placeholder_text="Optional: Select LKS file to append to...")
        self.entry_target.grid(row=1, column=1, padx=10, pady=10)
        
        self.btn_target = ctk.CTkButton(self.input_frame, text="Browse", width=80, command=self.browse_target)
        self.btn_target.grid(row=1, column=2, padx=10, pady=10)

        # ACTION BUTTON
        self.btn_start = ctk.CTkButton(self, text="START PROCESSING", fg_color="green", height=40, font=ctk.CTkFont(size=15, weight="bold"), command=self.start_thread)
        self.btn_start.grid(row=2, column=0, padx=20, pady=10, sticky="ew")

        # LOG AREA
        self.log_box = ctk.CTkTextbox(self, width=600, font=("Consolas", 12))
        self.log_box.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="nsew")
        
        # Redirect stdout
        sys.stdout = RedirectText(self.log_box)
        sys.stderr = RedirectText(self.log_box)
        
        print("Ready. Select files to begin.")

    def browse_data(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")])
        if filename:
            self.entry_data.delete(0, "end")
            self.entry_data.insert(0, filename)

    def browse_target(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsm *.xlsx")])
        if filename:
            self.entry_target.delete(0, "end")
            self.entry_target.insert(0, filename)

    def start_thread(self):
        # Disable button
        self.btn_start.configure(state="disabled", text="PROCESSING...")
        
        # Start in thread
        threading.Thread(target=self.run_logic, daemon=True).start()

    def run_logic(self):
        data_path = self.entry_data.get()
        target_path = self.entry_target.get()
        
        if not data_path:
            print("Error: Please select a Data file.")
            self.restore_button()
            return
            
        if not target_path:
            target_path = None
        
        try:
            print("-" * 50)
            print("Starting Automation...")
            print("-" * 50)
            
            run_process(data_path, target_path)
            
            print("\n" + "=" * 50)
            print("DONE! You can close this window.")
            print("=" * 50)
        except Exception as e:
            print(f"\nCRITICAL ERROR: {e}")
            import traceback
            traceback.print_exc()
        finally:
            self.restore_button()

    def restore_button(self):
        self.btn_start.configure(state="normal", text="START PROCESSING")

if __name__ == "__main__":
    app = App()
    app.mainloop()
