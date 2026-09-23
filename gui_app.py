#!/usr/bin/env python3
"""
================================================================================
COLLEGE TIMETABLE GENERATOR - DESKTOP CONTROL DASHBOARD (GUI)
================================================================================
A graphical interface for professors and administrators to:
  1. Trigger the Multi-Agent CP-SAT Solver pipeline with real-time logs
  2. Launch the Interactive Web Dashboard (Timetable_Viewer.html)
  3. Open Excel workbooks (Master and Workload sheets) directly
  4. Browse modular export directories (Sections, Staff, Rooms)
================================================================================
"""

import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import threading
import webbrowser
import os
import sys

class TimetableGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("College Timetable Control Center | ISE & CSBS")
        self.root.geometry("820x640")
        self.root.minsize(750, 550)

        # Base path
        self.base_dir = os.path.dirname(os.path.abspath(__file__))

        # Dark theme colors
        self.bg_color = "#0f172a"
        self.card_bg = "#1e293b"
        self.accent_blue = "#2563eb"
        self.accent_green = "#10b981"
        self.text_light = "#f8fafc"
        self.text_muted = "#94a3b8"
        self.border_color = "#334155"

        self.root.configure(bg=self.bg_color)
        self.is_running = False

        self.build_ui()

    def build_ui(self):
        # 1. Header Frame
        header = tk.Frame(self.root, bg=self.card_bg, padx=25, pady=18, highlightbackground=self.border_color, highlightthickness=1)
        header.pack(fill="x", padx=15, pady=(15, 10))

        title_lbl = tk.Label(
            header,
            text="🎓 College Timetable Control Center",
            font=("Segoe UI", 18, "bold"),
            bg=self.card_bg,
            fg=self.text_light
        )
        title_lbl.pack(anchor="w")

        sub_lbl = tk.Label(
            header,
            text="Multi-Agent Discrete Optimization (OR-Tools CP-SAT) | 13 Sections • 20 Faculty • 9 Rooms",
            font=("Segoe UI", 10),
            bg=self.card_bg,
            fg=self.text_muted
        )
        sub_lbl.pack(anchor="w", pady=(3, 0))

        # 2. Action Controls Frame
        controls = tk.Frame(self.root, bg=self.bg_color, padx=5, pady=5)
        controls.pack(fill="x", padx=15, pady=5)

        btn_font = ("Segoe UI", 10, "bold")

        # Run Pipeline Button
        self.run_btn = tk.Button(
            controls,
            text="▶ Run Complete Pipeline",
            font=btn_font,
            bg=self.accent_green,
            fg="white",
            activebackground="#059669",
            activeforeground="white",
            relief="flat",
            padx=16,
            pady=10,
            cursor="hand2",
            command=self.start_pipeline_thread
        )
        self.run_btn.grid(row=0, column=0, padx=(0, 8), pady=5)

        # Open Web Viewer
        web_btn = tk.Button(
            controls,
            text="🌐 Open Interactive Web UI",
            font=btn_font,
            bg=self.accent_blue,
            fg="white",
            activebackground="#1d4ed8",
            activeforeground="white",
            relief="flat",
            padx=14,
            pady=10,
            cursor="hand2",
            command=self.open_web_viewer
        )
        web_btn.grid(row=0, column=1, padx=8, pady=5)

        # Open Excel Master
        excel_btn = tk.Button(
            controls,
            text="📊 Master Excel",
            font=btn_font,
            bg="#3b82f6",
            fg="white",
            activebackground="#2563eb",
            activeforeground="white",
            relief="flat",
            padx=14,
            pady=10,
            cursor="hand2",
            command=self.open_excel_master
        )
        excel_btn.grid(row=0, column=2, padx=8, pady=5)

        # Open Workload Sheet
        workload_btn = tk.Button(
            controls,
            text="👥 Workload Sheet",
            font=btn_font,
            bg="#6366f1",
            fg="white",
            activebackground="#4f46e5",
            activeforeground="white",
            relief="flat",
            padx=14,
            pady=10,
            cursor="hand2",
            command=self.open_workload_sheet
        )
        workload_btn.grid(row=0, column=3, padx=8, pady=5)

        # Open Folder
        folder_btn = tk.Button(
            controls,
            text="📁 Browse Exports",
            font=btn_font,
            bg="#475569",
            fg="white",
            activebackground="#334155",
            activeforeground="white",
            relief="flat",
            padx=14,
            pady=10,
            cursor="hand2",
            command=self.open_exports_folder
        )
        folder_btn.grid(row=0, column=4, padx=(8, 0), pady=5)

        # 3. Progress Bar
        self.progress = ttk.Progressbar(self.root, mode="indeterminate")
        self.progress.pack(fill="x", padx=15, pady=(5, 5))

        # 4. Status and Log Console
        log_frame = tk.Frame(self.root, bg=self.card_bg, padx=15, pady=12, highlightbackground=self.border_color, highlightthickness=1)
        log_frame.pack(fill="both", expand=True, padx=15, pady=(5, 15))

        console_title = tk.Label(
            log_frame,
            text="System Execution & Solver Output Log:",
            font=("Segoe UI", 10, "bold"),
            bg=self.card_bg,
            fg=self.text_light
        )
        console_title.pack(anchor="w", pady=(0, 6))

        self.log_text = tk.Text(
            log_frame,
            bg="#090d16",
            fg="#e2e8f0",
            insertbackground="white",
            font=("Consolas", 9),
            wrap="word",
            relief="flat",
            padx=10,
            pady=10
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)

        # Initial greeting in log
        self.log("System initialized successfully.")
        self.log("Click '▶ Run Complete Pipeline' to re-solve constraints and generate all schedules.")
        self.log("Click '🌐 Open Interactive Web UI' to view the interactive dashboard in your browser.")

    def log(self, message):
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")

    def start_pipeline_thread(self):
        if self.is_running:
            return
        self.is_running = True
        self.run_btn.config(state="disabled", text="⏳ Processing...")
        self.progress.start(10)
        self.log("\n" + "=" * 60)
        self.log("Starting full generation & export pipeline...")
        self.log("=" * 60)

        threading.Thread(target=self.run_pipeline, daemon=True).start()

    def run_pipeline(self):
        try:
            cmd = [sys.executable, os.path.join(self.base_dir, "run_all.py")]
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                cwd=self.base_dir
            )

            for line in process.stdout:
                clean_line = line.rstrip()
                self.root.after(0, self.log, clean_line)

            process.wait()

            if process.returncode == 0:
                self.root.after(0, lambda: messagebox.showinfo("Success", "Timetable generation and exports completed successfully!"))
            else:
                self.root.after(0, lambda: messagebox.showerror("Execution Error", f"Process exited with error code {process.returncode}."))

        except Exception as e:
            self.root.after(0, self.log, f"[ERROR] {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, self.finish_pipeline)

    def finish_pipeline(self):
        self.is_running = False
        self.progress.stop()
        self.run_btn.config(state="normal", text="▶ Run Complete Pipeline")

    def open_web_viewer(self):
        html_path = os.path.join(self.base_dir, "Timetable_Viewer.html")
        if os.path.exists(html_path):
            webbrowser.open(f"file:///{os.path.abspath(html_path)}")
            self.log(f"Opened web viewer: {html_path}")
        else:
            messagebox.showwarning("File Missing", "Timetable_Viewer.html not found. Please run the pipeline first.")

    def open_excel_master(self):
        excel_path = os.path.join(self.base_dir, "Complete_Timetable.xlsx")
        if os.path.exists(excel_path):
            os.startfile(excel_path)
            self.log(f"Opened Excel workbook: {excel_path}")
        else:
            messagebox.showwarning("File Missing", "Complete_Timetable.xlsx not found.")

    def open_workload_sheet(self):
        sheet_path = os.path.join(self.base_dir, "Teacher_Workload_Sheet.xlsx")
        if os.path.exists(sheet_path):
            os.startfile(sheet_path)
            self.log(f"Opened Workload Sheet: {sheet_path}")
        else:
            messagebox.showwarning("File Missing", "Teacher_Workload_Sheet.xlsx not found.")

    def open_exports_folder(self):
        folder_path = os.path.join(self.base_dir, "Timetable_Exports")
        if os.path.exists(folder_path):
            os.startfile(folder_path)
            self.log(f"Opened folder: {folder_path}")
        else:
            os.startfile(self.base_dir)

if __name__ == "__main__":
    root = tk.Tk()
    app = TimetableGUI(root)
    root.mainloop()
