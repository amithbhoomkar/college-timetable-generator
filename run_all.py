#!/usr/bin/env python3
"""
================================================================================
MASTER ORCHESTRATION PIPELINE: COLLEGE TIMETABLE SYSTEM
================================================================================
This script executes the complete end-to-end timetable generation and export
pipeline in proper dependency order:
  Step 1: main.py                   -> Generate master schedule using CP-SAT solver
  Step 2: export_workload_sheet.py   -> Calculate workload & inject analytics sheet
  Step 3: export_separate_timetables.py -> Generate separate Excel workbooks
  Step 4: export_text_and_csv.py     -> Generate CSVs and Markdown text views
  Step 5: generate_html_viewer.py    -> Build self-contained HTML interactive viewer
================================================================================
"""

import sys
import time
import subprocess
import os

STEPS = [
    ("Step 1: Constraint Solving & Master Timetable Generation", "main.py"),
    ("Step 2: Faculty Workload Analytics & Distribution Sheet", "export_workload_sheet.py"),
    ("Step 3: Export Individual Excel Workbooks (Sections/Staff/Rooms/Depts)", "export_separate_timetables.py"),
    ("Step 4: Export CSV Datasets and Markdown Text Tables", "export_text_and_csv.py"),
    ("Step 5: Generate Interactive Web Dashboard (Timetable_Viewer.html)", "generate_html_viewer.py")
]

def run_pipeline():
    print("=" * 80)
    print("   AUTOMATED TIMETABLE GENERATION & EXPORT PIPELINE")
    print("=" * 80)
    start_total = time.time()

    for idx, (title, script_name) in enumerate(STEPS, 1):
        print(f"\n[{idx}/{len(STEPS)}] Running {title} ({script_name})...")
        print("-" * 80)
        step_start = time.time()
        
        result = subprocess.run([sys.executable, script_name], capture_output=False)
        if result.returncode != 0:
            print(f"\n[ERROR] Pipeline aborted: {script_name} failed with exit code {result.returncode}!")
            sys.exit(result.returncode)
            
        elapsed = round(time.time() - step_start, 2)
        print(f"--> [OK] Completed {script_name} in {elapsed}s")

    total_time = round(time.time() - start_total, 2)
    print("\n" + "=" * 80)
    print(f"   ALL PIPELINE STEPS COMPLETED SUCCESSFULLY IN {total_time}s")
    print("=" * 80)
    print("Generated Artifacts:")
    print("  - Master Workbook    : Complete_Timetable.xlsx")
    print("  - Workload Workbook  : Teacher_Workload_Sheet.xlsx")
    print("  - Interactive Viewer : Timetable_Viewer.html")
    print("  - Separate Workbooks : Timetable_Exports/ (Sections, Staff, Rooms, Departments)")
    print("  - CSV Files          : Timetable_CSVs/")
    print("  - Markdown Reports   : Timetable_Text_Views/ and ALL_TIMETABLES.md")
    print("=" * 80)

if __name__ == "__main__":
    run_pipeline()
