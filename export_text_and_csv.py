import openpyxl
import os
import csv

def generate_text_and_csv():
    wb = openpyxl.load_workbook("Complete_Timetable.xlsx")

    os.makedirs("Timetable_Text_Views", exist_ok=True)
    os.makedirs("Timetable_CSVs", exist_ok=True)

    section_sheets = [s for s in wb.sheetnames if s.startswith("ISE_S") or s.startswith("CSBS_S")]
    all_md_content = ["# COLLEGE TIMETABLES - ALL SECTIONS\n\n"]

    for s_name in section_sheets:
        ws = wb[s_name]

        # 1. Generate CSV
        csv_path = f"Timetable_CSVs/{s_name}.csv"
        with open(csv_path, "w", newline="", encoding="utf-8") as f_csv:
            writer = csv.writer(f_csv)
            for r in range(1, ws.max_row + 1):
                row_vals = [ws.cell(row=r, column=c).value or "" for c in range(1, ws.max_column + 1)]
                writer.writerow(row_vals)

        # 2. Generate Markdown table
        md_lines = [f"# Timetable: {s_name}\n\n"]
        headers = ["Time", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        md_lines.append("| " + " | ".join(headers) + " |")
        md_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")

        for r in range(4, ws.max_row + 1):
            time_slot = ws.cell(row=r, column=1).value
            if not time_slot:
                continue
            row_cells = [str(time_slot)]
            for c in range(2, 8):
                val = ws.cell(row=r, column=c).value or "FREE"
                # replace newlines with html breaks for markdown table cell
                val_str = str(val).replace("\n", "<br>")
                row_cells.append(val_str)
            md_lines.append("| " + " | ".join(row_cells) + " |")

        md_content = "\n".join(md_lines) + "\n\n"
        with open(f"Timetable_Text_Views/{s_name}.md", "w", encoding="utf-8") as f_md:
            f_md.write(md_content)

        all_md_content.append(md_content)
        all_md_content.append("\n---\n\n")

    # Save combined file
    with open("ALL_TIMETABLES.md", "w", encoding="utf-8") as f_all:
        f_all.write("\n".join(all_md_content))

    # Also Department Masters as Markdown & CSV
    for d_name in ["ISE_MASTER", "CSBS_MASTER"]:
        ws = wb[d_name]
        with open(f"Timetable_CSVs/{d_name}.csv", "w", newline="", encoding="utf-8") as f_csv:
            writer = csv.writer(f_csv)
            for r in range(1, ws.max_row + 1):
                row_vals = [ws.cell(row=r, column=c).value or "" for c in range(1, ws.max_column + 1)]
                writer.writerow(row_vals)

        md_lines = [f"# Master Timetable: {d_name}\n\n"]
        headers = ["Day", "Time", "Section", "Subject", "Type", "Teacher", "Room"]
        md_lines.append("| " + " | ".join(headers) + " |")
        md_lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
        for r in range(4, ws.max_row + 1):
            vals = [str(ws.cell(row=r, column=c).value or "") for c in range(1, 8)]
            if not vals[0]:
                continue
            md_lines.append("| " + " | ".join(vals) + " |")
        with open(f"Timetable_Text_Views/{d_name}.md", "w", encoding="utf-8") as f_md:
            f_md.write("\n".join(md_lines))

    # Staff timetable as CSV & Markdown
    staff_ws = wb["STAFF_TIMETABLE"]
    with open("Timetable_CSVs/STAFF_TIMETABLE.csv", "w", newline="", encoding="utf-8") as f_csv:
        writer = csv.writer(f_csv)
        for r in range(1, staff_ws.max_row + 1):
            row_vals = [staff_ws.cell(row=r, column=c).value or "" for c in range(1, staff_ws.max_column + 1)]
            writer.writerow(row_vals)

    # Room utilization as CSV & Markdown
    room_ws = wb["ROOM_UTILIZATION"]
    with open("Timetable_CSVs/ROOM_UTILIZATION.csv", "w", newline="", encoding="utf-8") as f_csv:
        writer = csv.writer(f_csv)
        for r in range(1, room_ws.max_row + 1):
            row_vals = [room_ws.cell(row=r, column=c).value or "" for c in range(1, room_ws.max_column + 1)]
            writer.writerow(row_vals)

    print("Created Timetable_Text_Views/, Timetable_CSVs/, and ALL_TIMETABLES.md!")

if __name__ == "__main__":
    generate_text_and_csv()
