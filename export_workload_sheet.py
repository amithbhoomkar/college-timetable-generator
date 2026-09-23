import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict
import csv
import os

def generate_workload_sheets():
    wb_src = openpyxl.load_workbook('Complete_Timetable.xlsx', data_only=True)
    ws_staff = wb_src['STAFF_TIMETABLE']
    rows = list(ws_staff.iter_rows(values_only=True))[1:]

    DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

    staff_dict = defaultdict(lambda: {
        'designation': '',
        'tot_hrs': 0,
        'theory_hrs': 0,
        'lab_sessions': 0,
        'lab_hrs': 0,
        'oe_hrs': 0,
        'pe_hrs': 0,
        'days': defaultdict(int),
        'subjects': set(),
        'sections': set(),
        'subject_details': defaultdict(lambda: {'type': '', 'hrs': 0, 'sections': set()})
    })

    for r in rows:
        teacher, desig, day, time_slot, subject, section, room, ev_type = r
        d = staff_dict[teacher]
        d['designation'] = desig
        hrs = 3 if ev_type == 'LAB' else 1
        d['tot_hrs'] += hrs
        if ev_type == 'THEORY':
            d['theory_hrs'] += 1
        elif ev_type == 'LAB':
            d['lab_sessions'] += 1
            d['lab_hrs'] += 3
        elif ev_type == 'OE':
            d['oe_hrs'] += 1
        elif ev_type == 'PE':
            d['pe_hrs'] += 1

        d['days'][day] += hrs
        d['subjects'].add(subject)
        d['sections'].add(section)
        d['subject_details'][subject]['type'] = ev_type
        d['subject_details'][subject]['hrs'] += hrs
        d['subject_details'][subject]['sections'].add(section)

    def sort_key(name):
        if 'Professor ' in name and 'Associate' not in name and 'Assistant' not in name:
            return (1, int(name.split()[-1]))
        elif 'Associate' in name:
            return (2, int(name.split()[-1]))
        else:
            return (3, int(name.split()[-1]))

    sorted_teachers = sorted(staff_dict.keys(), key=sort_key)

    # 1. CREATE WORKLOAD EXCEL WORKBOOK
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TEACHER_WORKLOAD"
    ws.views.sheetView[0].showGridLines = True

    # Styling definitions
    font_title = Font(name="Calibri", size=16, bold=True, color="1F4E79")
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="595959")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    font_small = Font(name="Calibri", size=10, italic=True)

    fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_day_header = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_subtotal = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_zebra = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    thin_border_side = Side(style="thin", color="D9D9D9")
    border_data = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_header = Border(
        left=Side(style="thin", color="FFFFFF"),
        right=Side(style="thin", color="FFFFFF"),
        top=Side(style="medium", color="1F4E79"),
        bottom=Side(style="medium", color="1F4E79")
    )
    double_bottom = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=Side(style="thin", color="000000"),
        bottom=Side(style="double", color="000000")
    )

    # Title Banner
    ws.merge_cells("A1:Q1")
    ws["A1"] = "COLLEGE OF ENGINEERING - FACULTY WORKLOAD DISTRIBUTION SHEET"
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    ws.merge_cells("A2:Q2")
    ws["A2"] = "Weekly Teaching Load Summary (Monday - Saturday) | ISE & CSBS Departments"
    ws["A2"].font = font_subtitle
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    headers = [
        "Sl No", "Faculty Name", "Designation", "Total Hours", 
        "Theory (Hrs)", "Lab (Sessions)", "Lab (Hrs)", "OE (Hrs)", "PE (Hrs)",
        "Mon", "Tue", "Wed", "Thu", "Fri", "Sat",
        "Assigned Subjects", "Assigned Sections"
    ]

    ws.row_dimensions[4].height = 28
    for col_num, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_num)
        cell.value = h
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_header
        if h in ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]:
            cell.fill = fill_day_header
        else:
            cell.fill = fill_header

    curr_row = 5
    tot_theory = 0
    tot_lab_sessions = 0
    tot_lab_hrs = 0
    tot_oe = 0
    tot_pe = 0
    grand_tot_hrs = 0
    day_totals = defaultdict(int)

    for idx, t in enumerate(sorted_teachers, 1):
        d = staff_dict[t]
        subj_str = ", ".join(sorted(d['subjects']))
        sec_str = ", ".join(sorted(d['sections']))

        row_vals = [
            idx,
            t,
            d['designation'],
            d['tot_hrs'],
            d['theory_hrs'],
            d['lab_sessions'],
            d['lab_hrs'],
            d['oe_hrs'],
            d['pe_hrs'],
            d['days']['Monday'],
            d['days']['Tuesday'],
            d['days']['Wednesday'],
            d['days']['Thursday'],
            d['days']['Friday'],
            d['days']['Saturday'],
            subj_str,
            sec_str
        ]

        ws.row_dimensions[curr_row].height = 24
        is_zebra = (idx % 2 == 0)

        for col_num, val in enumerate(row_vals, 1):
            cell = ws.cell(row=curr_row, column=col_num)
            cell.value = val
            cell.font = font_regular
            cell.border = border_data

            if is_zebra:
                cell.fill = fill_zebra

            if col_num in [1, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_num in [2, 3]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)

            if col_num == 4:
                cell.font = font_bold

        grand_tot_hrs += d['tot_hrs']
        tot_theory += d['theory_hrs']
        tot_lab_sessions += d['lab_sessions']
        tot_lab_hrs += d['lab_hrs']
        tot_oe += d['oe_hrs']
        tot_pe += d['pe_hrs']
        for day in DAYS:
            day_totals[day] += d['days'][day]

        curr_row += 1

    # Total Row
    ws.row_dimensions[curr_row].height = 26
    total_vals = [
        "", "TOTAL / SUMMARY", f"{len(sorted_teachers)} Faculty", grand_tot_hrs,
        tot_theory, tot_lab_sessions, tot_lab_hrs, tot_oe, tot_pe,
        day_totals['Monday'], day_totals['Tuesday'], day_totals['Wednesday'],
        day_totals['Thursday'], day_totals['Friday'], day_totals['Saturday'],
        "-", "-"
    ]

    for col_num, val in enumerate(total_vals, 1):
        cell = ws.cell(row=curr_row, column=col_num)
        cell.value = val
        cell.font = font_bold
        cell.fill = fill_subtotal
        cell.border = double_bottom
        if col_num in [1, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

    # Column widths
    col_widths = {
        1: 8,   # Sl No
        2: 26,  # Faculty Name
        3: 22,  # Designation
        4: 13,  # Total Hours
        5: 14,  # Theory
        6: 14,  # Lab Sessions
        7: 12,  # Lab Hrs
        8: 10,  # OE Hrs
        9: 10,  # PE Hrs
        10: 8,  # Mon
        11: 8,  # Tue
        12: 8,  # Wed
        13: 8,  # Thu
        14: 8,  # Fri
        15: 8,  # Sat
        16: 45, # Subjects
        17: 35  # Sections
    }
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Save standalone Excel workload sheet
    excel_path = "Teacher_Workload_Sheet.xlsx"
    wb.save(excel_path)
    wb.save("Timetable_Exports/Teacher_Workload_Sheet.xlsx")
    print(f"Saved {excel_path} and Timetable_Exports/Teacher_Workload_Sheet.xlsx")

    # Also add sheet to Complete_Timetable.xlsx
    wb_master = openpyxl.load_workbook("Complete_Timetable.xlsx")
    if "TEACHER_WORKLOAD" in wb_master.sheetnames:
        del wb_master["TEACHER_WORKLOAD"]
    # Copy sheet to wb_master
    ws_new = wb_master.create_sheet(title="TEACHER_WORKLOAD", index=0)
    # copy rows and styling
    for r in range(1, ws.max_row + 1):
        ws_new.row_dimensions[r].height = ws.row_dimensions[r].height
        for c in range(1, ws.max_column + 1):
            src_cell = ws.cell(row=r, column=c)
            dest_cell = ws_new.cell(row=r, column=c)
            dest_cell.value = src_cell.value
            if src_cell.has_style:
                dest_cell.font = src_cell.font.copy()
                dest_cell.fill = src_cell.fill.copy()
                dest_cell.border = src_cell.border.copy()
                dest_cell.alignment = src_cell.alignment.copy()

    for col_idx, width in col_widths.items():
        ws_new.column_dimensions[get_column_letter(col_idx)].width = width
    ws_new.views.sheetView[0].showGridLines = True
    ws_new.merge_cells("A1:Q1")
    ws_new.merge_cells("A2:Q2")
    wb_master.save("Complete_Timetable.xlsx")
    print("Updated Complete_Timetable.xlsx with TEACHER_WORKLOAD tab at position 1.")

    # 2. CREATE CSV EXPORT
    csv_path = "Timetable_CSVs/TEACHER_WORKLOAD.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for idx, t in enumerate(sorted_teachers, 1):
            d = staff_dict[t]
            writer.writerow([
                idx, t, d['designation'], d['tot_hrs'],
                d['theory_hrs'], d['lab_sessions'], d['lab_hrs'], d['oe_hrs'], d['pe_hrs'],
                d['days']['Monday'], d['days']['Tuesday'], d['days']['Wednesday'],
                d['days']['Thursday'], d['days']['Friday'], d['days']['Saturday'],
                "; ".join(sorted(d['subjects'])),
                "; ".join(sorted(d['sections']))
            ])
        writer.writerow(total_vals)
    print(f"Saved {csv_path}")

    # 3. CREATE MARKDOWN TEXT VIEW
    md_path = "Timetable_Text_Views/TEACHER_WORKLOAD.md"
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("# FACULTY WORKLOAD DISTRIBUTION SHEET\n\n")
        f.write("**Weekly Teaching Hours Summary (Monday to Saturday)**\n\n")
        f.write("| Sl | Faculty Name | Designation | Total Hrs | Theory | Lab (Hrs) | OE | PE | Mon | Tue | Wed | Thu | Fri | Sat | Subjects Handled | Sections |\n")
        f.write("|:---:|:---|:---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---|:---|\n")

        for idx, t in enumerate(sorted_teachers, 1):
            d = staff_dict[t]
            lab_col = f"{d['lab_sessions']} ({d['lab_hrs']}h)" if d['lab_sessions'] > 0 else "-"
            f.write(f"| {idx} | **{t}** | {d['designation']} | **{d['tot_hrs']}** | {d['theory_hrs']} | {lab_col} | {d['oe_hrs']} | {d['pe_hrs']} | {d['days']['Monday']} | {d['days']['Tuesday']} | {d['days']['Wednesday']} | {d['days']['Thursday']} | {d['days']['Friday']} | {d['days']['Saturday']} | {', '.join(sorted(d['subjects']))} | {', '.join(sorted(d['sections']))} |\n")

        f.write(f"| | **TOTAL** | **20 Faculty** | **{grand_tot_hrs}** | **{tot_theory}** | **{tot_lab_sessions} ({tot_lab_hrs}h)** | **{tot_oe}** | **{tot_pe}** | **{day_totals['Monday']}** | **{day_totals['Tuesday']}** | **{day_totals['Wednesday']}** | **{day_totals['Thursday']}** | **{day_totals['Friday']}** | **{day_totals['Saturday']}** | - | - |\n\n")

        f.write("## Detailed Faculty-Wise Subject Breakdown\n\n")
        for t in sorted_teachers:
            d = staff_dict[t]
            f.write(f"### {t} ({d['designation']})\n")
            f.write(f"- **Total Weekly Workload**: {d['tot_hrs']} Hours (Theory: {d['theory_hrs']}h, Lab: {d['lab_hrs']}h, OE: {d['oe_hrs']}h, PE: {d['pe_hrs']}h)\n")
            day_str = ", ".join([f"{day[:3]}: {d['days'][day]}h" for day in DAYS])
            f.write(f"- **Day-wise**: {day_str}\n")
            f.write("- **Courses Taught**:\n")
            for subj, sinfo in sorted(d['subject_details'].items()):
                f.write(f"  - `{subj}` [{sinfo['type']}]: {sinfo['hrs']} hrs/week (Sections: {', '.join(sorted(sinfo['sections']))})\n")
            f.write("\n")
    print(f"Saved {md_path}")

if __name__ == "__main__":
    generate_workload_sheets()
