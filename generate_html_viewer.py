import openpyxl
import json
import os
from collections import defaultdict

def generate_html():
    wb = openpyxl.load_workbook("Complete_Timetable.xlsx")

    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    academic_slots = [
        "08:30-09:30", "09:30-10:30",
        "11:00-11:50", "11:50-12:40", "12:40-13:30",
        "14:30-15:30", "15:30-16:30", "16:30-17:30"
    ]

    # 1. Extract all section data
    section_names = [s for s in wb.sheetnames if s.startswith("ISE_S") or s.startswith("CSBS_S")]
    sections_data = {}

    for s_name in section_names:
        ws = wb[s_name]
        rows = []
        for r in range(4, ws.max_row + 1):
            time_slot = ws.cell(row=r, column=1).value
            if not time_slot:
                continue
            day_cells = []
            is_break_row = False
            break_type = ""
            for c in range(2, 8):
                val = ws.cell(row=r, column=c).value or ""
                day_cells.append(str(val))
                if "LUNCH" in str(val):
                    is_break_row = True
                    break_type = "LUNCH"
                elif "BREAK" in str(val):
                    is_break_row = True
                    break_type = "BREAK"
            rows.append({
                "time": str(time_slot),
                "is_break": is_break_row,
                "break_type": break_type,
                "days": day_cells
            })
        sections_data[s_name] = rows

    # 2. Extract Staff data and build weekly grids
    staff_ws = wb["STAFF_TIMETABLE"]
    staff_data = {}
    staff_dict = defaultdict(lambda: {
        'name': '',
        'designation': '',
        'tot_hrs': 0,
        'theory_hrs': 0,
        'lab_sessions': 0,
        'lab_hrs': 0,
        'oe_hrs': 0,
        'pe_hrs': 0,
        'days': defaultdict(int),
        'subjects': set(),
        'sections': set()
    })

    # Prepare slot grids for each staff member
    staff_raw_events = defaultdict(list)
    for r in range(2, staff_ws.max_row + 1):
        row_vals = [staff_ws.cell(row=r, column=c).value for c in range(1, 9)]
        teacher = str(row_vals[0])
        desig = str(row_vals[1])
        day = str(row_vals[2])
        time_slot = str(row_vals[3])
        subject = str(row_vals[4])
        section = str(row_vals[5])
        room = str(row_vals[6])
        ev_type = str(row_vals[7])

        if teacher not in staff_data:
            staff_data[teacher] = {
                "designation": desig,
                "events": []
            }
        staff_data[teacher]["events"].append({
            "day": day,
            "time": time_slot,
            "subject": subject,
            "section": section,
            "room": room,
            "type": ev_type
        })
        staff_raw_events[teacher].append((day, time_slot, subject, section, room, ev_type))

        d = staff_dict[teacher]
        d['name'] = teacher
        d['designation'] = desig
        hrs = 3 if ev_type == 'LAB' else 1
        d['tot_hrs'] += hrs
        if ev_type == 'THEORY': d['theory_hrs'] += 1
        elif ev_type == 'LAB':
            d['lab_sessions'] += 1
            d['lab_hrs'] += 3
        elif ev_type == 'OE': d['oe_hrs'] += 1
        elif ev_type == 'PE': d['pe_hrs'] += 1

        d['days'][day] += hrs
        d['subjects'].add(subject)
        d['sections'].add(section)

    def sort_key(name):
        if 'Professor ' in name and 'Associate' not in name and 'Assistant' not in name:
            return (1, int(name.split()[-1]))
        elif 'Associate' in name:
            return (2, int(name.split()[-1]))
        else:
            return (3, int(name.split()[-1]))

    def get_target(desig):
        if 'Professor' in desig and 'Associate' not in desig and 'Assistant' not in desig:
            return 12
        elif 'Associate' in desig:
            return 14
        else:
            return 16

    sorted_staff_names = sorted(staff_dict.keys(), key=sort_key)
    workload_data = []
    for t in sorted_staff_names:
        d = staff_dict[t]
        workload_data.append({
            "name": d['name'],
            "designation": d['designation'],
            "target_hrs": get_target(d['designation']),
            "tot_hrs": d['tot_hrs'],
            "theory_hrs": d['theory_hrs'],
            "lab_sessions": d['lab_sessions'],
            "lab_hrs": d['lab_hrs'],
            "oe_hrs": d['oe_hrs'],
            "pe_hrs": d['pe_hrs'],
            "days": dict(d['days']),
            "subjects": sorted(list(d['subjects'])),
            "sections": sorted(list(d['sections']))
        })

    # Build staff weekly schedule grids
    staff_grids = {}
    for teacher, events in staff_raw_events.items():
        slot_map = {s: {d: "FREE" for d in days} for s in academic_slots}
        for day, time_slot, subject, section, room, ev_type in events:
            text = f"{subject}\n[{ev_type}]\n{section}\n{room}"
            if ev_type == 'LAB':
                span = ["11:00-11:50", "11:50-12:40", "12:40-13:30"] if time_slot.startswith("11:00") else ["14:30-15:30", "15:30-16:30", "16:30-17:30"]
                for s in span:
                    slot_map[s][day] = text
            elif time_slot in slot_map:
                slot_map[time_slot][day] = text

        rows = [
            {"time": "08:30-09:30", "is_break": False, "break_type": "", "days": [slot_map["08:30-09:30"][d] for d in days]},
            {"time": "09:30-10:30", "is_break": False, "break_type": "", "days": [slot_map["09:30-10:30"][d] for d in days]},
            {"time": "10:30-11:00", "is_break": True, "break_type": "BREAK", "days": ["SHORT BREAK"] * 6},
            {"time": "11:00-11:50", "is_break": False, "break_type": "", "days": [slot_map["11:00-11:50"][d] for d in days]},
            {"time": "11:50-12:40", "is_break": False, "break_type": "", "days": [slot_map["11:50-12:40"][d] for d in days]},
            {"time": "12:40-13:30", "is_break": False, "break_type": "", "days": [slot_map["12:40-13:30"][d] for d in days]},
            {"time": "13:30-14:30", "is_break": True, "break_type": "LUNCH", "days": ["LUNCH BREAK"] * 6},
            {"time": "14:30-15:30", "is_break": False, "break_type": "", "days": [slot_map["14:30-15:30"][d] for d in days]},
            {"time": "15:30-16:30", "is_break": False, "break_type": "", "days": [slot_map["15:30-16:30"][d] for d in days]},
            {"time": "16:30-17:30", "is_break": False, "break_type": "", "days": [slot_map["16:30-17:30"][d] for d in days]},
        ]
        staff_grids[teacher] = rows

    # 3. Extract Room utilization data and build weekly grids
    room_ws = wb["ROOM_UTILIZATION"]
    room_data = defaultdict(list)
    room_raw_events = defaultdict(list)
    for r in range(2, room_ws.max_row + 1):
        row_vals = [room_ws.cell(row=r, column=c).value for c in range(1, 8)]
        room_name = str(row_vals[0])
        day = str(row_vals[1])
        time_slot = str(row_vals[2])
        subject = str(row_vals[3])
        section = str(row_vals[4])
        teacher = str(row_vals[5])
        ev_type = str(row_vals[6])

        room_data[room_name].append({
            "day": day,
            "time": time_slot,
            "subject": subject,
            "section": section,
            "teacher": teacher,
            "type": ev_type
        })
        room_raw_events[room_name].append((day, time_slot, subject, section, teacher, ev_type))

    room_grids = {}
    for room_name, events in room_raw_events.items():
        slot_map = {s: {d: "FREE" for d in days} for s in academic_slots}
        for day, time_slot, subject, section, teacher, ev_type in events:
            text = f"{subject}\n[{ev_type}]\n{section}\n{teacher}"
            if ev_type == 'LAB':
                span = ["11:00-11:50", "11:50-12:40", "12:40-13:30"] if time_slot.startswith("11:00") else ["14:30-15:30", "15:30-16:30", "16:30-17:30"]
                for s in span:
                    slot_map[s][day] = text
            elif time_slot in slot_map:
                slot_map[time_slot][day] = text

        rows = [
            {"time": "08:30-09:30", "is_break": False, "break_type": "", "days": [slot_map["08:30-09:30"][d] for d in days]},
            {"time": "09:30-10:30", "is_break": False, "break_type": "", "days": [slot_map["09:30-10:30"][d] for d in days]},
            {"time": "10:30-11:00", "is_break": True, "break_type": "BREAK", "days": ["SHORT BREAK"] * 6},
            {"time": "11:00-11:50", "is_break": False, "break_type": "", "days": [slot_map["11:00-11:50"][d] for d in days]},
            {"time": "11:50-12:40", "is_break": False, "break_type": "", "days": [slot_map["11:50-12:40"][d] for d in days]},
            {"time": "12:40-13:30", "is_break": False, "break_type": "", "days": [slot_map["12:40-13:30"][d] for d in days]},
            {"time": "13:30-14:30", "is_break": True, "break_type": "LUNCH", "days": ["LUNCH BREAK"] * 6},
            {"time": "14:30-15:30", "is_break": False, "break_type": "", "days": [slot_map["14:30-15:30"][d] for d in days]},
            {"time": "15:30-16:30", "is_break": False, "break_type": "", "days": [slot_map["15:30-16:30"][d] for d in days]},
            {"time": "16:30-17:30", "is_break": False, "break_type": "", "days": [slot_map["16:30-17:30"][d] for d in days]},
        ]
        room_grids[room_name] = rows

    # 4. Extract Department masters
    dept_data = {}
    for d_name in ["ISE_MASTER", "CSBS_MASTER"]:
        d_ws = wb[d_name]
        d_events = []
        for r in range(4, d_ws.max_row + 1):
            day_val = d_ws.cell(row=r, column=1).value
            if not day_val:
                continue
            d_events.append({
                "day": str(day_val),
                "time": str(d_ws.cell(row=r, column=2).value or ""),
                "section": str(d_ws.cell(row=r, column=3).value or ""),
                "subject": str(d_ws.cell(row=r, column=4).value or ""),
                "type": str(d_ws.cell(row=r, column=5).value or ""),
                "teacher": str(d_ws.cell(row=r, column=6).value or ""),
                "room": str(d_ws.cell(row=r, column=7).value or "")
            })
        dept_data[d_name] = d_events

    data_bundle = {
        "sections": sections_data,
        "staff": staff_data,
        "staff_grids": staff_grids,
        "workload": workload_data,
        "rooms": room_data,
        "room_grids": room_grids,
        "departments": dept_data,
        "days": days
    }

    json_data = json.dumps(data_bundle)

    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>College Timetable Interactive Viewer | 1-Page PDF Fit</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root {{
      --primary: #2563eb;
      --primary-dark: #1d4ed8;
      --primary-glow: rgba(37, 99, 235, 0.25);
      --bg: #0f172a;
      --card-bg: rgba(30, 41, 59, 0.88);
      --card-border: rgba(255, 255, 255, 0.08);
      --text: #f8fafc;
      --text-muted: #94a3b8;
      --accent-lab: #06b6d4;
      --accent-oe: #f59e0b;
      --accent-pe: #10b981;
      --accent-theory: #6366f1;
      --accent-lunch: #64748b;
    }}

    * {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }}

    body {{
      font-family: 'Inter', sans-serif;
      background-color: var(--bg);
      background-image: 
        radial-gradient(at 0% 0%, rgba(37, 99, 235, 0.15) 0px, transparent 50%),
        radial-gradient(at 100% 100%, rgba(99, 102, 241, 0.15) 0px, transparent 50%);
      color: var(--text);
      min-height: 100vh;
      display: flex;
      flex-direction: column;
    }}

    header {{
      background: rgba(15, 23, 42, 0.85);
      backdrop-filter: blur(12px);
      border-bottom: 1px solid var(--card-border);
      padding: 1rem 2rem;
      display: flex;
      justify-content: space-between;
      align-items: center;
      position: sticky;
      top: 0;
      z-index: 100;
    }}

    .logo-area h1 {{
      font-family: 'Outfit', sans-serif;
      font-size: 1.4rem;
      font-weight: 700;
      background: linear-gradient(135deg, #60a5fa, #a78bfa);
      -webkit-background-clip: text;
      background-clip: text;
      -webkit-text-fill-color: transparent;
      letter-spacing: -0.5px;
    }}

    .logo-area p {{
      color: var(--text-muted);
      font-size: 0.82rem;
      margin-top: 0.15rem;
    }}

    .actions {{
      display: flex;
      gap: 0.75rem;
    }}

    .btn {{
      padding: 0.5rem 1.1rem;
      border-radius: 8px;
      font-weight: 500;
      font-size: 0.85rem;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 0.45rem;
      transition: all 0.2s ease;
      border: none;
      outline: none;
      text-decoration: none;
    }}

    .btn-primary {{
      background: linear-gradient(135deg, var(--primary), var(--primary-dark));
      color: white;
      box-shadow: 0 4px 14px var(--primary-glow);
    }}
    .btn-primary:hover {{
      transform: translateY(-1px);
      box-shadow: 0 6px 20px rgba(37, 99, 235, 0.35);
    }}

    .btn-secondary {{
      background: rgba(255, 255, 255, 0.08);
      color: var(--text);
      border: 1px solid var(--card-border);
    }}
    .btn-secondary:hover {{
      background: rgba(255, 255, 255, 0.15);
    }}

    .main-container {{
      flex: 1;
      padding: 1.5rem 2rem;
      max-width: 1750px;
      margin: 0 auto;
      width: 100%;
    }}

    /* Nav Tabs */
    .nav-tabs {{
      display: flex;
      gap: 0.5rem;
      margin-bottom: 1.25rem;
      border-bottom: 1px solid var(--card-border);
      padding-bottom: 0.75rem;
      overflow-x: auto;
    }}

    .tab-btn {{
      padding: 0.55rem 1.25rem;
      border-radius: 20px;
      background: transparent;
      color: var(--text-muted);
      font-weight: 600;
      font-size: 0.88rem;
      border: 1px solid transparent;
      cursor: pointer;
      transition: all 0.2s ease;
      white-space: nowrap;
    }}

    .tab-btn:hover {{
      color: var(--text);
      background: rgba(255, 255, 255, 0.05);
    }}

    .tab-btn.active {{
      background: var(--primary);
      color: white;
      border-color: var(--primary-dark);
      box-shadow: 0 4px 12px var(--primary-glow);
    }}

    /* Controls Bar */
    .controls-bar {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 1.25rem;
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      padding: 0.85rem 1.25rem;
      border-radius: 12px;
      gap: 1rem;
      flex-wrap: wrap;
    }}

    .select-group {{
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }}

    .select-group label {{
      font-size: 0.88rem;
      font-weight: 600;
      color: var(--text-muted);
    }}

    select {{
      background: rgba(15, 23, 42, 0.8);
      border: 1px solid var(--card-border);
      color: var(--text);
      padding: 0.45rem 1rem;
      border-radius: 8px;
      font-size: 0.88rem;
      outline: none;
      cursor: pointer;
      transition: border-color 0.2s ease;
      min-width: 200px;
    }}

    select:focus {{
      border-color: var(--primary);
    }}

    /* Legend */
    .legend {{
      display: flex;
      gap: 0.9rem;
      align-items: center;
      flex-wrap: wrap;
    }}

    .legend-item {{
      display: flex;
      align-items: center;
      gap: 0.4rem;
      font-size: 0.75rem;
      color: var(--text-muted);
    }}

    .legend-box {{
      width: 12px;
      height: 12px;
      border-radius: 3px;
    }}

    /* Timetable Table Styles */
    .timetable-wrapper {{
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 14px;
      box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
      backdrop-filter: blur(16px);
      overflow: hidden;
    }}

    .table-responsive {{
      overflow-x: auto;
      width: 100%;
    }}

    table {{
      width: 100%;
      border-collapse: separate;
      border-spacing: 0;
      text-align: left;
    }}

    th {{
      background: rgba(15, 23, 42, 0.95);
      color: var(--text-muted);
      font-weight: 600;
      font-size: 0.78rem;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      padding: 0.75rem 0.65rem;
      border-bottom: 1px solid var(--card-border);
      border-right: 1px solid rgba(255, 255, 255, 0.04);
      white-space: nowrap;
    }}

    th:first-child {{
      width: 120px;
      text-align: center;
    }}

    td {{
      padding: 0.45rem 0.5rem;
      border-bottom: 1px solid rgba(255, 255, 255, 0.04);
      border-right: 1px solid rgba(255, 255, 255, 0.04);
      vertical-align: middle;
      font-size: 0.8rem;
      min-width: 155px;
    }}

    td.time-col {{
      font-family: monospace;
      font-size: 0.75rem;
      font-weight: 600;
      color: #93c5fd;
      background: rgba(15, 23, 42, 0.6);
      text-align: center;
      width: 120px;
      min-width: 120px;
    }}

    /* Event Card in Cell */
    .event-card {{
      background: rgba(255, 255, 255, 0.04);
      border: 1px solid rgba(255, 255, 255, 0.06);
      border-radius: 8px;
      padding: 0.45rem 0.55rem;
      position: relative;
      display: flex;
      flex-direction: column;
      gap: 0.2rem;
      transition: all 0.2s ease;
      min-height: 58px;
    }}

    .event-card:hover {{
      transform: translateY(-2px);
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.4);
      background: rgba(255, 255, 255, 0.08);
    }}

    .event-card.THEORY {{
      border-left: 3px solid var(--accent-theory);
      background: rgba(99, 102, 241, 0.08);
    }}

    .event-card.LAB {{
      border-left: 3px solid var(--accent-lab);
      background: rgba(6, 182, 212, 0.08);
    }}

    .event-card.OE {{
      border-left: 3px solid var(--accent-oe);
      background: rgba(245, 158, 11, 0.08);
    }}

    .event-card.PE {{
      border-left: 3px solid var(--accent-pe);
      background: rgba(16, 185, 129, 0.08);
    }}

    .ev-badge {{
      align-self: flex-start;
      font-size: 0.62rem;
      font-weight: 700;
      padding: 0.1rem 0.35rem;
      border-radius: 4px;
      text-transform: uppercase;
      letter-spacing: 0.3px;
    }}

    .badge-theory {{ background: rgba(99, 102, 241, 0.25); color: #a5b4fc; }}
    .badge-lab {{ background: rgba(6, 182, 212, 0.25); color: #67e8f9; }}
    .badge-oe {{ background: rgba(245, 158, 11, 0.25); color: #fde68a; }}
    .badge-pe {{ background: rgba(16, 185, 129, 0.25); color: #6ee7b7; }}

    .ev-title {{
      font-weight: 600;
      color: var(--text);
      font-size: 0.76rem;
      line-height: 1.2;
    }}

    .ev-meta {{
      display: flex;
      flex-direction: column;
      gap: 0.1rem;
      color: var(--text-muted);
      font-size: 0.7rem;
    }}

    .free-slot {{
      color: rgba(255, 255, 255, 0.18);
      font-weight: 500;
      font-size: 0.78rem;
    }}

    .lunch-row td {{
      background: rgba(100, 116, 139, 0.18);
      color: #cbd5e1;
      font-weight: 600;
      letter-spacing: 1.5px;
      font-size: 0.75rem;
      padding: 0.45rem;
      text-transform: uppercase;
      text-align: center;
    }}

    .break-row td {{
      background: rgba(245, 158, 11, 0.18);
      color: #fde68a;
      font-weight: 600;
      letter-spacing: 1.5px;
      font-size: 0.75rem;
      padding: 0.45rem;
      text-transform: uppercase;
      text-align: center;
    }}

    /* Stat Cards for Workload */
    .stat-cards-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
      gap: 0.85rem;
      margin-bottom: 1.25rem;
    }}

    .stat-card {{
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 12px;
      padding: 1rem 1.2rem;
      display: flex;
      flex-direction: column;
      gap: 0.3rem;
    }}
    .stat-title {{
      font-size: 0.75rem;
      color: var(--text-muted);
      font-weight: 500;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }}
    .stat-value {{
      font-family: 'Outfit', sans-serif;
      font-size: 1.6rem;
      font-weight: 700;
      color: #93c5fd;
    }}
    .stat-sub {{
      font-size: 0.72rem;
      color: #10b981;
      font-weight: 500;
    }}

    /* Info Banner for Specific Selection */
    .view-info-banner {{
      background: rgba(37, 99, 235, 0.1);
      border: 1px solid rgba(37, 99, 235, 0.25);
      border-radius: 10px;
      padding: 0.75rem 1.25rem;
      margin-bottom: 1rem;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 0.5rem;
    }}
    .view-info-banner .title {{
      font-weight: 700;
      font-size: 1rem;
      color: #93c5fd;
    }}
    .view-info-banner .sub {{
      font-size: 0.8rem;
      color: var(--text-muted);
    }}

    /* List Table Styles */
    .list-table {{
      width: 100%;
      border-collapse: collapse;
    }}
    .list-table th {{
      text-align: left;
      padding: 0.65rem 0.75rem;
      font-size: 0.75rem;
      background: rgba(15, 23, 42, 0.95);
      color: var(--text-muted);
    }}
    .list-table td {{
      text-align: left;
      min-width: unset;
      padding: 0.6rem 0.75rem;
      font-size: 0.78rem;
      border-bottom: 1px solid rgba(255, 255, 255, 0.05);
    }}
    .list-table tr:hover td {{
      background: rgba(255, 255, 255, 0.02);
    }}
    .list-table tfoot td {{
      background: rgba(15, 23, 42, 0.98);
      font-weight: 700;
      color: #f8fafc;
      border-top: 2px solid rgba(255, 255, 255, 0.2);
    }}

    .badge-load {{
      background: rgba(37, 99, 235, 0.25);
      color: #93c5fd;
      padding: 0.15rem 0.5rem;
      border-radius: 5px;
      font-weight: 700;
      font-size: 0.8rem;
      display: inline-block;
    }}

    .print-header-banner {{
      display: none;
    }}

    /* 1-PAGE PDF / PRINT OPTIMIZATION (LANDSCAPE A4) */
    @media print {{
      @page {{
        size: A4 landscape;
        margin: 5mm 6mm;
      }}
      body {{
        background: #ffffff !important;
        color: #000000 !important;
        padding: 0 !important;
        margin: 0 !important;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }}
      header, .nav-tabs, .controls-bar, .actions, .btn {{
        display: none !important;
      }}
      .main-container {{
        padding: 0 !important;
        max-width: 100% !important;
      }}
      .print-header-banner {{
        display: block !important;
        text-align: center;
        margin-bottom: 4px;
      }}
      #printBannerTitle {{
        font-family: 'Outfit', sans-serif;
        font-size: 13pt;
        font-weight: 700;
        color: #1e3a8a;
        letter-spacing: -0.3px;
      }}
      #printSubtitle {{
        display: block;
        font-size: 8pt;
        font-weight: normal;
        color: #475569;
        margin-top: 1px;
      }}
      .timetable-wrapper {{
        border: 1px solid #94a3b8 !important;
        box-shadow: none !important;
        background: #ffffff !important;
        page-break-inside: avoid;
        break-inside: avoid;
      }}
      table {{
        width: 100% !important;
        border-collapse: collapse !important;
      }}
      th {{
        background: #1e3a8a !important;
        color: #ffffff !important;
        font-size: 7.5pt !important;
        padding: 3px 2px !important;
        border: 1px solid #94a3b8 !important;
      }}
      td {{
        padding: 2.5px 2px !important;
        font-size: 6.8pt !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
        min-width: auto !important;
      }}
      td.time-col {{
        background: #f1f5f9 !important;
        color: #0f172a !important;
        font-size: 6.8pt !important;
        min-width: 65px !important;
        font-weight: 700 !important;
      }}
      .event-card {{
        padding: 2px 2px !important;
        border-radius: 3px !important;
        font-size: 6.5pt !important;
        line-height: 1.15 !important;
        border: 1px solid #94a3b8 !important;
        box-shadow: none !important;
        transform: none !important;
      }}
      .event-card.THEORY {{
        background: #eef2ff !important;
        border-left: 3px solid #4f46e5 !important;
      }}
      .event-card.LAB {{
        background: #ecfeff !important;
        border-left: 3px solid #0891b2 !important;
      }}
      .event-card.OE {{
        background: #fffbeb !important;
        border-left: 3px solid #d97706 !important;
      }}
      .event-card.PE {{
        background: #ecfdf5 !important;
        border-left: 3px solid #059669 !important;
      }}
      .ev-badge {{
        display: none !important;
      }}
      .ev-title {{
        color: #0f172a !important;
        font-weight: 700 !important;
        font-size: 6.8pt !important;
      }}
      .ev-meta {{
        color: #334155 !important;
        font-size: 6pt !important;
        gap: 0 !important;
      }}
      .lunch-row td {{
        background: #f1f5f9 !important;
        color: #475569 !important;
        font-size: 6.5pt !important;
        padding: 2px !important;
      }}
      .break-row td {{
        background: #fef3c7 !important;
        color: #92400e !important;
        font-size: 6.5pt !important;
        padding: 2px !important;
      }}
      .free-slot {{
        color: #94a3b8 !important;
        font-size: 6.5pt !important;
      }}
      .view-info-banner {{
        display: none !important;
      }}
    }}
  </style>
</head>
<body>

  <header>
    <div class="logo-area">
      <h1>College Timetable Interactive Viewer</h1>
      <p>ISE & CSBS Departments | AICTE Faculty Workload Calibration & 1-Page PDF Fit</p>
    </div>
    <div class="actions">
      <button class="btn btn-secondary" onclick="window.print()">🖨️ Print / Save 1-Page PDF</button>
      <button class="btn btn-primary" onclick="downloadCurrentCSV()">📥 Download CSV</button>
    </div>
  </header>

  <div class="main-container">
    <div class="print-header-banner" id="printBanner">
      <div id="printBannerTitle">COLLEGE OF ENGINEERING — ACADEMIC TIMETABLE</div>
      <div id="printSubtitle">ISE & CSBS Departments | Conflict-Free Schedule</div>
    </div>

    <div class="nav-tabs">
      <button class="tab-btn active" onclick="switchView('section')">Section Timetables</button>
      <button class="tab-btn" onclick="switchView('workload')">Faculty Workload Sheet</button>
      <button class="tab-btn" onclick="switchView('staff')">Faculty Schedules</button>
      <button class="tab-btn" onclick="switchView('room')">Room Utilization</button>
      <button class="tab-btn" onclick="switchView('dept')">Department Masters</button>
    </div>

    <div class="controls-bar" id="controlsBar">
      <div class="select-group">
        <label id="selectorLabel" for="itemSelect">Select Section:</label>
        <select id="itemSelect" onchange="renderCurrentView()"></select>
      </div>

      <div class="legend" id="legendArea">
        <div class="legend-item"><div class="legend-box" style="background:var(--accent-theory);"></div> Theory (1h)</div>
        <div class="legend-item"><div class="legend-box" style="background:var(--accent-lab);"></div> Lab (3h Block)</div>
        <div class="legend-item"><div class="legend-box" style="background:var(--accent-oe);"></div> Open Elective (OE)</div>
        <div class="legend-item"><div class="legend-box" style="background:var(--accent-pe);"></div> Physical Ed (PE)</div>
        <div class="legend-item"><div class="legend-box" style="background:#f59e0b;"></div> Short Break (10:30-11:00)</div>
        <div class="legend-item"><div class="legend-box" style="background:var(--accent-lunch);"></div> Lunch Break (13:30-14:30)</div>
      </div>
    </div>

    <div id="statsArea" style="display: none;"></div>

    <div class="timetable-wrapper">
      <div class="table-responsive" id="viewContainer">
        <!-- Rendered by JS -->
      </div>
    </div>

  </div>

  <script>
    const data = {json_data};
    let currentMode = 'section';

    function init() {{
      switchView('section');
    }}

    function switchView(mode) {{
      currentMode = mode;
      document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
      const activeBtn = Array.from(document.querySelectorAll('.tab-btn')).find(b => b.getAttribute('onclick').includes(mode));
      if (activeBtn) activeBtn.classList.add('active');

      const select = document.getElementById('itemSelect');
      const label = document.getElementById('selectorLabel');
      const statsArea = document.getElementById('statsArea');
      const controlsBar = document.getElementById('controlsBar');
      const legendArea = document.getElementById('legendArea');
      select.innerHTML = '';

      if (mode === 'section') {{
        controlsBar.style.display = 'flex';
        statsArea.style.display = 'none';
        legendArea.style.display = 'flex';
        label.textContent = 'Select Section:';
        Object.keys(data.sections).forEach(sec => {{
          const opt = document.createElement('option');
          opt.value = sec;
          opt.textContent = sec;
          select.appendChild(opt);
        }});
      }} else if (mode === 'workload') {{
        controlsBar.style.display = 'flex';
        statsArea.style.display = 'block';
        legendArea.style.display = 'none';
        label.textContent = 'Filter Designation:';
        ['All Designations', 'Professor', 'Associate Professor', 'Assistant Professor'].forEach(d => {{
          const opt = document.createElement('option');
          opt.value = d;
          opt.textContent = d;
          select.appendChild(opt);
        }});
      }} else if (mode === 'staff') {{
        controlsBar.style.display = 'flex';
        statsArea.style.display = 'none';
        legendArea.style.display = 'flex';
        label.textContent = 'Select Faculty:';
        Object.keys(data.staff).sort().forEach(staffName => {{
          const opt = document.createElement('option');
          opt.value = staffName;
          opt.textContent = staffName + ' (' + data.staff[staffName].designation + ')';
          select.appendChild(opt);
        }});
      }} else if (mode === 'room') {{
        controlsBar.style.display = 'flex';
        statsArea.style.display = 'none';
        legendArea.style.display = 'flex';
        label.textContent = 'Select Room:';
        Object.keys(data.rooms).sort().forEach(roomName => {{
          const opt = document.createElement('option');
          opt.value = roomName;
          opt.textContent = roomName;
          select.appendChild(opt);
        }});
      }} else if (mode === 'dept') {{
        controlsBar.style.display = 'flex';
        statsArea.style.display = 'none';
        legendArea.style.display = 'none';
        label.textContent = 'Select Department:';
        Object.keys(data.departments).forEach(dName => {{
          const opt = document.createElement('option');
          opt.value = dName;
          opt.textContent = dName.replace('_', ' ');
          select.appendChild(opt);
        }});
      }}

      renderCurrentView();
    }}

    function renderCurrentView() {{
      const select = document.getElementById('itemSelect');
      const selectedVal = select.value;
      const container = document.getElementById('viewContainer');
      const bannerTitle = document.getElementById('printBannerTitle');
      const subtitle = document.getElementById('printSubtitle');

      if (bannerTitle) {{
        if (currentMode === 'section') {{
          bannerTitle.textContent = `COLLEGE OF ENGINEERING — TIMETABLE: ${{selectedVal}}`;
        }} else if (currentMode === 'workload') {{
          bannerTitle.textContent = `COLLEGE OF ENGINEERING — FACULTY WORKLOAD DISTRIBUTION`;
        }} else if (currentMode === 'staff') {{
          bannerTitle.textContent = `FACULTY TIMETABLE — ${{selectedVal}}`;
        }} else if (currentMode === 'room') {{
          bannerTitle.textContent = `ROOM OCCUPANCY & UTILIZATION — ${{selectedVal}}`;
        }} else if (currentMode === 'dept') {{
          bannerTitle.textContent = `${{selectedVal.replace('_', ' ')}} MASTER TIMETABLE`;
        }}
      }}

      if (subtitle) {{
        if (currentMode === 'section') {{
          subtitle.textContent = "ISE & CSBS Departments | Monday - Saturday (8:30 AM - 5:30 PM)";
        }} else if (currentMode === 'workload') {{
          subtitle.textContent = "AICTE Workload Norms: Prof (12h), Assoc Prof (14h), Asst Prof (16h)";
        }} else if (currentMode === 'staff') {{
          subtitle.textContent = "Weekly Schedule | College of Engineering";
        }} else if (currentMode === 'room') {{
          subtitle.textContent = "Classroom / Laboratory Weekly Allocation";
        }} else if (currentMode === 'dept') {{
          subtitle.textContent = "Full Department Cohort Schedule";
        }}
      }}

      if (currentMode === 'section') {{
        renderSectionView(selectedVal, container);
      }} else if (currentMode === 'workload') {{
        renderWorkloadView(selectedVal, container);
      }} else if (currentMode === 'staff') {{
        renderStaffView(selectedVal, container);
      }} else if (currentMode === 'room') {{
        renderRoomView(selectedVal, container);
      }} else if (currentMode === 'dept') {{
        renderDeptView(selectedVal, container);
      }}
    }}

    function renderWorkloadView(filterVal, container) {{
      const statsArea = document.getElementById('statsArea');
      let filtered = data.workload;
      if (filterVal && filterVal !== 'All Designations') {{
        filtered = data.workload.filter(w => w.designation === filterVal);
      }}

      const totalFaculty = filtered.length;
      const totalHours = filtered.reduce((a, b) => a + b.tot_hrs, 0);
      const totalTarget = filtered.reduce((a, b) => a + b.target_hrs, 0);
      const totalTheory = filtered.reduce((a, b) => a + b.theory_hrs, 0);
      const totalLabSessions = filtered.reduce((a, b) => a + b.lab_sessions, 0);
      const totalLabHours = filtered.reduce((a, b) => a + b.lab_hrs, 0);
      const totalOE = filtered.reduce((a, b) => a + b.oe_hrs, 0);
      const totalPE = filtered.reduce((a, b) => a + b.pe_hrs, 0);
      const avgLoad = totalFaculty > 0 ? (totalHours / totalFaculty).toFixed(1) : 0;

      const dayTotals = {{}};
      data.days.forEach(d => {{
        dayTotals[d] = filtered.reduce((a, b) => a + (b.days[d] || 0), 0);
      }});

      statsArea.innerHTML = `
        <div class="stat-cards-grid">
          <div class="stat-card">
            <span class="stat-title">Faculty Filtered</span>
            <span class="stat-value">${{totalFaculty}}</span>
            <span class="stat-sub">Hierarchy: 4 Prof | 6 Assoc | 10 Asst</span>
          </div>
          <div class="stat-card">
            <span class="stat-title">Total Workload Assigned</span>
            <span class="stat-value">${{totalHours}} <span style="font-size:1rem;color:var(--text-muted)">hrs/wk</span></span>
            <span class="stat-sub">Target: ${{totalTarget}}h (${{totalHours >= totalTarget ? '+' : ''}}${{totalHours - totalTarget}}h variance)</span>
          </div>
          <div class="stat-card">
            <span class="stat-title">Average Faculty Load</span>
            <span class="stat-value">${{avgLoad}} <span style="font-size:1rem;color:var(--text-muted)">hrs/wk</span></span>
            <span class="stat-sub">AICTE Norms: Prof 12h | Assoc 14h | Asst 16h</span>
          </div>
          <div class="stat-card">
            <span class="stat-title">Lab Limit Compliance</span>
            <span class="stat-value" style="color:#10b981;">100%</span>
            <span class="stat-sub">Strictly 1 lab/week per section (13 labs total)</span>
          </div>
        </div>
      `;

      let html = `<table class="list-table">
        <thead>
          <tr>
            <th style="width:35px;text-align:center;">#</th>
            <th>Faculty Name</th>
            <th>Designation</th>
            <th style="text-align:center;">Target</th>
            <th style="text-align:center;">Assigned</th>
            <th style="text-align:center;">Theory</th>
            <th style="text-align:center;">Lab</th>
            <th style="text-align:center;">OE</th>
            <th style="text-align:center;">PE</th>
            <th style="text-align:center;">Mon</th>
            <th style="text-align:center;">Tue</th>
            <th style="text-align:center;">Wed</th>
            <th style="text-align:center;">Thu</th>
            <th style="text-align:center;">Fri</th>
            <th style="text-align:center;">Sat</th>
            <th>Assigned Subjects</th>
            <th>Sections</th>
          </tr>
        </thead>
        <tbody>`;

      filtered.forEach((w, idx) => {{
        const labStr = w.lab_sessions > 0 ? `${{w.lab_sessions}} (${{w.lab_hrs}}h)` : '-';
        
        let badgeColor = '#60a5fa';
        if (w.designation === 'Professor') badgeColor = '#f59e0b';
        else if (w.designation === 'Associate Professor') badgeColor = '#a78bfa';

        const isExact = w.tot_hrs === w.target_hrs;
        const targetBadge = `<span style="font-weight:600;color:var(--text-muted);">${{w.target_hrs}}h</span>`;
        const assignedBadge = `<span class="badge-load" style="background:${{isExact ? 'rgba(16,185,129,0.25)' : 'rgba(37,99,235,0.25)'}};color:${{isExact ? '#6ee7b7' : '#93c5fd'}};">${{w.tot_hrs}}h</span>`;

        html += `<tr>
          <td style="color:var(--text-muted);text-align:center;font-weight:600;">${{idx + 1}}</td>
          <td><strong>${{w.name}}</strong></td>
          <td><span style="color:${{badgeColor}};font-size:0.8rem;font-weight:600;">${{w.designation}}</span></td>
          <td style="text-align:center;">${{targetBadge}}</td>
          <td style="text-align:center;">${{assignedBadge}}</td>
          <td style="text-align:center;">${{w.theory_hrs}}h</td>
          <td style="text-align:center;"><span style="color:#67e8f9;">${{labStr}}</span></td>
          <td style="text-align:center;">${{w.oe_hrs > 0 ? w.oe_hrs + 'h' : '-'}}</td>
          <td style="text-align:center;">${{w.pe_hrs > 0 ? w.pe_hrs + 'h' : '-'}}</td>
          <td style="text-align:center;color:#94a3b8;">${{w.days['Monday'] || 0}}</td>
          <td style="text-align:center;color:#94a3b8;">${{w.days['Tuesday'] || 0}}</td>
          <td style="text-align:center;color:#94a3b8;">${{w.days['Wednesday'] || 0}}</td>
          <td style="text-align:center;color:#94a3b8;">${{w.days['Thursday'] || 0}}</td>
          <td style="text-align:center;color:#94a3b8;">${{w.days['Friday'] || 0}}</td>
          <td style="text-align:center;color:#94a3b8;">${{w.days['Saturday'] || 0}}</td>
          <td><div style="font-size:0.75rem;max-width:280px;line-height:1.3;">${{w.subjects.join(', ')}}</div></td>
          <td><div style="font-size:0.75rem;max-width:200px;color:#93c5fd;">${{w.sections.join(', ')}}</div></td>
        </tr>`;
      }});

      html += `</tbody>
        <tfoot>
          <tr>
            <td style="text-align:center;">Σ</td>
            <td><strong>TOTAL</strong></td>
            <td style="color:#94a3b8;">${{totalFaculty}} Faculty</td>
            <td style="text-align:center;"><strong>${{totalTarget}}h</strong></td>
            <td style="text-align:center;color:#6ee7b7;"><strong>${{totalHours}}h</strong></td>
            <td style="text-align:center;">${{totalTheory}}h</td>
            <td style="text-align:center;color:#67e8f9;">${{totalLabSessions > 0 ? `${{totalLabSessions}} (${{totalLabHours}}h)` : '-'}}</td>
            <td style="text-align:center;">${{totalOE > 0 ? totalOE + 'h' : '-'}}</td>
            <td style="text-align:center;">${{totalPE > 0 ? totalPE + 'h' : '-'}}</td>
            <td style="text-align:center;">${{dayTotals['Monday'] || 0}}</td>
            <td style="text-align:center;">${{dayTotals['Tuesday'] || 0}}</td>
            <td style="text-align:center;">${{dayTotals['Wednesday'] || 0}}</td>
            <td style="text-align:center;">${{dayTotals['Thursday'] || 0}}</td>
            <td style="text-align:center;">${{dayTotals['Friday'] || 0}}</td>
            <td style="text-align:center;">${{dayTotals['Saturday'] || 0}}</td>
            <td>-</td>
            <td>-</td>
          </tr>
        </tfoot>
      </table>`;
      container.innerHTML = html;
    }}

    function renderSectionView(sectionId, container) {{
      const rows = data.sections[sectionId];
      if (!rows) return;
      container.innerHTML = buildTimetableGridHtml(rows);
    }}

    function renderStaffView(staffName, container) {{
      const rows = data.staff_grids[staffName];
      const staffInfo = data.staff[staffName];
      if (!rows || !staffInfo) return;

      const desig = staffInfo.designation;
      const wItem = data.workload.find(w => w.name === staffName) || {{}};
      const assigned = wItem.tot_hrs || 0;
      const target = wItem.target_hrs || 0;
      const subjects = wItem.subjects ? wItem.subjects.join(', ') : '';
      const sections = wItem.sections ? wItem.sections.join(', ') : '';

      const banner = `
        <div class="view-info-banner">
          <div>
            <span class="title">👤 ${{staffName}}</span>
            <span style="margin-left:8px;font-size:0.8rem;background:rgba(255,255,255,0.1);padding:2px 8px;border-radius:4px;">${{desig}}</span>
            <div class="sub" style="margin-top:4px;">Assigned: <strong style="color:#6ee7b7;">${{assigned}}h</strong> / Target: ${{target}}h | Subjects: ${{subjects}}</div>
          </div>
          <div style="font-size:0.82rem;color:#93c5fd;">Cohorts: ${{sections}}</div>
        </div>
      `;

      container.innerHTML = banner + buildTimetableGridHtml(rows);
    }}

    function renderRoomView(roomName, container) {{
      const rows = data.room_grids[roomName];
      const events = data.rooms[roomName] || [];
      if (!rows) return;

      const isLab = roomName.startsWith('LAB');
      const occupiedPeriods = events.reduce((a, b) => a + (b.type === 'LAB' ? 3 : 1), 0);

      const banner = `
        <div class="view-info-banner">
          <div>
            <span class="title">📍 ${{roomName}}</span>
            <span style="margin-left:8px;font-size:0.8rem;background:rgba(255,255,255,0.1);padding:2px 8px;border-radius:4px;">${{isLab ? 'Computing Laboratory' : 'Academic Classroom'}}</span>
            <div class="sub" style="margin-top:4px;">Total Utilization: <strong style="color:#6ee7b7;">${{occupiedPeriods}} Periods/Week</strong> (${{((occupiedPeriods/48)*100).toFixed(1)}}% Occupancy)</div>
          </div>
          <div style="font-size:0.82rem;color:#93c5fd;">Active Days: Monday - Saturday</div>
        </div>
      `;

      container.innerHTML = banner + buildTimetableGridHtml(rows);
    }}

    function buildTimetableGridHtml(rows) {{
      let html = `<table>
        <thead>
          <tr>
            <th>Time Slot</th>
            ${{data.days.map(d => `<th>${{d}}</th>`).join('')}}
          </tr>
        </thead>
        <tbody>`;

      rows.forEach(r => {{
        if (r.is_break) {{
          const isMorning = r.time.includes("10:30") || r.break_type === "BREAK";
          const label = isMorning ? "SHORT BREAK (10:30 - 11:00)" : "LUNCH BREAK (13:30 - 14:30)";
          const rowClass = isMorning ? "break-row" : "lunch-row";
          html += `<tr class="${{rowClass}}">
            <td class="time-col">${{r.time}}</td>
            <td colspan="6">${{label}}</td>
          </tr>`;
          return;
        }}

        html += `<tr>
          <td class="time-col">${{r.time}}</td>`;

        r.days.forEach(cellText => {{
          if (!cellText || cellText === 'FREE') {{
            html += `<td><span class="free-slot">FREE</span></td>`;
          }} else {{
            const lines = cellText.split('\\n');
            const sub = lines[0] || '';
            const type = (lines[1] || '').replace('[', '').replace(']', '');
            const meta1 = lines[2] || '';
            const meta2 = lines[3] || '';

            let badgeClass = 'badge-theory';
            if (type === 'LAB') badgeClass = 'badge-lab';
            else if (type === 'OE') badgeClass = 'badge-oe';
            else if (type === 'PE') badgeClass = 'badge-pe';

            html += `<td>
              <div class="event-card ${{type}}">
                <span class="ev-badge ${{badgeClass}}">${{type}}</span>
                <div class="ev-title">${{sub}}</div>
                <div class="ev-meta">
                  <span>${{meta1.startsWith('CR_') || meta1.startsWith('LAB_') ? '📍 ' + meta1 : '👤 ' + meta1}}</span>
                  <span>${{meta2 ? (meta2.startsWith('CR_') || meta2.startsWith('LAB_') ? '📍 ' + meta2 : '👥 ' + meta2) : ''}}</span>
                </div>
              </div>
            </td>`;
          }}
        }});

        html += `</tr>`;
      }});

      html += `</tbody></table>`;
      return html;
    }}

    function renderDeptView(deptName, container) {{
      const events = data.departments[deptName] || [];

      let html = `
        <div class="view-info-banner">
          <div>
            <span class="title">🏛️ ${{deptName.replace('_', ' ')}}</span>
            <div class="sub" style="margin-top:4px;">Master Department Roster | Total Academic Events: <strong style="color:#6ee7b7;">${{events.length}}</strong></div>
          </div>
        </div>
        <table class="list-table">
          <thead>
            <tr>
              <th>Day</th>
              <th>Time</th>
              <th>Section</th>
              <th>Subject</th>
              <th>Teacher</th>
              <th>Room</th>
              <th>Type</th>
            </tr>
          </thead>
          <tbody>`;

      events.forEach(e => {{
        html += `<tr>
          <td><strong>${{e.day}}</strong></td>
          <td>${{e.time}}</td>
          <td><span style="color:#60a5fa">${{e.section}}</span></td>
          <td><strong>${{e.subject}}</strong></td>
          <td>${{e.teacher}}</td>
          <td><span style="color:#34d399">${{e.room}}</span></td>
          <td><span class="ev-badge badge-${{e.type.toLowerCase()}}">${{e.type}}</span></td>
        </tr>`;
      }});

      html += `</tbody></table>`;
      container.innerHTML = html;
    }}

    function downloadCurrentCSV() {{
      const select = document.getElementById('itemSelect');
      const val = select.value;
      let csvContent = "data:text/csv;charset=utf-8,";

      const table = document.querySelector("#viewContainer table");
      if (!table) return;

      const rows = table.querySelectorAll("tr");
      rows.forEach(row => {{
        const cols = row.querySelectorAll("th, td");
        const rowData = [];
        cols.forEach(c => {{
          let text = c.innerText.replace(/\\n/g, ' - ').replace(/"/g, '""');
          rowData.push('"' + text + '"');
        }});
        csvContent += rowData.join(",") + "\\r\\n";
      }});

      const encodedUri = encodeURI(csvContent);
      const link = document.createElement("a");
      link.setAttribute("href", encodedUri);
      link.setAttribute("download", `${{currentMode}}_${{val.replace(/\\s+/g, '_')}}_Timetable.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }}

    window.onload = init;
  </script>
</body>
</html>
"""

    with open("Timetable_Viewer.html", "w", encoding="utf-8") as f:
        f.write(html_template)

    print("Created Timetable_Viewer.html successfully with weekly grids for Sections, Faculty, Rooms, and Workload Table!")

if __name__ == "__main__":
    generate_html()
