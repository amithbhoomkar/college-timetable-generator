import openpyxl
import json
import os
from collections import defaultdict

def generate_html():
    wb = openpyxl.load_workbook("Complete_Timetable.xlsx")

    # Extract all section data
    section_names = [s for s in wb.sheetnames if s.startswith("ISE_S") or s.startswith("CSBS_S")]
    sections_data = {}

    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

    for s_name in section_names:
        ws = wb[s_name]
        rows = []
        for r in range(4, ws.max_row + 1):
            time_slot = ws.cell(row=r, column=1).value
            if not time_slot:
                continue
            day_cells = []
            is_lunch_row = False
            for c in range(2, 8):
                val = ws.cell(row=r, column=c).value or ""
                day_cells.append(str(val))
                if "LUNCH" in str(val):
                    is_lunch_row = True
            rows.append({
                "time": str(time_slot),
                "is_lunch": is_lunch_row,
                "days": day_cells
            })
        sections_data[s_name] = rows

    # Extract Staff data
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

    sorted_staff_names = sorted(staff_dict.keys(), key=sort_key)
    workload_data = []
    for t in sorted_staff_names:
        d = staff_dict[t]
        workload_data.append({
            "name": d['name'],
            "designation": d['designation'],
            "tot_hrs": d['tot_hrs'],
            "theory_hrs": d['theory_hrs'],
            "lab_sessions": d['lab_sessions'],
            "lab_hrs": d['lab_hrs'],
            "oe_hrs": d['oe_hrs'],
            "pe_hrs": d['pe_hrs'],
            "days": {day: d['days'][day] for day in days},
            "subjects": sorted(list(d['subjects'])),
            "sections": sorted(list(d['sections']))
        })

    # Extract Room data
    room_ws = wb["ROOM_UTILIZATION"]
    room_data = {}
    for r in range(2, room_ws.max_row + 1):
        row_vals = [room_ws.cell(row=r, column=c).value for c in range(1, 8)]
        room = str(row_vals[0])
        if room not in room_data:
            room_data[room] = []
        room_data[room].append({
            "day": str(row_vals[1]),
            "time": str(row_vals[2]),
            "subject": str(row_vals[3]),
            "section": str(row_vals[4]),
            "teacher": str(row_vals[5]),
            "type": str(row_vals[6])
        })

    # Department Master data
    dept_data = {}
    for d_name in ["ISE_MASTER", "CSBS_MASTER"]:
        ws = wb[d_name]
        d_events = []
        for r in range(4, ws.max_row + 1):
            row_vals = [ws.cell(row=r, column=c).value for c in range(1, 8)]
            if not row_vals[0]:
                continue
            d_events.append({
                "day": str(row_vals[0]),
                "time": str(row_vals[1]),
                "section": str(row_vals[2]),
                "subject": str(row_vals[3]),
                "type": str(row_vals[4]),
                "teacher": str(row_vals[5]),
                "room": str(row_vals[6])
            })
        dept_data[d_name] = d_events

    data_bundle = {
        "sections": sections_data,
        "staff": staff_data,
        "workload": workload_data,
        "rooms": room_data,
        "departments": dept_data,
        "days": days
    }

    json_data = json.dumps(data_bundle)

    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>College Timetable Interactive Viewer | ISE & CSBS</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root {{
      --primary: #2563eb;
      --primary-dark: #1d4ed8;
      --primary-glow: rgba(37, 99, 235, 0.2);
      --bg: #0f172a;
      --card-bg: rgba(30, 41, 59, 0.85);
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
      background: rgba(15, 23, 42, 0.8);
      backdrop-filter: blur(12px);
      border-bottom: 1px solid var(--card-border);
      padding: 1.2rem 2.5rem;
      display: flex;
      justify-content: space-between;
      align-items: center;
      position: sticky;
      top: 0;
      z-index: 100;
    }}

    .logo-area h1 {{
      font-family: 'Outfit', sans-serif;
      font-size: 1.5rem;
      font-weight: 700;
      background: linear-gradient(135deg, #60a5fa, #a78bfa);
      -webkit-background-clip: text;
      background-clip: text;
      -webkit-text-fill-color: transparent;
      letter-spacing: -0.5px;
    }}

    .logo-area p {{
      color: var(--text-muted);
      font-size: 0.85rem;
      margin-top: 0.2rem;
    }}

    .actions {{
      display: flex;
      gap: 0.8rem;
    }}

    .btn {{
      padding: 0.55rem 1.2rem;
      border-radius: 8px;
      font-weight: 500;
      font-size: 0.875rem;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 0.5rem;
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
      background: rgba(255, 255, 255, 0.05);
      color: var(--text);
      border: 1px solid var(--card-border);
    }}
    .btn-secondary:hover {{
      background: rgba(255, 255, 255, 0.1);
    }}

    .main-container {{
      flex: 1;
      padding: 2rem 2.5rem;
      max-width: 1700px;
      margin: 0 auto;
      width: 100%;
    }}

    .nav-tabs {{
      display: flex;
      gap: 0.5rem;
      background: rgba(15, 23, 42, 0.6);
      padding: 0.4rem;
      border-radius: 12px;
      border: 1px solid var(--card-border);
      margin-bottom: 1.5rem;
      width: fit-content;
      flex-wrap: wrap;
    }}

    .tab-btn {{
      padding: 0.6rem 1.4rem;
      border-radius: 8px;
      background: transparent;
      border: none;
      color: var(--text-muted);
      font-weight: 500;
      font-size: 0.9rem;
      cursor: pointer;
      transition: all 0.2s ease;
      font-family: 'Inter', sans-serif;
    }}

    .tab-btn:hover {{
      color: var(--text);
    }}

    .tab-btn.active {{
      background: var(--primary);
      color: white;
      box-shadow: 0 2px 10px var(--primary-glow);
    }}

    .controls-bar {{
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 12px;
      padding: 1rem 1.5rem;
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 1.5rem;
      flex-wrap: wrap;
      gap: 1rem;
    }}

    .select-group {{
      display: flex;
      align-items: center;
      gap: 0.8rem;
    }}

    .select-group label {{
      font-size: 0.9rem;
      color: var(--text-muted);
      font-weight: 500;
    }}

    select {{
      background: #1e293b;
      color: white;
      border: 1px solid rgba(255, 255, 255, 0.15);
      padding: 0.55rem 1.2rem;
      border-radius: 8px;
      font-size: 0.9rem;
      outline: none;
      cursor: pointer;
    }}
    select:focus {{
      border-color: var(--primary);
      box-shadow: 0 0 0 2px var(--primary-glow);
    }}

    .legend {{
      display: flex;
      gap: 1rem;
      flex-wrap: wrap;
      font-size: 0.8rem;
    }}
    .legend-item {{
      display: flex;
      align-items: center;
      gap: 0.4rem;
    }}
    .legend-box {{
      width: 12px;
      height: 12px;
      border-radius: 3px;
    }}

    .timetable-wrapper {{
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 14px;
      overflow: hidden;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
    }}

    .table-responsive {{
      overflow-x: auto;
    }}

    table {{
      width: 100%;
      border-collapse: collapse;
      text-align: center;
    }}

    th {{
      background: #1e293b;
      color: #93c5fd;
      font-family: 'Outfit', sans-serif;
      font-weight: 600;
      padding: 1rem 0.75rem;
      font-size: 0.875rem;
      border-bottom: 2px solid rgba(255, 255, 255, 0.1);
      border-right: 1px solid rgba(255, 255, 255, 0.05);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }}

    td {{
      padding: 0.75rem 0.5rem;
      border-bottom: 1px solid rgba(255, 255, 255, 0.05);
      border-right: 1px solid rgba(255, 255, 255, 0.05);
      font-size: 0.825rem;
      vertical-align: middle;
      min-width: 160px;
    }}

    td.time-col {{
      background: #1e293b;
      color: #f1f5f9;
      font-family: 'Outfit', sans-serif;
      font-weight: 600;
      font-size: 0.85rem;
      min-width: 120px;
    }}

    .event-card {{
      padding: 0.6rem 0.5rem;
      border-radius: 8px;
      text-align: left;
      font-size: 0.78rem;
      position: relative;
      display: flex;
      flex-direction: column;
      gap: 0.25rem;
      transition: transform 0.15s ease, box-shadow 0.15s ease;
      background: rgba(255, 255, 255, 0.03);
      border: 1px solid rgba(255, 255, 255, 0.06);
    }}

    .event-card:hover {{
      transform: translateY(-2px);
      box-shadow: 0 4px 12px rgba(0,0,0,0.25);
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
      display: inline-block;
      font-size: 0.65rem;
      font-weight: 700;
      padding: 0.1rem 0.4rem;
      border-radius: 4px;
      width: fit-content;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }}
    .badge-theory {{ background: rgba(99, 102, 241, 0.25); color: #a5b4fc; }}
    .badge-lab {{ background: rgba(6, 182, 212, 0.25); color: #67e8f9; }}
    .badge-oe {{ background: rgba(245, 158, 11, 0.25); color: #fde68a; }}
    .badge-pe {{ background: rgba(16, 185, 129, 0.25); color: #6ee7b7; }}

    .ev-title {{
      font-weight: 600;
      color: #f8fafc;
      line-height: 1.25;
    }}

    .ev-meta {{
      display: flex;
      flex-direction: column;
      gap: 0.15rem;
      color: var(--text-muted);
      font-size: 0.72rem;
    }}

    .free-slot {{
      color: rgba(255, 255, 255, 0.15);
      font-weight: 500;
      font-size: 0.8rem;
    }}

    .lunch-row td {{
      background: rgba(100, 116, 139, 0.15);
      color: #94a3b8;
      font-weight: 600;
      letter-spacing: 2px;
      font-size: 0.8rem;
      padding: 0.6rem;
      text-transform: uppercase;
    }}

    /* Stat Cards for Workload */
    .stat-cards-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 1rem;
      margin-bottom: 1.5rem;
    }}

    .stat-card {{
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: 12px;
      padding: 1.2rem;
      display: flex;
      flex-direction: column;
      gap: 0.4rem;
    }}
    .stat-title {{
      font-size: 0.8rem;
      color: var(--text-muted);
      font-weight: 500;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }}
    .stat-value {{
      font-family: 'Outfit', sans-serif;
      font-size: 1.8rem;
      font-weight: 700;
      color: #93c5fd;
    }}
    .stat-sub {{
      font-size: 0.75rem;
      color: #10b981;
      font-weight: 500;
    }}

    /* List Table Styles */
    .list-table th {{
      text-align: left;
    }}
    .list-table td {{
      text-align: left;
      min-width: unset;
      padding: 0.85rem 1rem;
    }}
    .list-table tr:hover td {{
      background: rgba(255, 255, 255, 0.02);
    }}

    .badge-load {{
      background: rgba(37, 99, 235, 0.25);
      color: #93c5fd;
      padding: 0.2rem 0.6rem;
      border-radius: 6px;
      font-weight: 700;
      font-size: 0.85rem;
      display: inline-block;
    }}

    @media print {{
      body {{ background: white !important; color: black !important; }}
      header, .controls-bar, .nav-tabs {{ display: none !important; }}
      .timetable-wrapper {{ border: 1px solid #ccc !important; box-shadow: none !important; }}
      th, td {{ color: black !important; border: 1px solid #ccc !important; }}
      td.time-col {{ background: #eee !important; color: black !important; }}
      .event-card {{ background: none !important; border: 1px solid #ddd !important; }}
      .ev-title {{ color: black !important; }}
      .ev-meta {{ color: #555 !important; }}
    }}
  </style>
</head>
<body>

  <header>
    <div class="logo-area">
      <h1>College Timetable Interactive Viewer</h1>
      <p>ISE & CSBS Departments | Generated via Google OR-Tools CP-SAT</p>
    </div>
    <div class="actions">
      <button class="btn btn-secondary" onclick="window.print()">Print / Save PDF</button>
      <button class="btn btn-primary" onclick="downloadCurrentCSV()">Download CSV</button>
    </div>
  </header>

  <div class="main-container">
    
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
        <div class="legend-item"><div class="legend-box" style="background:var(--accent-theory);"></div> Theory</div>
        <div class="legend-item"><div class="legend-box" style="background:var(--accent-lab);"></div> Lab (3h)</div>
        <div class="legend-item"><div class="legend-box" style="background:var(--accent-oe);"></div> Open Elective (OE)</div>
        <div class="legend-item"><div class="legend-box" style="background:var(--accent-pe);"></div> Physical Ed (PE)</div>
        <div class="legend-item"><div class="legend-box" style="background:var(--accent-lunch);"></div> Lunch Break</div>
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
      event.target.classList.add('active');

      const select = document.getElementById('itemSelect');
      const label = document.getElementById('selectorLabel');
      const statsArea = document.getElementById('statsArea');
      const controlsBar = document.getElementById('controlsBar');
      select.innerHTML = '';

      if (mode === 'section') {{
        controlsBar.style.display = 'flex';
        statsArea.style.display = 'none';
        label.textContent = 'Select Section:';
        Object.keys(data.sections).forEach(sec => {{
          const opt = document.createElement('option');
          opt.value = sec;
          opt.textContent = sec;
          select.appendChild(opt);
        }});
        document.getElementById('legendArea').style.display = 'flex';
      }} else if (mode === 'workload') {{
        controlsBar.style.display = 'flex';
        statsArea.style.display = 'block';
        label.textContent = 'Filter Designation:';
        ['All Designations', 'Professor', 'Associate Professor', 'Assistant Professor'].forEach(d => {{
          const opt = document.createElement('option');
          opt.value = d;
          opt.textContent = d;
          select.appendChild(opt);
        }});
        document.getElementById('legendArea').style.display = 'none';
      }} else if (mode === 'staff') {{
        controlsBar.style.display = 'flex';
        statsArea.style.display = 'none';
        label.textContent = 'Select Faculty:';
        Object.keys(data.staff).sort().forEach(staffName => {{
          const opt = document.createElement('option');
          opt.value = staffName;
          opt.textContent = staffName + ' (' + data.staff[staffName].designation + ')';
          select.appendChild(opt);
        }});
        document.getElementById('legendArea').style.display = 'none';
      }} else if (mode === 'room') {{
        controlsBar.style.display = 'flex';
        statsArea.style.display = 'none';
        label.textContent = 'Select Room:';
        Object.keys(data.rooms).sort().forEach(roomName => {{
          const opt = document.createElement('option');
          opt.value = roomName;
          opt.textContent = roomName;
          select.appendChild(opt);
        }});
        document.getElementById('legendArea').style.display = 'none';
      }} else if (mode === 'dept') {{
        controlsBar.style.display = 'flex';
        statsArea.style.display = 'none';
        label.textContent = 'Select Department:';
        Object.keys(data.departments).forEach(dName => {{
          const opt = document.createElement('option');
          opt.value = dName;
          opt.textContent = dName.replace('_', ' ');
          select.appendChild(opt);
        }});
        document.getElementById('legendArea').style.display = 'none';
      }}

      renderCurrentView();
    }}

    function renderCurrentView() {{
      const select = document.getElementById('itemSelect');
      const selectedVal = select.value;
      const container = document.getElementById('viewContainer');

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

      const totalFaculty = data.workload.length;
      const totalHours = data.workload.reduce((a, b) => a + b.tot_hrs, 0);
      const avgLoad = (totalHours / totalFaculty).toFixed(1);

      statsArea.innerHTML = `
        <div class="stat-cards-grid">
          <div class="stat-card">
            <span class="stat-title">Total Faculty Members</span>
            <span class="stat-value">${{totalFaculty}}</span>
            <span class="stat-sub">4 Profs | 6 Assoc | 10 Asst</span>
          </div>
          <div class="stat-card">
            <span class="stat-title">Total Teaching Load</span>
            <span class="stat-value">${{totalHours}} <span style="font-size:1rem;color:var(--text-muted)">hrs/wk</span></span>
            <span class="stat-sub">275 Theory + 39 Lab + 20 OE/PE</span>
          </div>
          <div class="stat-card">
            <span class="stat-title">Average Faculty Load</span>
            <span class="stat-value">${{avgLoad}} <span style="font-size:1rem;color:var(--text-muted)">hrs/wk</span></span>
            <span class="stat-sub">Strict range: 12 - 22 hrs</span>
          </div>
          <div class="stat-card">
            <span class="stat-title">Quota Compliance</span>
            <span class="stat-value" style="color:#10b981;">100%</span>
            <span class="stat-sub">Max 3 Subjs & 3 Secs per Faculty</span>
          </div>
        </div>
      `;

      let html = `<table class="list-table">
        <thead>
          <tr>
            <th style="width:45px;">#</th>
            <th>Faculty Name</th>
            <th>Designation</th>
            <th style="text-align:center;">Total Hours</th>
            <th style="text-align:center;">Theory</th>
            <th style="text-align:center;">Lab</th>
            <th style="text-align:center;">OE</th>
            <th style="text-align:center;">PE</th>
            <th style="text-align:center;">Daily (M-S)</th>
            <th>Assigned Subjects</th>
            <th>Sections</th>
          </tr>
        </thead>
        <tbody>`;

      filtered.forEach((w, idx) => {{
        const dailyStr = data.days.map(d => w.days[d] || 0).join('-');
        const labStr = w.lab_sessions > 0 ? `${{w.lab_sessions}} (${{w.lab_hrs}}h)` : '-';
        
        let badgeColor = '#60a5fa';
        if (w.designation === 'Professor') badgeColor = '#f59e0b';
        else if (w.designation === 'Associate Professor') badgeColor = '#a78bfa';

        html += `<tr>
          <td style="color:var(--text-muted);font-weight:600;">${{idx + 1}}</td>
          <td><strong>${{w.name}}</strong></td>
          <td><span style="color:${{badgeColor}};font-size:0.8rem;font-weight:600;">${{w.designation}}</span></td>
          <td style="text-align:center;"><span class="badge-load">${{w.tot_hrs}}h</span></td>
          <td style="text-align:center;">${{w.theory_hrs}}h</td>
          <td style="text-align:center;"><span style="color:#67e8f9;">${{labStr}}</span></td>
          <td style="text-align:center;">${{w.oe_hrs > 0 ? w.oe_hrs + 'h' : '-'}}</td>
          <td style="text-align:center;">${{w.pe_hrs > 0 ? w.pe_hrs + 'h' : '-'}}</td>
          <td style="text-align:center;font-family:monospace;color:#94a3b8;">${{dailyStr}}</td>
          <td><div style="font-size:0.75rem;max-width:320px;line-height:1.3;">${{w.subjects.join(', ')}}</div></td>
          <td><div style="font-size:0.75rem;max-width:240px;color:#93c5fd;">${{w.sections.join(', ')}}</div></td>
        </tr>`;
      }});

      html += `</tbody></table>`;
      container.innerHTML = html;
    }}

    function renderSectionView(sectionId, container) {{
      const rows = data.sections[sectionId];
      if (!rows) return;

      let html = `<table>
        <thead>
          <tr>
            <th>Time Slot</th>
            ${{data.days.map(d => `<th>${{d}}</th>`).join('')}}
          </tr>
        </thead>
        <tbody>`;

      rows.forEach(r => {{
        if (r.is_lunch) {{
          html += `<tr class="lunch-row">
            <td class="time-col">${{r.time}}</td>
            <td colspan="6">LUNCH BREAK (13:30 - 14:30)</td>
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
            const teacher = lines[2] || '';
            const room = lines[3] || '';

            let badgeClass = 'badge-theory';
            if (type === 'LAB') badgeClass = 'badge-lab';
            else if (type === 'OE') badgeClass = 'badge-oe';
            else if (type === 'PE') badgeClass = 'badge-pe';

            html += `<td>
              <div class="event-card ${{type}}">
                <span class="ev-badge ${{badgeClass}}">${{type}}</span>
                <div class="ev-title">${{sub}}</div>
                <div class="ev-meta">
                  <span>👤 ${{teacher}}</span>
                  <span>📍 ${{room}}</span>
                </div>
              </div>
            </td>`;
          }}
        }});

        html += `</tr>`;
      }});

      html += `</tbody></table>`;
      container.innerHTML = html;
    }}

    function renderStaffView(staffName, container) {{
      const staffInfo = data.staff[staffName];
      if (!staffInfo) return;

      const sortedEvents = [...staffInfo.events].sort((a, b) => {{
        const dayOrder = data.days.indexOf(a.day) - data.days.indexOf(b.day);
        if (dayOrder !== 0) return dayOrder;
        return a.time.localeCompare(b.time);
      }});

      let html = `<table class="list-table">
        <thead>
          <tr>
            <th>Day</th>
            <th>Time</th>
            <th>Subject</th>
            <th>Section</th>
            <th>Room</th>
            <th>Type</th>
          </tr>
        </thead>
        <tbody>`;

      sortedEvents.forEach(e => {{
        html += `<tr>
          <td><strong>${{e.day}}</strong></td>
          <td>${{e.time}}</td>
          <td><strong>${{e.subject}}</strong></td>
          <td><span style="color:#60a5fa">${{e.section}}</span></td>
          <td><span style="color:#34d399">${{e.room}}</span></td>
          <td><span class="ev-badge badge-${{e.type.toLowerCase()}}">${{e.type}}</span></td>
        </tr>`;
      }});

      html += `</tbody></table>`;
      container.innerHTML = html;
    }}

    function renderRoomView(roomName, container) {{
      const events = data.rooms[roomName] || [];
      const sortedEvents = [...events].sort((a, b) => {{
        const dayOrder = data.days.indexOf(a.day) - data.days.indexOf(b.day);
        if (dayOrder !== 0) return dayOrder;
        return a.time.localeCompare(b.time);
      }});

      let html = `<table class="list-table">
        <thead>
          <tr>
            <th>Day</th>
            <th>Time</th>
            <th>Subject</th>
            <th>Section</th>
            <th>Teacher</th>
            <th>Type</th>
          </tr>
        </thead>
        <tbody>`;

      sortedEvents.forEach(e => {{
        html += `<tr>
          <td><strong>${{e.day}}</strong></td>
          <td>${{e.time}}</td>
          <td><strong>${{e.subject}}</strong></td>
          <td><span style="color:#60a5fa">${{e.section}}</span></td>
          <td>${{e.teacher}}</td>
          <td><span class="ev-badge badge-${{e.type.toLowerCase()}}">${{e.type}}</span></td>
        </tr>`;
      }});

      html += `</tbody></table>`;
      container.innerHTML = html;
    }}

    function renderDeptView(deptName, container) {{
      const events = data.departments[deptName] || [];

      let html = `<table class="list-table">
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

    print("Created Timetable_Viewer.html successfully with Faculty Workload Sheet view!")

if __name__ == "__main__":
    generate_html()
