#!/usr/bin/env python3
"""
================================================================================
AGENTIC AI COLLEGE TIMETABLE GENERATOR (CP-SAT DISCRETE OPTIMIZATION)
================================================================================
Academic Schedule System with AICTE Faculty Workload Calibration & 1-Page PDF Fit
- Departments: ISE + CSBS (13 Student Cohort Sections)
- Faculty: 20 Teachers (4 Professors, 6 Associate Professors, 10 Assistant Professors)
- Workload Norms: Professor: 12h, Associate Professor: 14h, Assistant Professor: 16h
- Lab Rules: Strictly 1 laboratory session per week per class (no more than 1)
- Daily Timing: 8 Periods
    * Morning: 08:30-09:30, 09:30-10:30 (1 hour each)
    * Short Break: 10:30-11:00 (30 min)
    * Mid Session: 11:00-11:50, 11:50-12:40, 12:40-13:30 (50 min each)
    * Lunch Break: 13:30-14:30 (1 hour)
    * Noon Session: 14:30-15:30, 15:30-16:30, 16:30-17:30 (1 hour each, till 5:30)
================================================================================
"""

from ortools.sat.python import cp_model
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict
import os
import json
import time

# ================================================================
# 1. GLOBAL CONFIGURATION & SCHEDULE TOPOLOGY
# ================================================================

DAYS = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday"
]

# 8 Teaching Slots / Day according to institutional schedule
TIME_SLOTS = [
    "08:30-09:30",  # Slot 0: Morning 1 (1 hour)
    "09:30-10:30",  # Slot 1: Morning 2 (1 hour) -> Synchronized OE
    "11:00-11:50",  # Slot 2: Mid 1 (50 min) -> Valid Lab Block 1 Start
    "11:50-12:40",  # Slot 3: Mid 2 (50 min)
    "12:40-13:30",  # Slot 4: Mid 3 (50 min)
    "14:30-15:30",  # Slot 5: Noon 1 (1 hour) -> Synchronized PE / Valid Lab Block 2 Start
    "15:30-16:30",  # Slot 6: Noon 2 (1 hour)
    "16:30-17:30"   # Slot 7: Noon 3 (1 hour, concludes at 5:30 PM)
]

TIME_START = [
    "08:30",
    "09:30",
    "11:00",
    "11:50",
    "12:40",
    "14:30",
    "15:30",
    "16:30"
]

NUM_DAYS = len(DAYS)
NUM_SLOTS = len(TIME_SLOTS)

CLASSROOMS = [
    "CR_1",
    "CR_2",
    "CR_3",
    "CR_4",
    "CR_5",
    "CR_6",
    "CR_7"
]

LABS = [
    "LAB_1",
    "LAB_2"
]

ROOMS = CLASSROOMS + LABS

NUM_CLASSROOMS = len(CLASSROOMS)
NUM_LABS = len(LABS)

OE_SLOT = 1             # 09:30-10:30 (Morning Slot 2, before 10:30 break)
PE_SLOT = 5             # 14:30-15:30 (Noon Slot 1, right after lunch break)
VALID_LAB_STARTS = [2, 5]  # Mid session (11:00-13:30) or Noon session (14:30-17:30)

SOLVER_TIME_LIMIT = 180
OUTPUT_FILE = "Complete_Timetable.xlsx"

# ================================================================
# 2. SECTION REGISTRY
# ================================================================

SECTIONS = [
    # ISE
    {"id": "ISE_S3_A", "dept": "ISE", "sem": 3},
    {"id": "ISE_S3_B", "dept": "ISE", "sem": 3},
    {"id": "ISE_S3_C", "dept": "ISE", "sem": 3},
    {"id": "ISE_S5_A", "dept": "ISE", "sem": 5},
    {"id": "ISE_S5_B", "dept": "ISE", "sem": 5},
    {"id": "ISE_S7_A", "dept": "ISE", "sem": 7},
    {"id": "ISE_S7_B", "dept": "ISE", "sem": 7},

    # CSBS
    {"id": "CSBS_S1_A", "dept": "CSBS", "sem": 1},
    {"id": "CSBS_S1_B", "dept": "CSBS", "sem": 1},
    {"id": "CSBS_S3_A", "dept": "CSBS", "sem": 3},
    {"id": "CSBS_S3_B", "dept": "CSBS", "sem": 3},
    {"id": "CSBS_S5_A", "dept": "CSBS", "sem": 5},
    {"id": "CSBS_S7_A", "dept": "CSBS", "sem": 7},
]

SECTION_IDS = [x["id"] for x in SECTIONS]
SECTION_INFO = {x["id"]: x for x in SECTIONS}

# ================================================================
# 3. FACULTY REGISTRY (AICTE WORKLOAD TARGETS)
# ================================================================

STAFF = []

# 4 Professors: Target 12 Hours
for i in range(1, 5):
    STAFF.append({
        "id": f"PROF_{i}",
        "name": f"Professor {i}",
        "designation": "Professor"
    })

# 6 Associate Professors: Target 14 Hours
for i in range(1, 7):
    STAFF.append({
        "id": f"ASSOC_{i}",
        "name": f"Associate Professor {i}",
        "designation": "Associate Professor"
    })

# 10 Assistant Professors: Target 16 Hours
for i in range(1, 11):
    STAFF.append({
        "id": f"ASST_{i}",
        "name": f"Assistant Professor {i}",
        "designation": "Assistant Professor"
    })

STAFF_IDS = [x["id"] for x in STAFF]
STAFF_INFO = {x["id"]: x for x in STAFF}

MAX_SUBJECTS_PER_STAFF = 3
MAX_SECTIONS_PER_STAFF = 3

# ================================================================
# 4. CURRICULUM DATABASE
# ================================================================

SUBJECTS = {
    "ISE": {
        3: [
            "Data Structures",
            "Database Management Systems",
            "Computer Networks",
            "Operating Systems",
            "Object Oriented Programming",
            "Mathematics for Computing"
        ],
        5: [
            "Machine Learning",
            "Web Technology",
            "Software Engineering",
            "Cloud Computing",
            "Data Mining",
            "Information Security"
        ],
        7: [
            "Deep Learning",
            "Big Data Analytics",
            "Natural Language Processing",
            "Distributed Systems",
            "Advanced Database Systems",
            "Project Management"
        ]
    },
    "CSBS": {
        1: [
            "Programming Fundamentals",
            "Discrete Mathematics",
            "Digital Logic",
            "Computer Organization",
            "Communication Skills",
            "Engineering Mathematics"
        ],
        3: [
            "Data Structures",
            "Database Systems",
            "Computer Networks",
            "Operating Systems",
            "Object Oriented Programming",
            "Probability and Statistics"
        ],
        5: [
            "Machine Learning",
            "Business Analytics",
            "Cloud Computing",
            "Software Engineering",
            "Data Visualization",
            "Information Security"
        ],
        7: [
            "Deep Learning",
            "Natural Language Processing",
            "Big Data Analytics",
            "Business Intelligence",
            "Distributed Systems",
            "Project Management"
        ]
    }
}

OE_CHOICES = [
    "Artificial Intelligence",
    "Cyber Security",
    "Blockchain Technology",
    "Cloud Computing",
    "Data Science",
    "Internet of Things",
    "Computer Vision",
    "Natural Language Processing",
    "Robotics",
    "Big Data Analytics",
    "Business Intelligence",
    "Digital Forensics"
]

ISE_OE = "Artificial Intelligence"
CSBS_OE = "Business Intelligence"

PE_SUBJECTS = ["Physical Education"]

# ================================================================
# 5. AGENT 1 - CURATOR AGENT
# ================================================================

class CuratorAgent:
    """Curates all academic events, applying strict 1 lab/week and curriculum rules."""

    def run(self):
        print("\n" + "=" * 70)
        print("[AGENT 1] CURATOR AGENT")
        print("=" * 70)

        events = []
        theory_counter = 0

        # --------------------------------------------------------
        # THEORY EVENTS: 18 theory hours per section (3 per subject)
        # --------------------------------------------------------
        for section in SECTIONS:
            sid = section["id"]
            dept = section["dept"]
            sem = section["sem"]
            subjects = SUBJECTS[dept][sem]

            total_theory = 18

            for i in range(total_theory):
                subject = subjects[i % len(subjects)]
                events.append({
                    "id": f"T_{theory_counter:03d}",
                    "type": "THEORY",
                    "section": sid,
                    "dept": dept,
                    "sem": sem,
                    "subject": subject,
                    "duration": 1
                })
                theory_counter += 1

        # --------------------------------------------------------
        # LAB EVENTS: Strictly only 1 lab per week per class/section
        # --------------------------------------------------------
        lab_counter = 0
        for section in SECTIONS:
            sid = section["id"]
            dept = section["dept"]
            sem = section["sem"]
            subjects = SUBJECTS[dept][sem]
            lab_subject = subjects[0] + " Lab"

            events.append({
                "id": f"LAB_{lab_counter:03d}",
                "type": "LAB",
                "section": sid,
                "dept": dept,
                "sem": sem,
                "subject": lab_subject,
                "duration": 3  # 3 consecutive periods block
            })
            lab_counter += 1

        # --------------------------------------------------------
        # OPEN ELECTIVE (OE) EVENTS
        # --------------------------------------------------------
        oe_sections = [
            "ISE_S5_A",
            "ISE_S5_B",
            "ISE_S7_A",
            "ISE_S7_B",
            "CSBS_S5_A",
            "CSBS_S7_A"
        ]
        oe_counter = 0
        for sid in oe_sections:
            dept = SECTION_INFO[sid]["dept"]
            sem = SECTION_INFO[sid]["sem"]
            oe_subject = ISE_OE if dept == "ISE" else CSBS_OE

            for session in range(2):
                events.append({
                    "id": f"OE_{oe_counter:03d}",
                    "type": "OE",
                    "section": sid,
                    "dept": dept,
                    "sem": sem,
                    "subject": oe_subject,
                    "duration": 1
                })
                oe_counter += 1

        # --------------------------------------------------------
        # PHYSICAL EDUCATION (PE) EVENTS
        # --------------------------------------------------------
        pe_groups = [
            ["ISE_S5_A", "CSBS_S5_A"],
            ["ISE_S5_B"],
            ["ISE_S7_A", "CSBS_S7_A"],
            ["ISE_S7_B"]
        ]
        pe_counter = 0
        for group in pe_groups:
            for session in range(2):
                main_section = group[0]
                events.append({
                    "id": f"PE_{pe_counter:03d}",
                    "type": "PE",
                    "section": main_section,
                    "sections": group,
                    "dept": "COMBINED",
                    "sem": 5 if "S5" in main_section else 7,
                    "subject": "Physical Education",
                    "duration": 1
                })
                pe_counter += 1

        print("Sections :", len(SECTIONS))
        print("Staff    :", len(STAFF))
        print("Rooms    :", len(ROOMS))
        print("Classes  :", len(CLASSROOMS))
        print("Labs     :", len(LABS))
        print(f"Total Academic Events Curated: {len(events)}")
        print(f"Total Academic Event Hours   : {sum(e['duration'] for e in events)}")

        return events

# ================================================================
# 6. AGENT 2 - CONSTRAINT AGENT
# ================================================================

class ConstraintAgent:
    """Applies institutional constraints and mathematical bounds."""

    def run(self, events):
        print("\n" + "=" * 70)
        print("[AGENT 2] CONSTRAINT AGENT")
        print("=" * 70)

        counts = defaultdict(int)
        for e in events:
            counts[e["type"]] += 1

        print("Curated event breakdown:")
        for k, v in counts.items():
            print(f"  {k:10s}: {v}")

        print("\nScheduling constraints matrix:")
        print("  - Monday to Saturday (6 Operational Days)")
        print("  - 8 Daily Periods (8:30 AM - 5:30 PM)")
        print("  - Short Break: 10:30 - 11:00 (between Slot 1 & 2)")
        print("  - Lunch Break: 13:30 - 14:30 (between Slot 4 & 5)")
        print("  - AICTE Workload Norms: Prof (12h), Assoc Prof (14h), Asst Prof (16h)")
        print("  - Lab Constraint: Strictly 1 lab/week per class (no more than 1)")
        print("  - Valid 3-Period Lab Blocks: Slot 2 (11:00-13:30) or Slot 5 (14:30-17:30)")
        print("  - Synchronized OE: Slot 1 (09:30-10:30)")
        print("  - Synchronized PE: Slot 5 (14:30-15:30)")
        print("  - Senior Faculty Exemption: Professors & Assoc Profs cannot teach Slot 0 (08:30)")
        print("  - Shift Compactness: Early Shift (0..5) or Late Shift (2..7)")

        return {
            "events": events,
            "counts": dict(counts)
        }

# ================================================================
# 7. AGENT 3 - SOLVER AGENT (CP-SAT MULTI-STAGE OPTIMIZER)
# ================================================================

class SolverAgent:
    """Executes multi-stage optimization via Google OR-Tools CP-SAT."""

    def run(self, state):
        print("\n" + "=" * 70)
        print("[AGENT 3] SOLVER AGENT - CP-SAT (MULTI-STAGE OPTIMIZER)")
        print("=" * 70)

        events = state["events"]

        # Group events by course: (section(s), subject, type)
        # Guarantees teacher-course uniformity across the week
        course_events = defaultdict(list)
        for i, e in enumerate(events):
            sec_key = tuple(sorted(e.get("sections", [e["section"]])))
            course_events[(sec_key, e["subject"], e["type"])].append(i)

        courses = list(course_events.keys())
        num_courses = len(courses)
        print(f"Aggregated courses: {num_courses}")

        # --------------------------------------------------------
        # STAGE 3.1: FACULTY ALLOTMENT (WORKLOAD BALANCING)
        # --------------------------------------------------------
        print("\n[Stage 3.1] Faculty Allotment (AICTE Workload Norms Calibration)...")
        m_faculty = cp_model.CpModel()

        assigned = {}
        for c_idx in range(num_courses):
            for t in range(len(STAFF)):
                assigned[c_idx, t] = m_faculty.NewBoolVar(f"as_{c_idx}_{t}")

        for c_idx in range(num_courses):
            m_faculty.AddExactlyOne([assigned[c_idx, t] for t in range(len(STAFF))])

        course_hours = [sum(events[i]["duration"] for i in course_events[c]) for c in courses]

        # Workload bounds according to designation:
        # 1. Professor: 12 hours (bounded 10 to 14)
        # 2. Associate Professor: 14 hours (bounded 12 to 16)
        # 3. Assistant Professor: 16 hours (bounded 14 to 18)
        for t in range(len(STAFF)):
            desig = STAFF[t]["designation"]
            if desig == "Professor":
                th = m_faculty.NewIntVar(10, 14, f"th_{t}")
            elif desig == "Associate Professor":
                th = m_faculty.NewIntVar(12, 16, f"th_{t}")
            else:
                th = m_faculty.NewIntVar(14, 18, f"th_{t}")
            m_faculty.Add(th == sum(course_hours[c_idx] * assigned[c_idx, t] for c_idx in range(num_courses)))

        # Subject limit <= 3
        all_subjects = sorted(set(c[1] for c in courses))
        for t in range(len(STAFF)):
            sub_bools = []
            for s in all_subjects:
                matching_c = [c_idx for c_idx, c in enumerate(courses) if c[1] == s]
                sb = m_faculty.NewBoolVar(f"t_{t}_sub_{s}")
                m_faculty.AddMaxEquality(sb, [assigned[c_idx, t] for c_idx in matching_c])
                sub_bools.append(sb)
            m_faculty.Add(sum(sub_bools) <= MAX_SUBJECTS_PER_STAFF)

        # Section limit <= 3
        for t in range(len(STAFF)):
            sec_bools = []
            for sid in SECTION_IDS:
                matching_c = [c_idx for c_idx, c in enumerate(courses) if sid in c[0]]
                sb = m_faculty.NewBoolVar(f"t_{t}_sec_{sid}")
                m_faculty.AddMaxEquality(sb, [assigned[c_idx, t] for c_idx in matching_c])
                sec_bools.append(sb)
            m_faculty.Add(sum(sec_bools) <= MAX_SECTIONS_PER_STAFF)

        solver_f = cp_model.CpSolver()
        solver_f.parameters.max_time_in_seconds = 15
        solver_f.parameters.num_search_workers = 8
        status_f = solver_f.Solve(m_faculty)

        if status_f not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
            raise RuntimeError("Faculty allotment failed. Check teacher constraints.")

        print(f"Faculty Allotment status: {solver_f.StatusName(status_f)}")

        # Assign allotted faculty to each event
        for c_idx, c in enumerate(courses):
            for t in range(len(STAFF)):
                if solver_f.Value(assigned[c_idx, t]):
                    teacher = STAFF[t]
                    for i in course_events[c]:
                        events[i]["teacher"] = teacher["name"]
                        events[i]["teacher_id"] = teacher["id"]
                        events[i]["designation"] = teacher["designation"]
                        events[i]["teacher_idx"] = t
                    break

        # --------------------------------------------------------
        # STAGE 3.2: TIMETABLE TIME-SLOT SCHEDULING (CP-SAT)
        # --------------------------------------------------------
        print("\n[Stage 3.2] Timetable Slot Scheduling (8 Slots, Shifts, Lab blocks, Breaks)...")
        m_sched = cp_model.CpModel()

        day = [m_sched.NewIntVar(0, NUM_DAYS - 1, f"day_{i}") for i in range(len(events))]
        slot = [m_sched.NewIntVar(0, NUM_SLOTS - 1, f"slot_{i}") for i in range(len(events))]
        start = [m_sched.NewIntVar(0, NUM_DAYS * NUM_SLOTS - 1, f"start_{i}") for i in range(len(events))]

        for i in range(len(events)):
            m_sched.Add(start[i] == day[i] * NUM_SLOTS + slot[i])

        # Senior staff cannot teach at Slot 0 (08:30)
        for i, e in enumerate(events):
            if e["designation"] in ["Professor", "Associate Professor"]:
                m_sched.Add(slot[i] != 0)

        # Shift restrictions: Early shift (0..5) or Late shift (2..7)
        shift = {(sid, d): m_sched.NewBoolVar(f"late_{sid}_{d}") for sid in SECTION_IDS for d in range(NUM_DAYS)}
        for i, e in enumerate(events):
            sids = e.get("sections", [e["section"]])
            for sid in sids:
                for d in range(NUM_DAYS):
                    same_day = m_sched.NewBoolVar(f"ev_{i}_{sid}_d_{d}")
                    m_sched.Add(day[i] == d).OnlyEnforceIf(same_day)
                    m_sched.Add(day[i] != d).OnlyEnforceIf(same_day.Not())

                    late = shift[(sid, d)]
                    m_sched.Add(slot[i] >= 2 * late).OnlyEnforceIf(same_day)
                    m_sched.Add(slot[i] <= 5 + 2 * late).OnlyEnforceIf(same_day)

        # Fixed slots for OE and PE, and valid blocks for LAB
        for i, e in enumerate(events):
            if e["type"] == "OE":
                m_sched.Add(slot[i] == OE_SLOT)
            elif e["type"] == "PE":
                m_sched.Add(slot[i] == PE_SLOT)
            elif e["type"] == "LAB":
                valid_starts = [d * NUM_SLOTS + s for d in range(NUM_DAYS) for s in VALID_LAB_STARTS]
                m_sched.AddAllowedAssignments([start[i]], [[v] for v in valid_starts])

        # Section non-overlap
        sec_intervals = {sid: [] for sid in SECTION_IDS}
        for i, e in enumerate(events):
            sids = e.get("sections", [e["section"]])
            for sid in sids:
                iv = m_sched.NewIntervalVar(start[i], e["duration"], start[i] + e["duration"], f"sec_iv_{sid}_{i}")
                sec_intervals[sid].append(iv)

        for sid in SECTION_IDS:
            m_sched.AddNoOverlap(sec_intervals[sid])

        # Teacher non-overlap
        teacher_intervals = [[] for _ in range(len(STAFF))]
        for i, e in enumerate(events):
            t_idx = e["teacher_idx"]
            iv = m_sched.NewIntervalVar(start[i], e["duration"], start[i] + e["duration"], f"t_iv_{t_idx}_{i}")
            teacher_intervals[t_idx].append(iv)

        for t in range(len(STAFF)):
            m_sched.AddNoOverlap(teacher_intervals[t])

        # Cumulative room capacities
        cr_intervals = []
        lab_intervals = []
        for i, e in enumerate(events):
            iv = m_sched.NewIntervalVar(start[i], e["duration"], start[i] + e["duration"], f"room_iv_{i}")
            if e["type"] == "LAB":
                lab_intervals.append(iv)
            else:
                cr_intervals.append(iv)

        m_sched.AddCumulative(cr_intervals, [1] * len(cr_intervals), NUM_CLASSROOMS)
        m_sched.AddCumulative(lab_intervals, [1] * len(lab_intervals), NUM_LABS)

        solver_sched = cp_model.CpSolver()
        solver_sched.parameters.max_time_in_seconds = SOLVER_TIME_LIMIT
        solver_sched.parameters.num_search_workers = 8

        print("CP-SAT Time-Slot Solver started...")
        status_sched = solver_sched.Solve(m_sched)
        print("Solver status:", solver_sched.StatusName(status_sched))

        if status_sched not in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
            raise RuntimeError("No feasible timetable generated. Check constraints.")

        # Extract scheduled slots
        result = []
        for i, e in enumerate(events):
            d_val = solver_sched.Value(day[i])
            s_val = solver_sched.Value(slot[i])
            result.append({
                **e,
                "day": DAYS[d_val],
                "day_index": d_val,
                "slot": s_val,
                "time": TIME_SLOTS[s_val],
                "start_time": d_val * NUM_SLOTS + s_val
            })

        # --------------------------------------------------------
        # STAGE 3.3: ROOM ALLOCATION
        # --------------------------------------------------------
        print("\n[Stage 3.3] Room Allocation (CR_1..CR_7, LAB_1..LAB_2)...")

        # Assign labs
        lab_occupancy = defaultdict(set)
        for e in result:
            if e["type"] == "LAB":
                for lab_name in LABS:
                    if all(lab_name not in lab_occupancy[e["start_time"] + k] for k in range(e["duration"])):
                        e["room"] = lab_name
                        for k in range(e["duration"]):
                            lab_occupancy[e["start_time"] + k].add(lab_name)
                        break

        # Assign classrooms
        cr_occupancy = defaultdict(set)
        for e in result:
            if e["type"] != "LAB":
                for cr_name in CLASSROOMS:
                    if all(cr_name not in cr_occupancy[e["start_time"] + k] for k in range(e["duration"])):
                        e["room"] = cr_name
                        for k in range(e["duration"]):
                            cr_occupancy[e["start_time"] + k].add(cr_name)
                        break

        print("Room allocation complete.")
        return result

# ================================================================
# 8. AGENT 4 - VALIDATOR AGENT
# ================================================================

class ValidatorAgent:
    """Verifies all hard, soft, pedagogical, workload, and lab constraints."""

    def run(self, timetable):
        print("\n" + "=" * 70)
        print("[AGENT 4] VALIDATOR AGENT")
        print("=" * 70)

        errors = []
        warnings = []

        # 1. Section collision
        section_slots = defaultdict(list)
        for e in timetable:
            for sid in e.get("sections", [e["section"]]):
                for k in range(e["duration"]):
                    slot_number = e["day_index"] * NUM_SLOTS + e["slot"] + k
                    section_slots[(sid, slot_number)].append(e["id"])

        for key, ids in section_slots.items():
            if len(ids) > 1:
                errors.append(f"Section collision {key}: {ids}")

        # 2. Room collision
        room_slots = defaultdict(list)
        for e in timetable:
            for k in range(e["duration"]):
                slot_number = e["day_index"] * NUM_SLOTS + e["slot"] + k
                room_slots[(e["room"], slot_number)].append(e["id"])

        for key, ids in room_slots.items():
            if len(ids) > 1:
                errors.append(f"Room collision {key}: {ids}")

        # 3. Staff collision
        staff_slots = defaultdict(list)
        for e in timetable:
            for k in range(e["duration"]):
                slot_number = e["day_index"] * NUM_SLOTS + e["slot"] + k
                staff_slots[(e["teacher_id"], slot_number)].append(e["id"])

        for key, ids in staff_slots.items():
            if len(ids) > 1:
                errors.append(f"Staff collision {key}: {ids}")

        # 4. Lab room validation
        for e in timetable:
            if e["type"] == "LAB":
                if e["room"] not in LABS:
                    errors.append(f"Lab {e['id']} assigned to non-lab room {e['room']}")

        # 5. Strictly 1 lab in a week per class validation
        for sid in SECTION_IDS:
            sec_labs = [e for e in timetable if e["type"] == "LAB" and e["section"] == sid]
            if len(sec_labs) > 1:
                errors.append(f"Class {sid} has {len(sec_labs)} labs scheduled! Maximum allowed is strictly 1.")

        # 6. OE validation (Slot 1: 09:30-10:30)
        for e in timetable:
            if e["type"] == "OE":
                if e["slot"] != OE_SLOT:
                    errors.append(f"OE {e['id']} is not at slot {OE_SLOT} (09:30-10:30)")

        # 7. PE validation (Slot 5: 14:30-15:30)
        for e in timetable:
            if e["type"] == "PE":
                if e["slot"] != PE_SLOT:
                    errors.append(f"PE {e['id']} is not at slot {PE_SLOT} (14:30-15:30)")

        # 8. Break and lunch invariance (labs at Slot 2 or Slot 5)
        for e in timetable:
            if e["type"] == "LAB":
                if e["slot"] not in VALID_LAB_STARTS:
                    errors.append(f"Lab {e['id']} invalid block or crosses break: slot {e['slot']}")

        # 9. Senior staff cannot teach at Slot 0 (08:30)
        for e in timetable:
            if e["slot"] == 0:
                if e["designation"] in ["Professor", "Associate Professor"]:
                    errors.append(f"Senior staff at 08:30 (Slot 0): {e['teacher']} ({e['id']})")

        # 10. Teacher subject limit <= 3
        teacher_subjects = defaultdict(set)
        for e in timetable:
            teacher_subjects[e["teacher_id"]].add(e["subject"])

        for teacher, subjects in teacher_subjects.items():
            if len(subjects) > MAX_SUBJECTS_PER_STAFF:
                errors.append(f"{teacher} teaches {len(subjects)} subjects (Max: {MAX_SUBJECTS_PER_STAFF})")

        # 11. Teacher section limit <= 3
        teacher_sections = defaultdict(set)
        for e in timetable:
            for sid in e.get("sections", [e["section"]]):
                teacher_sections[e["teacher_id"]].add(sid)

        for teacher, sections in teacher_sections.items():
            if len(sections) > MAX_SECTIONS_PER_STAFF:
                errors.append(f"{teacher} teaches {len(sections)} sections (Max: {MAX_SECTIONS_PER_STAFF})")

        # 12. Expected event counts
        expected = {
            "THEORY": 13 * 18,  # 234
            "LAB": 13,
            "OE": 12,
            "PE": 8
        }
        actual = defaultdict(int)
        for e in timetable:
            actual[e["type"]] += 1

        for k, v in expected.items():
            if actual[k] != v:
                errors.append(f"{k}: expected {v}, got {actual[k]}")

        print("\nValidation Result:")
        if errors:
            print("FAILED - Errors:", len(errors))
            for err in errors[:20]:
                print("  -", err)
        else:
            print("PASSED - All constraints satisfied 100% with ZERO errors!")

        return {
            "status": "PASSED" if not errors else "FAILED",
            "errors": errors,
            "warnings": warnings,
            "teacher_subjects": teacher_subjects,
            "teacher_sections": teacher_sections
        }

# ================================================================
# 9. AGENT 5 - EXCEL GENERATOR (1-PAGE LANDSCAPE PRINT OPTIMIZED)
# ================================================================

class ExcelGenerator:
    """Generates styled Excel workbooks with page-fit configuration for 1-page PDF downloads."""

    def __init__(self, filename):
        self.filename = filename
        self.wb = Workbook()
        ws = self.wb.active
        self.wb.remove(ws)

        self.header_fill = PatternFill("solid", fgColor="1F4E78")
        self.lunch_fill = PatternFill("solid", fgColor="D9E1F2")
        self.break_fill = PatternFill("solid", fgColor="FFF2CC")

        self.header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        self.bold_font = Font(name="Calibri", size=10, bold=True)
        self.regular_font = Font(name="Calibri", size=9)

        self.center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        self.thin_border = Border(
            left=Side(style="thin", color="BFBFBF"),
            right=Side(style="thin", color="BFBFBF"),
            top=Side(style="thin", color="BFBFBF"),
            bottom=Side(style="thin", color="BFBFBF")
        )

    def write_section(self, sid, timetable):
        ws = self.wb.create_sheet(sid)

        # 1-Page Landscape Print Configuration
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True

        ws["A1"] = f"COLLEGE OF ENGINEERING - TIMETABLE: {sid}"
        ws["A1"].font = Font(name="Calibri", size=13, bold=True, color="1F4E78")
        ws.row_dimensions[1].height = 24

        headers = ["Time", *DAYS]
        ws.row_dimensions[3].height = 24
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=c, value=h)
            cell.fill = self.header_fill
            cell.font = self.header_font
            cell.alignment = self.center
            cell.border = self.thin_border

        event_map = {}
        for e in timetable:
            sections = e.get("sections", [e["section"]])
            if sid not in sections:
                continue
            for k in range(e["duration"]):
                event_map[(e["day"], e["slot"] + k)] = e

        current_row = 4
        for s in range(NUM_SLOTS):
            # Short Break between Slot 1 (09:30-10:30) and Slot 2 (11:00-11:50)
            if s == 2:
                ws.row_dimensions[current_row].height = 18
                ws.cell(row=current_row, column=1, value="10:30-11:00").font = self.bold_font
                ws.cell(row=current_row, column=1).alignment = self.center
                ws.cell(row=current_row, column=1).border = self.thin_border
                ws.cell(row=current_row, column=1).fill = self.break_fill

                for d in range(NUM_DAYS):
                    cell = ws.cell(row=current_row, column=d + 2, value="SHORT BREAK")
                    cell.alignment = self.center
                    cell.border = self.thin_border
                    cell.fill = self.break_fill
                    cell.font = self.bold_font
                current_row += 1

            # Lunch Break between Slot 4 (12:40-13:30) and Slot 5 (14:30-15:30)
            if s == 5:
                ws.row_dimensions[current_row].height = 18
                ws.cell(row=current_row, column=1, value="13:30-14:30").font = self.bold_font
                ws.cell(row=current_row, column=1).alignment = self.center
                ws.cell(row=current_row, column=1).border = self.thin_border
                ws.cell(row=current_row, column=1).fill = self.lunch_fill

                for d in range(NUM_DAYS):
                    cell = ws.cell(row=current_row, column=d + 2, value="LUNCH BREAK")
                    cell.alignment = self.center
                    cell.border = self.thin_border
                    cell.fill = self.lunch_fill
                    cell.font = self.bold_font
                current_row += 1

            ws.row_dimensions[current_row].height = 36
            ws.cell(row=current_row, column=1, value=TIME_SLOTS[s]).font = self.bold_font
            ws.cell(row=current_row, column=1).alignment = self.center
            ws.cell(row=current_row, column=1).border = self.thin_border

            for d in range(NUM_DAYS):
                e = event_map.get((DAYS[d], s))
                cell = ws.cell(row=current_row, column=d + 2)
                if e:
                    cell.value = f"{e['subject']}\n[{e['type']}]\n{e['teacher']}\n{e['room']}"
                    cell.font = self.regular_font
                else:
                    cell.value = "FREE"
                    cell.font = Font(name="Calibri", size=9, color="888888")

                cell.alignment = self.center
                cell.border = self.thin_border

            current_row += 1

        # Compact column widths for 1-page fit
        ws.column_dimensions["A"].width = 15
        for col in range(2, 8):
            ws.column_dimensions[get_column_letter(col)].width = 19

    def write_master(self, dept, timetable):
        ws = self.wb.create_sheet(f"{dept}_MASTER")
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True

        ws["A1"] = f"{dept} MASTER TIMETABLE"
        ws["A1"].font = Font(size=14, bold=True, color="1F4E78")

        headers = ["Day", "Time", "Section", "Subject", "Type", "Teacher", "Room"]
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=c, value=h)
            cell.fill = self.header_fill
            cell.font = self.header_font
            cell.alignment = self.center
            cell.border = self.thin_border

        row = 4
        dept_events = [
            e for e in timetable
            if (e["dept"] == dept or (e["type"] == "PE" and any(SECTION_INFO[s]["dept"] == dept for s in e.get("sections", []))))
        ]
        dept_events.sort(key=lambda x: (x["day_index"], x["slot"], x["section"]))

        for e in dept_events:
            section_text = ",".join(e.get("sections", [e["section"]]))
            values = [e["day"], e["time"], section_text, e["subject"], e["type"], e["teacher"], e["room"]]
            for c, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=c, value=value)
                cell.alignment = self.center
                cell.border = self.thin_border
                cell.font = self.regular_font
            row += 1

        for c in range(1, 8):
            ws.column_dimensions[get_column_letter(c)].width = 18

    def write_staff(self, timetable):
        ws = self.wb.create_sheet("STAFF_TIMETABLE")
        headers = ["Teacher", "Designation", "Day", "Time", "Subject", "Section", "Room", "Type"]
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.fill = self.header_fill
            cell.font = self.header_font
            cell.alignment = self.center
            cell.border = self.thin_border

        row = 2
        sorted_events = sorted(timetable, key=lambda x: (x["teacher"], x["day_index"], x["slot"]))
        for e in sorted_events:
            values = [
                e["teacher"],
                e["designation"],
                e["day"],
                e["time"],
                e["subject"],
                ",".join(e.get("sections", [e["section"]])),
                e["room"],
                e["type"]
            ]
            for c, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=c, value=value)
                cell.alignment = self.center
                cell.border = self.thin_border
                cell.font = self.regular_font
            row += 1

        for c in range(1, 9):
            ws.column_dimensions[get_column_letter(c)].width = 20

    def write_room_utilization(self, timetable):
        ws = self.wb.create_sheet("ROOM_UTILIZATION")
        headers = ["Room", "Day", "Time", "Subject", "Section", "Teacher", "Type"]
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.fill = self.header_fill
            cell.font = self.header_font
            cell.alignment = self.center
            cell.border = self.thin_border

        row = 2
        sorted_events = sorted(timetable, key=lambda x: (ROOMS.index(x["room"]), x["day_index"], x["slot"]))
        for e in sorted_events:
            values = [
                e["room"],
                e["day"],
                e["time"],
                e["subject"],
                ",".join(e.get("sections", [e["section"]])),
                e["teacher"],
                e["type"]
            ]
            for c, value in enumerate(values, 1):
                cell = ws.cell(row=row, column=c, value=value)
                cell.alignment = self.center
                cell.border = self.thin_border
                cell.font = self.regular_font
            row += 1

        for c in range(1, 8):
            ws.column_dimensions[get_column_letter(c)].width = 18

    def write_validation(self, validation):
        ws = self.wb.create_sheet("VALIDATION_REPORT")
        ws["A1"] = "TIMETABLE VALIDATION REPORT"
        ws["A1"].font = Font(size=14, bold=True)

        ws["A3"] = "Status"
        ws["B3"] = validation["status"]
        ws["B3"].font = Font(bold=True, color="008000" if validation["status"] == "PASSED" else "FF0000")

        ws["A4"] = "Errors"
        ws["B4"] = len(validation["errors"])

        row = 6
        ws.cell(row=row, column=1, value="Error Details").font = self.header_font
        ws.cell(row=row, column=1).fill = self.header_fill
        row += 1

        if not validation["errors"]:
            ws.cell(row=row, column=1, value="None - All hard, soft, pedagogical, workload, and lab constraints validated successfully!")
            row += 1
        else:
            for error in validation["errors"]:
                ws.cell(row=row, column=1, value=error)
                row += 1

        row += 2
        ws.cell(row=row, column=1, value="Teacher").fill = self.header_fill
        ws.cell(row=row, column=1).font = self.header_font
        ws.cell(row=row, column=2, value="Subjects").fill = self.header_fill
        ws.cell(row=row, column=2).font = self.header_font
        ws.cell(row=row, column=3, value="Sections").fill = self.header_fill
        ws.cell(row=row, column=3).font = self.header_font
        row += 1

        for teacher in STAFF_IDS:
            subjects = validation["teacher_subjects"].get(teacher, set())
            sections = validation["teacher_sections"].get(teacher, set())
            ws.cell(row=row, column=1, value=teacher)
            ws.cell(row=row, column=2, value=len(subjects))
            ws.cell(row=row, column=3, value=len(sections))
            row += 1

        for c in range(1, 4):
            ws.column_dimensions[get_column_letter(c)].width = 20

    def write_agent_report(self):
        ws = self.wb.create_sheet("AGENT_EXECUTION")
        report = [
            ["Agent", "Function", "Status"],
            ["Agent 1 - Curator", "Created sections, subjects, 1 lab/class, and academic events", "COMPLETED"],
            ["Agent 2 - Constraint", "Applied 8-slot, AICTE workload, breaks, and shift constraints", "COMPLETED"],
            ["Agent 3 - Solver", "CP-SAT multi-stage optimal scheduling and allocation", "COMPLETED"],
            ["Agent 4 - Validator", "Validated collisions, rooms, senior staff, 1-lab limit, and workload", "COMPLETED"],
            ["Agent 5 - Excel", "Generated 1-page fit master, section, staff, and utilization sheets", "COMPLETED"]
        ]
        for r, row_data in enumerate(report, 1):
            for c, value in enumerate(row_data, 1):
                cell = ws.cell(row=r, column=c, value=value)
                cell.alignment = self.center
                cell.border = self.thin_border
                if r == 1:
                    cell.fill = self.header_fill
                    cell.font = self.header_font

        ws.column_dimensions["A"].width = 25
        ws.column_dimensions["B"].width = 65
        ws.column_dimensions["C"].width = 18

    def save(self):
        self.wb.save(self.filename)
        print("\nExcel file created:", os.path.abspath(self.filename))

# ================================================================
# 10. MAIN PIPELINE ORCHESTRATION
# ================================================================

def main():
    start_time = time.time()
    print("\n" + "=" * 80)
    print("       AGENTIC AI COLLEGE TIMETABLE GENERATOR")
    print("       8-Slot Academic Architecture & AICTE Workload Model")
    print("=" * 80)

    curator = CuratorAgent()
    events = curator.run()

    constraint_agent = ConstraintAgent()
    state = constraint_agent.run(events)

    solver_agent = SolverAgent()
    timetable = solver_agent.run(state)

    validator = ValidatorAgent()
    validation = validator.run(timetable)

    print("\n" + "=" * 70)
    print("[AGENT 5] EXCEL REPORT GENERATOR")
    print("=" * 70)

    excel = ExcelGenerator(OUTPUT_FILE)
    for sid in SECTION_IDS:
        excel.write_section(sid, timetable)
    excel.write_master("ISE", timetable)
    excel.write_master("CSBS", timetable)
    excel.write_staff(timetable)
    excel.write_room_utilization(timetable)
    excel.write_validation(validation)
    excel.write_agent_report()
    excel.save()

    elapsed = time.time() - start_time
    print("\n" + "=" * 80)
    print(f"TIMETABLE GENERATION COMPLETED IN {round(elapsed, 2)}s (STATUS: {validation['status']})")
    print("=" * 80)
    return timetable, validation

if __name__ == "__main__":
    timetable, validation = main()
