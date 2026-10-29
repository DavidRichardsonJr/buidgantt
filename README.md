build_gantt.py — Torres Trail Gantt Schedule Generator

Generates TOR_Gantt_Schedule.xlsx — a colour-coded project schedule for the Torres Trail (TOR) rehabilitation at Royal Palm Beach Pines Natural Area.

Requirements
pip install openpyxl

Python 3.8 or later.

Quick start
bash
python3 build_gantt.py

Outputs TOR_Gantt_Schedule.xlsx in the same directory.

What's inside the workbook
Sheet	Contents
Gantt Schedule	13 work packages across 5 phases with colour-coded bars, weekend shading, today marker
Legend & Notes	Colour key for each action type + usage notes
Budget Sketch	Line-item budget skeleton — fill unit costs in column D to auto-calculate totals
Customising for another trail or project

All configuration lives at the top of build_gantt.py:

python
PROJECT_START = datetime.date(2026, 11, 1)   # ← change start date
TRAIL         = "Torres Trail (TOR)"          # ← change trail/project name
SITE          = "Royal Palm Beach Pines Natural Area"
PREPARED_BY   = "Trail Asset Management Platform (TAMP)"
Changing the work packages

Edit the TASKS list. Each entry is a tuple:

python
(id, phase, task_label, action_type, count, start_offset_days, duration_days, team, notes)
Field	Description
id	Short unique string, e.g. "WP-01"
phase	Phase heading — rows with the same phase are grouped
task_label	Display name for the task
action_type	Controls bar colour: Install, Replace, Paint, Admin, QC, No Action
count	Scope quantity (posts, items) — shown in the Count column; 0 shows —
start_offset_days	Days after PROJECT_START when this task begins
duration_days	How many calendar days the task spans
team	Team or person responsible
notes	Appears only in the script; add a Notes column to the sheet if needed
Adding a new action type / bar colour

Add an entry to the ACTION_COLORS dict:

python
ACTION_COLORS = {
    "Install":   C["bar_install"],   # blue
    "Replace":   C["bar_replace"],   # orange-red
    "Paint":     C["bar_paint"],     # green
    "Admin":     C["bar_admin"],     # grey
    "QC":        C["bar_qc"],        # purple
    "No Action": C["bar_monitor"],   # olive
    "Survey":    "E6A817",           # ← add new type here (hex without #)
}
Re-using for a completely different project
Change PROJECT_START, TRAIL, SITE at the top.
Replace the TASKS list with your own work packages.
Run python3 build_gantt.py.

The script has no dependencies on Torres Trail data — TASKS is the entire project definition.

Relationship to TAMP web app

build_gantt.py is a standalone generator 

Output file

TOR_Gantt_Schedule.xlsx — safe to re-run; overwrites the previous file.
