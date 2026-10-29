"""
Torres Trail Rehabilitation — Gantt Chart Generator
Royal Palm Beach Pines Natural Area (RPBPNA), Palm Beach County, FL

Generates: TOR_Gantt_Schedule.xlsx
Run:       python3 build_gantt.py

Dependencies: pip install openpyxl
"""

from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_DATE_DATETIME
import datetime

# ── Project configuration ────────────────────────────────────────────────────

PROJECT_START = datetime.date(2026, 11, 1)   # Adjust to actual project start
TRAIL         = "Torres Trail (TOR)"
SITE          = "Royal Palm Beach Pines Natural Area"
PREPARED_BY   = "Trail Asset Management Platform (TAMP)"
PREPARED_DATE = datetime.date.today().strftime("%B %d, %Y")

# ── Work packages ─────────────────────────────────────────────────────────────
# Fields: (id, phase, task, action_type, scope_count, start_offset_days, duration_days, team, notes)

TASKS = [
    # Phase 1 — Mobilization
    ("WP-01", "Phase 1 – Mobilization",   "Project kickoff & site orientation",  "Admin",      0, 0,  1, "PM",       "Confirm scope, distribute SOW to all teams"),
    ("WP-02", "Phase 1 – Mobilization",   "Materials procurement — paint",        "Admin",      47, 0,  3, "PM",       "Brown/green trail marker paint + primer"),
    ("WP-03", "Phase 1 – Mobilization",   "Materials procurement — posts",        "Admin",      23, 0,  3, "PM",       "12 replacement + 11 new install posts"),
    ("WP-04", "Phase 1 – Mobilization",   "Tool & equipment staging",             "Admin",      0, 2,  1, "Logistics", "Post-hole digger, level, concrete mix"),

    # Phase 2 — Post Installation (new posts first so paint team follows)
    ("WP-05", "Phase 2 – Installation",   "Install new marker posts (11 posts)",  "Install",    11, 4,  3, "Team A",   "TOR-061–071 gap locations per site manager"),
    ("WP-06", "Phase 2 – Installation",   "Replace deteriorated posts (12 posts)","Replace",    12, 4,  4, "Team B",   "Worst-condition markers per condition score ≤2"),
    ("WP-07", "Phase 2 – Installation",   "Concrete cure / inspection hold",      "Admin",      0, 8,  2, "PM",       "Allow 48 hrs cure before painting new posts"),

    # Phase 3 — Paint
    ("WP-08", "Phase 3 – Paint",          "Paint existing posts (47 posts)",      "Paint",      47, 10, 5, "Team C",   "2-coat application; include repaired posts"),
    ("WP-09", "Phase 3 – Paint",          "Paint newly installed posts (11+12)",  "Paint",      23, 10, 3, "Team C",   "After concrete cure; schedule with WP-08"),

    # Phase 4 — No Action / Monitor
    ("WP-10", "Phase 4 – Monitor",        "Tag & document No-Action marker",      "No Action",  1, 14, 1, "Team D",   "TOR-022 metal directional sign — monitor only"),

    # Phase 5 — Quality Control
    ("WP-11", "Phase 5 – QC & Closeout",  "Trail walk-through inspection",        "QC",         0, 16, 1, "PM",       "GPS verify all 71 markers; photo each post"),
    ("WP-12", "Phase 5 – QC & Closeout",  "Update TAMP condition scores",         "Admin",      0, 17, 1, "PM",       "Log post-rehab scores in TAMF_v4 app"),
    ("WP-13", "Phase 5 – QC & Closeout",  "Final sign-off & documentation",       "Admin",      0, 18, 1, "PM",       "Submit completion report to site manager"),
]

TOTAL_DAYS = max(t[5] + t[6] for t in TASKS) + 2   # project duration in days

# ── Color palette ─────────────────────────────────────────────────────────────

C = {
    "header_bg":   "2D6A4F",   # TAMP primary green
    "header_fg":   "FFFFFF",
    "phase_bg":    "D8EDDF",   # light green phase rows
    "phase_fg":    "1B4332",
    "bar_install": "1D6FA0",   # blue
    "bar_replace": "C1440E",   # orange-red
    "bar_paint":   "2D6A4F",   # green
    "bar_admin":   "6C757D",   # grey
    "bar_qc":      "7B2D8B",   # purple
    "bar_monitor": "5A6E00",   # olive
    "bar_default": "AAAAAA",
    "weekend":     "F0F0F0",   # light grey for weekend columns
    "today":       "FFD166",   # yellow today marker
    "row_alt":     "F7FAF8",
    "border":      "B7D5C3",
}

ACTION_COLORS = {
    "Install":   C["bar_install"],
    "Replace":   C["bar_replace"],
    "Paint":     C["bar_paint"],
    "Admin":     C["bar_admin"],
    "QC":        C["bar_qc"],
    "No Action": C["bar_monitor"],
}

def hex_fill(hex_color, fill_type="solid"):
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type=fill_type)

def thin_border(left=True, right=True, top=True, bottom=True):
    s = Side(style="thin", color="B7D5C3")
    n = Side(style=None)
    return Border(
        left=s if left else n,
        right=s if right else n,
        top=s if top else n,
        bottom=s if bottom else n,
    )

def bold_border_bottom():
    return Border(bottom=Side(style="medium", color="2D6A4F"))

# ── Build workbook ─────────────────────────────────────────────────────────────

wb = Workbook()

# ── Sheet 1: Gantt Chart ───────────────────────────────────────────────────────

ws = wb.active
ws.title = "Gantt Schedule"
ws.sheet_view.showGridLines = False
ws.freeze_panes = "G4"   # freeze task columns; scroll through date columns

# Column layout:
# A=WP ID, B=Phase, C=Task, D=Action Type, E=Count, F=Team, G=Start Date, H=End Date, I=Duration
# J onwards = daily Gantt columns

FIXED_COLS   = 9
GANTT_START  = FIXED_COLS + 1   # column index of first date column (J)

col_widths = [8, 22, 38, 12, 7, 10, 13, 13, 10]
col_labels = ["WP #", "Phase", "Task / Description", "Action", "Count", "Team",
              "Start Date", "End Date", "Days"]

# ── Row 1: Title ──────────────────────────────────────────────────────────────
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=FIXED_COLS + TOTAL_DAYS)
title_cell = ws.cell(row=1, column=1,
    value=f"Torres Trail Rehabilitation — Project Schedule  |  {SITE}")
title_cell.font      = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
title_cell.fill      = hex_fill(C["header_bg"])
title_cell.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 26

# ── Row 2: Metadata ───────────────────────────────────────────────────────────
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=FIXED_COLS + TOTAL_DAYS)
meta_cell = ws.cell(row=2, column=1,
    value=f"Prepared by: {PREPARED_BY}   |   Date: {PREPARED_DATE}   |   "
          f"Project Start: {PROJECT_START.strftime('%B %d, %Y')}   |   "
          f"Scope: Paint 47 · Replace 12 · Install 11 · No Action 1")
meta_cell.font      = Font(name="Calibri", italic=True, size=10, color="FFFFFF")
meta_cell.fill      = hex_fill("1B4332")
meta_cell.alignment = Alignment(horizontal="left", vertical="center",
                                indent=1)
ws.row_dimensions[2].height = 16

# ── Row 3: Column headers (fixed) ────────────────────────────────────────────
for ci, (label, width) in enumerate(zip(col_labels, col_widths), start=1):
    c = ws.cell(row=3, column=ci, value=label)
    c.font      = Font(name="Calibri", bold=True, size=10, color=C["header_fg"])
    c.fill      = hex_fill(C["header_bg"])
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border    = thin_border()
    ws.column_dimensions[get_column_letter(ci)].width = width
ws.row_dimensions[3].height = 30

# ── Row 3: Date headers (Gantt columns) ───────────────────────────────────────
for d in range(TOTAL_DAYS):
    col  = GANTT_START + d
    date = PROJECT_START + datetime.timedelta(days=d)
    is_weekend = date.weekday() >= 5
    is_today   = date == datetime.date.today()

    c = ws.cell(row=3, column=col,
                value=date.strftime("%m/%d"))
    c.font      = Font(name="Calibri", bold=True, size=8,
                       color=C["header_fg"] if not is_weekend else "888888")
    c.fill      = hex_fill(C["today"] if is_today else
                           (C["weekend"] if is_weekend else C["header_bg"]))
    c.alignment = Alignment(horizontal="center", vertical="center",
                            text_rotation=90)
    c.border    = thin_border()
    ws.column_dimensions[get_column_letter(col)].width = 3.2

# ── Rows 4+: Tasks ────────────────────────────────────────────────────────────
current_phase = None
row = 4

for task in TASKS:
    wp_id, phase, task_name, action, count, start_off, dur, team, notes = task

    # Phase header row when phase changes
    if phase != current_phase:
        current_phase = phase
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=FIXED_COLS + TOTAL_DAYS)
        ph = ws.cell(row=row, column=1, value=f"  {phase}")
        ph.font      = Font(name="Calibri", bold=True, size=10, color=C["phase_fg"])
        ph.fill      = hex_fill(C["phase_bg"])
        ph.alignment = Alignment(vertical="center")
        ph.border    = bold_border_bottom()
        ws.row_dimensions[row].height = 18
        row += 1

    # Task start / end dates
    start_date = PROJECT_START + datetime.timedelta(days=start_off)
    end_date   = start_date   + datetime.timedelta(days=dur - 1)

    is_alt = (row % 2 == 0)
    row_fill = hex_fill(C["row_alt"]) if is_alt else PatternFill()

    # Fixed columns
    fixed_vals = [wp_id, phase, task_name, action,
                  count if count > 0 else "—",
                  team, start_date, end_date, dur]
    for ci, val in enumerate(fixed_vals, start=1):
        c = ws.cell(row=row, column=ci, value=val)
        c.font      = Font(name="Calibri", size=10)
        c.fill      = row_fill
        c.border    = thin_border()
        c.alignment = Alignment(vertical="center",
                                horizontal="center" if ci in (1,5,6,9) else "left",
                                wrap_text=(ci == 3))
        if ci in (7, 8):   # date columns
            c.number_format = "MM/DD/YYYY"
            c.alignment     = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[row].height = 22

    # Gantt bar columns
    bar_color = ACTION_COLORS.get(action, C["bar_default"])
    for d in range(TOTAL_DAYS):
        col  = GANTT_START + d
        date = PROJECT_START + datetime.timedelta(days=d)
        is_weekend = date.weekday() >= 5
        is_today   = date == datetime.date.today()
        in_bar     = (start_date <= date <= end_date)

        c = ws.cell(row=row, column=col, value="")
        c.border = thin_border()

        if in_bar:
            c.fill = hex_fill(bar_color)
        elif is_today:
            c.fill = hex_fill(C["today"])
        elif is_weekend:
            c.fill = hex_fill(C["weekend"])
        elif is_alt:
            c.fill = hex_fill(C["row_alt"])

    row += 1

# ── Totals row ────────────────────────────────────────────────────────────────
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
tot = ws.cell(row=row, column=1, value="PROJECT TOTALS")
tot.font      = Font(name="Calibri", bold=True, size=10, color=C["header_fg"])
tot.fill      = hex_fill(C["header_bg"])
tot.alignment = Alignment(horizontal="right", vertical="center", indent=1)

totals = {"Paint": 47, "Replace": 12, "Install": 11, "No Action": 1, "Total Posts": 71}
for ci, (label, val) in enumerate(totals.items(), start=5):
    hc = ws.cell(row=row - 1 + 1, column=ci)   # reuse same row
    pass

# simpler: single summary cell
c5 = ws.cell(row=row, column=5,
             value="Paint 47 · Replace 12 · Install 11 · No Action 1 = 71 posts total")
ws.merge_cells(start_row=row, start_column=5, end_row=row,
               end_column=FIXED_COLS + TOTAL_DAYS)
c5.font      = Font(name="Calibri", bold=True, size=10, color=C["header_fg"])
c5.fill      = hex_fill(C["header_bg"])
c5.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[row].height = 20

# ── Sheet 2: Legend ────────────────────────────────────────────────────────────

wl = wb.create_sheet("Legend & Notes")
wl.sheet_view.showGridLines = False
wl.column_dimensions["A"].width = 4
wl.column_dimensions["B"].width = 18
wl.column_dimensions["C"].width = 50
wl.column_dimensions["D"].width = 20

# Title
wl.merge_cells("A1:D1")
t = wl.cell(row=1, column=1, value="Legend & Project Notes — Torres Trail Rehabilitation")
t.font      = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
t.fill      = hex_fill(C["header_bg"])
t.alignment = Alignment(horizontal="center", vertical="center")
wl.row_dimensions[1].height = 24

legend_items = [
    ("Install",   C["bar_install"], "New marker post installation (11 posts, TOR-061–071 gap locations)"),
    ("Replace",   C["bar_replace"], "Deteriorated post replacement (12 posts, condition score ≤ 2)"),
    ("Paint",     C["bar_paint"],   "Post repainting — 2-coat application (47 existing + 23 new/replaced)"),
    ("Admin",     C["bar_admin"],   "Project management, procurement, documentation"),
    ("QC",        C["bar_qc"],      "Quality control inspection and TAMP data entry"),
    ("No Action", C["bar_monitor"], "Monitor only — TOR-022 metal directional sign, no work required"),
]

wl.cell(row=3, column=2, value="Action Type").font = Font(bold=True, size=11)
wl.cell(row=3, column=3, value="Description").font  = Font(bold=True, size=11)

for i, (label, color, desc) in enumerate(legend_items, start=4):
    swatch = wl.cell(row=i, column=1, value=" ")
    swatch.fill = hex_fill(color)

    lc = wl.cell(row=i, column=2, value=label)
    lc.font = Font(name="Calibri", bold=True, size=10)

    dc = wl.cell(row=i, column=3, value=desc)
    dc.font = Font(name="Calibri", size=10)
    dc.alignment = Alignment(wrap_text=True)

    wl.row_dimensions[i].height = 18

# Notes section
notes_start = len(legend_items) + 6
wl.merge_cells(f"B{notes_start}:D{notes_start}")
nh = wl.cell(row=notes_start, column=2, value="Project Notes")
nh.font = Font(name="Calibri", bold=True, size=11, color=C["header_bg"])

notes = [
    "• Project start date is configurable — change PROJECT_START in build_gantt.py and re-run.",
    "• Weekend columns (Sat/Sun) are shaded grey; adjust task durations to account for non-working days.",
    "• Today's date column is highlighted in yellow for quick visual reference.",
    "• Scope source: Site manager SOW — Paint 47, Replace 12, Install 11 new, No Action 1 (TOR-022).",
    "• 11 new posts (WP-05) correspond to gap locations identified in field survey; assign TOR-061–071.",
    "• This schedule was generated by build_gantt.py using openpyxl — part of the TAMP project package.",
    "• For marker-level tracking, condition scores, and GPS data — use TAMF_v4.html (TAMP web app).",
]

for j, note in enumerate(notes, start=notes_start + 1):
    nc = wl.cell(row=j, column=2, value=note)
    nc.font      = Font(name="Calibri", size=10)
    nc.alignment = Alignment(wrap_text=True)
    wl.merge_cells(f"B{j}:D{j}")
    wl.row_dimensions[j].height = 16

# ── Sheet 3: Budget skeleton ────────────────────────────────────────────────────

wb_sheet = wb.create_sheet("Budget Sketch")
wb_sheet.sheet_view.showGridLines = False
wb_sheet.column_dimensions["A"].width = 6
wb_sheet.column_dimensions["B"].width = 30
wb_sheet.column_dimensions["C"].width = 12
wb_sheet.column_dimensions["D"].width = 12
wb_sheet.column_dimensions["E"].width = 14
wb_sheet.column_dimensions["F"].width = 22

wb_sheet.merge_cells("A1:F1")
bt = wb_sheet.cell(row=1, column=1,
                   value="Torres Trail Rehabilitation — Budget Sketch (fill in unit costs)")
bt.font      = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
bt.fill      = hex_fill(C["header_bg"])
bt.alignment = Alignment(horizontal="center", vertical="center")
wb_sheet.row_dimensions[1].height = 24

headers = ["#", "Item", "Qty", "Unit Cost ($)", "Total ($)", "Notes"]
for ci, h in enumerate(headers, start=1):
    hc = wb_sheet.cell(row=3, column=ci, value=h)
    hc.font      = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    hc.fill      = hex_fill(C["header_bg"])
    hc.alignment = Alignment(horizontal="center", vertical="center")
    hc.border    = thin_border()

budget_items = [
    (1,  "Trail marker posts (replacement)",         12, "", "=C4*D4",   "Pressure-treated 4x4, 42\" length"),
    (2,  "Trail marker posts (new install)",          11, "", "=C5*D5",   "Same spec as replacement posts"),
    (3,  "Concrete mix (bags)",                       23, "", "=C6*D6",   "1 bag per post hole"),
    (4,  "Trail marker paint — brown/green",           4, "", "=C7*D7",   "Quarts; covers ~47 posts 2-coat"),
    (5,  "Primer (posts before painting)",             2, "", "=C8*D8",   "For new/replaced posts only"),
    (6,  "Brushes / rollers / trays",                  1, "", "=C9*D9",   "Consumables kit"),
    (7,  "Post-hole digger rental",                    1, "", "=C10*D10", "1-day rental"),
    (8,  "Volunteer labor — installation (hrs)",      40, "", "=C11*D11", "In-kind; set unit cost = $0 or imputed"),
    (9,  "Volunteer labor — painting (hrs)",          60, "", "=C12*D12", "In-kind"),
    (10, "Volunteer labor — QC walk-through (hrs)",    8, "", "=C13*D13", "In-kind"),
    (11, "Miscellaneous / contingency (5%)",           1, "", "",          "=sum of items 1-10 × 5%"),
]

for i, (num, item, qty, unit, total, note) in enumerate(budget_items, start=4):
    row_data = [num, item, qty, unit, total, note]
    is_alt   = (i % 2 == 0)
    for ci, val in enumerate(row_data, start=1):
        c = wb_sheet.cell(row=i, column=ci, value=val)
        c.font      = Font(name="Calibri", size=10)
        c.fill      = hex_fill(C["row_alt"]) if is_alt else PatternFill()
        c.border    = thin_border()
        c.alignment = Alignment(horizontal="center" if ci in (1, 3) else "left",
                                vertical="center", wrap_text=(ci == 6))

# Total row
tr = len(budget_items) + 4
wb_sheet.merge_cells(f"A{tr}:D{tr}")
tc = wb_sheet.cell(row=tr, column=1, value="PROJECT TOTAL")
tc.font      = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
tc.fill      = hex_fill(C["header_bg"])
tc.alignment = Alignment(horizontal="right", vertical="center", indent=1)

total_formula = f"=SUM(E4:E{tr-1})"
te = wb_sheet.cell(row=tr, column=5, value=total_formula)
te.font      = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
te.fill      = hex_fill(C["header_bg"])
te.number_format = "$#,##0.00"
te.alignment = Alignment(horizontal="center", vertical="center")

wb_sheet.cell(row=tr, column=6, value="Fill unit costs in column D to auto-calculate").font = \
    Font(name="Calibri", italic=True, size=9, color="555555")

wb_sheet.row_dimensions[tr].height = 20
for r in range(4, tr):
    wb_sheet.row_dimensions[r].height = 18

# ── Save ──────────────────────────────────────────────────────────────────────

OUT_PATH = "/mnt/user-data/outputs/TOR_Gantt_Schedule.xlsx"
wb.save(OUT_PATH)
print(f"Saved: {OUT_PATH}")
print(f"Sheets: {[s.title for s in wb.worksheets]}")
print(f"Project start: {PROJECT_START}  |  Duration: {TOTAL_DAYS} calendar days")
