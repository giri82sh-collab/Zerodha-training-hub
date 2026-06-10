#!/usr/bin/env python3
"""
Zerodha Training Hub — Full Excel Report Generator (Template-Based)
Loads master sheet template, populates dynamic data, and produces high-fidelity styled Excel report with charts.
Usage: python3 /Users/girisha/Desktop/Zerodha-training-hub/generate_zerodha_report.py
"""

import json, re, urllib.request, os, sys, copy
from datetime import datetime

# ── Dependencies ──────────────────────────────────────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, PieChart, Reference, Series
    from openpyxl.chart.series import DataPoint
    from openpyxl.chart.label import DataLabelList
except ImportError:
    os.system("pip3 install openpyxl -q")
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, PieChart, Reference, Series
    from openpyxl.chart.series import DataPoint
    from openpyxl.chart.label import DataLabelList

# ── Config ─────────────────────────────────────────────────────────────────────
def find_html_path():
    paths = [
        os.path.expanduser("~/Desktop/zerodha-training-hub.html"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "zerodha-training-hub.index.html"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "index.html"),
        "./zerodha-training-hub.html",
        "./index.html"
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return paths[0]

def find_template_path():
    paths = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "template.xlsx"),
        "/Users/girisha/Desktop/Zerodha Trainer Dashboard Master (1).xlsx",
        os.path.expanduser("~/Desktop/Zerodha Trainer Dashboard Master (1).xlsx")
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    raise FileNotFoundError("Could not find Zerodha Trainer Dashboard Master (1).xlsx template file.")

FIREBASE_URL = "https://zerodha-training-hub-default-rtdb.firebaseio.com"
HTML_PATH    = find_html_path()
TEMPLATE_PATH = find_template_path()
OUT_PATH     = os.path.expanduser("~/Desktop/Zerodha_Training_Dashboard.xlsx")
MONTHS_ORDER = [
    "December 2025", "January 2026", "February 2026", "March 2026", "April 2026",
    "May 2026", "June 2026", "July 2026", "August 2026", "September 2026",
    "October 2026", "November 2026", "December 2026"
]
TRAINERS     = ["Gaurav Kumar", "Raju Vernekar", "Freeda Reshma Dsouza", "Harish Bhat", "Girish A"]

# ── Colors (exact from template) ─────────────────────────────────────────────
TEAL         = "00BFA5"
DARK_TEAL    = "00A187"
DARK_BG      = "232840"
DARK_BG2     = "1B2033"
WHITE        = "FFFFFF"
GREEN_HDR    = "258229"
GREEN_VAL    = "43A047"
PURPLE_HDR   = "5D43E1"
PURPLE_VAL   = "7B61FF"
AMBER_HDR    = "D78805"
AMBER_VAL    = "F5A623"
BLUE_HDR     = "0047A2"
BLUE_VAL     = "1565C0"
LIGHT_TEXT   = "E8EBF0"
DIM_TEXT     = "546E7A"
DARK_TEXT    = "1F2937"
ALT_ROW      = "F5F7FA"

TRAINER_COLORS = {
    "Gaurav Kumar": TEAL,
    "Raju Vernekar": PURPLE_VAL,
    "Freeda Reshma Dsouza": AMBER_VAL,
    "Harish Bhat": BLUE_VAL,
    "Girish A": GREEN_VAL
}

def fill(hex_rgb):
    return PatternFill("solid", fgColor=hex_rgb)

def font(hex_rgb, bold=False, italic=False, size=10, name="Calibri"):
    return Font(color=hex_rgb, bold=bold, italic=italic, size=size, name=name)

def align(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def thin_border(color="CCCCCC"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def to_float(val, default=0.0):
    try:
        return float(val) if val is not None else default
    except (ValueError, TypeError):
        return default

def style_range(ws, cell_range, fill_=None, font_=None, align_=None, border_=None):
    if ":" in cell_range:
        cells = ws[cell_range]
    else:
        cells = [[ws[cell_range]]]
    for row in cells:
        for cell in row:
            if fill_ is not None:    cell.fill = fill_
            if font_ is not None:    cell.font = font_
            if align_ is not None:   cell.alignment = align_
            if border_ is not None:  cell.border = border_

# ── Fetch data ─────────────────────────────────────────────────────────────────
def fetch_firebase(key):
    try:
        url = f"{FIREBASE_URL}/{key}.json"
        with urllib.request.urlopen(url, timeout=10) as r:
            data = json.loads(r.read())
        if isinstance(data, dict):
            if "v" in data:
                data = data["v"]
            else:
                data = list(data.values())
        if isinstance(data, list) and len(data) == 1 and isinstance(data[0], list):
            data = data[0]
        return data or []
    except Exception as e:
        print(f"  Firebase fetch '{key}' failed: {e}")
        return []

def parse_tr_activities():
    """Extract hardcoded TR.activities from the HTML file."""
    try:
        with open(HTML_PATH, "r", encoding="utf-8") as f:
            html = f.read()
        m = re.search(r'"activities"\s*:\s*(\[.*?\])\s*,\s*"kra"', html, re.DOTALL)
        if m:
            return json.loads(m.group(1))
    except Exception as e:
        print(f"  Could not parse HTML: {e}")
    return []

def parse_tr_monthly():
    """Extract TR.monthly_hours from the HTML file."""
    try:
        with open(HTML_PATH, "r", encoding="utf-8") as f:
            html = f.read()
        m = re.search(r'"monthly_hours"\s*:\s*(\{.*?\})\s*,\s*"activities"', html, re.DOTALL)
        if m:
            return json.loads(m.group(1))
    except: pass
    return {}

print("Fetching data…")
base_acts  = parse_tr_activities()
logs       = fetch_firebase("zh_logs")
if isinstance(logs, list):
    all_logs = [l for l in logs if isinstance(l, dict)]
else:
    all_logs = []
all_acts   = base_acts + all_logs

# Filter to report months only
report_set = set(MONTHS_ORDER)
all_acts   = [a for a in all_acts if a.get("month") in report_set]
all_acts.sort(key=lambda a: MONTHS_ORDER.index(a.get("month", MONTHS_ORDER[-1]))
              if a.get("month") in MONTHS_ORDER else 99)

tr_mhrs    = parse_tr_monthly()

# Build monthly hours (TR base + log additions)
def get_monthly_hours():
    base = copy.deepcopy(tr_mhrs)
    for lg in all_logs:
        m = lg.get("month","")
        if m not in report_set: continue
        if m not in base: base[m] = {t:0 for t in TRAINERS}
        tr = lg.get("trainer","")
        hrs = to_float(lg.get("hours",0))
        for t in TRAINERS:
            if t.split()[0] in tr or t in tr:
                base[m][t] = base[m].get(t,0) + hrs
    return base

mhrs = get_monthly_hours()

# Collect all activity types
act_types = {}
for a in all_acts:
    t = a.get("type","Other")
    act_types[t] = act_types.get(t,0) + to_float(a.get("hours",0))

all_types_sorted = sorted(act_types.items(), key=lambda x: -x[1])

print(f"  {len(all_acts)} activities across {len(set(a['month'] for a in all_acts))} months")

# ── Create workbook from template ──────────────────────────────────────────────
print(f"Loading template from {TEMPLATE_PATH}…")
wb = openpyxl.load_workbook(TEMPLATE_PATH)

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1 — DATA
# ══════════════════════════════════════════════════════════════════════════════
print("Updating Data sheet…")
ws_data = wb["Data"]

# Clear old data rows
if ws_data.max_row >= 2:
    ws_data.delete_rows(2, ws_data.max_row - 1)

# Write dynamic activities
for i, act in enumerate(all_acts):
    row = i + 2
    is_online = act.get("session","") == "Online"
    bg_color = "E3F2FD" if is_online else "E8F5E9"
    vals = [
        act.get("month",""),
        act.get("type",""),
        act.get("details") or act.get("training_details",""),
        act.get("start") or act.get("date",""),
        act.get("end",""),
        act.get("trainer",""),
        act.get("session",""),
        act.get("topic",""),
        to_float(act.get("hours",0)),
    ]
    for j, v in enumerate(vals):
        cell = ws_data.cell(row=row, column=j+1, value=v)
        cell.fill  = fill(bg_color)
        cell.font  = font("111827", size=10)
        cell.alignment = align("left" if j < 8 else "right", "center")
        cell.border = thin_border("EEEEEE")

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2 — DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
print("Updating Dashboard sheet…")
ws_dash = wb["Dashboard"]

# Title row subtitle update
ws_dash["A3"].value = f"Training Activity Dashboard   |   {MONTHS_ORDER[0]} – {MONTHS_ORDER[-1]}"

# Filter cell default
ws_dash["D6"].value = "All Months"

# Row resizing calculations for active months
active_months = [m for m in MONTHS_ORDER if m in mhrs]
num_active = len(active_months)
template_months_count = 5  # Template has Dec 2025 - Apr 2026 (rows 19 to 23)
diff = num_active - template_months_count

# Shift cells, heights, and merged ranges below summary table if month count differs
if diff > 0:
    ws_dash.insert_rows(24, diff)
    
    # Adjust merged ranges below row 24
    merged_ranges = list(ws_dash.merged_cells.ranges)
    for r in merged_ranges:
        ws_dash.merged_cells.remove(r)
    for r in merged_ranges:
        if r.min_row >= 24:
            r.shift(row_shift=diff)
        ws_dash.merged_cells.add(r)
        
    # Copy layout heights
    for r_idx in range(24, 24 + diff):
        ws_dash.row_dimensions[r_idx].height = 19.5
elif diff < 0:
    ws_dash.delete_rows(24, -diff)
    
    merged_ranges = list(ws_dash.merged_cells.ranges)
    for r in merged_ranges:
        ws_dash.merged_cells.remove(r)
    for r in merged_ranges:
        if r.min_row >= 24:
            r.shift(row_shift=diff)
        ws_dash.merged_cells.add(r)

# Write summary rows B19 to O{18 + num_active}
for idx, m in enumerate(active_months):
    row = 19 + idx
    bg = WHITE if idx % 2 == 0 else ALT_ROW
    
    # Merge Month Name (B to C)
    ws_dash.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
    ws_dash.cell(row=row, column=2, value=m)
    style_range(ws_dash, f"B{row}:C{row}", fill_=fill(bg), font_=font("1F2937", bold=True, size=9), align_=align("left", "center"), border_=Border(bottom=Side(style="thin", color="E8EBF0")))
    
    # Trainer values
    tr_cols = [4, 6, 8, 10, 12]  # D, F, H, J, L
    for col_idx, trainer in zip(tr_cols, TRAINERS):
        ws_dash.merge_cells(start_row=row, start_column=col_idx, end_row=row, end_column=col_idx+1)
        c_letter = get_column_letter(col_idx)
        c_letter2 = get_column_letter(col_idx+1)
        cell = ws_dash.cell(row=row, column=col_idx)
        cell.value = f'=SUMIFS(Data!$I$2:$I$200,Data!$A$2:$A$200,"{m}",Data!$F$2:$F$200,"*{trainer.split()[0]}*")'
        style_range(ws_dash, f"{c_letter}{row}:{c_letter2}{row}", fill_=fill(bg), font_=font("1F2937", size=9), align_=align("right", "center"), border_=Border(bottom=Side(style="thin", color="E8EBF0")))
        cell.number_format = "0.00"
        ws_dash.cell(row=row, column=col_idx+1).number_format = "0.00"
        
    # Grand Total (N to O)
    ws_dash.merge_cells(start_row=row, start_column=14, end_row=row, end_column=15)
    cell_gt = ws_dash.cell(row=row, column=14)
    cell_gt.value = f"=SUM(D{row},F{row},H{row},J{row},L{row})"
    style_range(ws_dash, f"N{row}:O{row}", fill_=fill(bg), font_=font("7B61FF", bold=True, size=9), align_=align("right", "center"), border_=Border(bottom=Side(style="thin", color="E8EBF0")))
    cell_gt.number_format = "0.00"
    ws_dash.cell(row=row, column=15).number_format = "0.00"

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3 — CD (Chart Data helper)
# ══════════════════════════════════════════════════════════════════════════════
print("Updating CD helper sheet…")
ws_cd = wb["CD"]

# Clear everything in CD sheet
if ws_cd.max_row >= 1:
    ws_cd.delete_rows(1, ws_cd.max_row)

# 1. Write monthly hours per trainer (rows 1 to 1 + len(active_months))
ws_cd.cell(row=1, column=1, value="Month")
for j, t in enumerate(TRAINERS):
    ws_cd.cell(row=1, column=j+2, value=t)

for i, m in enumerate(active_months):
    row = i + 2
    ws_cd.cell(row=row, column=1).value = m
    for j, t in enumerate(TRAINERS):
        ws_cd.cell(row=row, column=j+2).value = (
            f'=IF(OR(Dashboard!$D$6="All Months",Dashboard!$D$6="{m}"),'
            f'SUMPRODUCT((Data!$A$2:$A$200="{m}")*ISNUMBER(SEARCH("{t.split()[0]}",Data!$F$2:$F$200))*Data!$I$2:$I$200),0)'
        )

# 2. Write session type breakdown (Online vs Offline)
row_session = 3 + len(active_months)
ws_cd.cell(row=row_session, column=1, value="Session Type")
ws_cd.cell(row=row_session, column=2, value="Hours")

ws_cd.cell(row=row_session+1, column=1, value="Online")
ws_cd.cell(row=row_session+1, column=2, value=(
    f'=IF(Dashboard!$D$6="All Months",'
    f'SUMIF(Data!$G$2:$G$200,"Online",Data!$I$2:$I$200),'
    f'SUMPRODUCT((Data!$A$2:$A$200=Dashboard!$D$6)*(Data!$G$2:$G$200="Online")*Data!$I$2:$I$200))'
))

ws_cd.cell(row=row_session+2, column=1, value="Offline")
ws_cd.cell(row=row_session+2, column=2, value=(
    f'=IF(Dashboard!$D$6="All Months",'
    f'SUMIF(Data!$G$2:$G$200,"Offline",Data!$I$2:$I$200),'
    f'SUMPRODUCT((Data!$A$2:$A$200=Dashboard!$D$6)*(Data!$G$2:$G$200="Offline")*Data!$I$2:$I$200))'
))

# 3. Write activity breakdown
row_act = row_session + 4
ws_cd.cell(row=row_act, column=1, value="Activity")
ws_cd.cell(row=row_act, column=2, value="Hours")

n_types = len(all_types_sorted)
for idx, (act, _) in enumerate(all_types_sorted):
    r = row_act + 1 + idx
    ws_cd.cell(row=r, column=1, value=act)
    ws_cd.cell(row=r, column=2, value=(
        f'=IF(Dashboard!$D$6="All Months",'
        f'SUMIF(Data!$B$2:$B$200,"{act}",Data!$I$2:$I$200),'
        f'SUMPRODUCT((Data!$A$2:$A$200=Dashboard!$D$6)*(Data!$B$2:$B$200="{act}")*Data!$I$2:$I$200))'
    ))

# ── RE-CREATE CHARTS ──────────────────────────────────────────────────────────
print("Re-creating charts…")
ws_dash._charts.clear()

chart_lbl_row = 19 + num_active + 2
chart_anchor_row = chart_lbl_row + 1

# Chart 1: Stacked Bar — Monthly Hours by Trainer
bar1 = BarChart()
bar1.type    = "bar"
bar1.grouping = "stacked"
bar1.overlap  = 100
bar1.width   = 15
bar1.height  = 12
bar1.title   = None
bar1.legend.position = "b"
bar1.y_axis.title = "Hours"

cats = Reference(ws_cd, min_col=1, max_col=1, min_row=2, max_row=1+num_active)
for j, trainer in enumerate(TRAINERS):
    vals = Reference(ws_cd, min_col=j+2, max_col=j+2, min_row=2, max_row=1+num_active)
    s = Series(vals, title=trainer)
    s.graphicalProperties.solidFill = TRAINER_COLORS.get(trainer, "CCCCCC")
    bar1.series.append(s)
bar1.set_categories(cats)
ws_dash.add_chart(bar1, f"B{chart_anchor_row}")

# Chart 2: Pie — Online vs Offline
pie = PieChart()
pie.width  = 10
pie.height = 12
pie.title  = None
pie.legend.position = "b"
pie_vals  = Reference(ws_cd, min_col=2, min_row=row_session+1, max_row=row_session+2)
pie_cats  = Reference(ws_cd, min_col=1, min_row=row_session+1, max_row=row_session+2)
pie_s     = Series(pie_vals, title="Session Type")

slice_online = DataPoint(idx=0)
slice_online.graphicalProperties.solidFill = BLUE_VAL
slice_offline = DataPoint(idx=1)
slice_offline.graphicalProperties.solidFill = GREEN_VAL
pie_s.points = [slice_online, slice_offline]

pie.series.append(pie_s)
pie.dLbls = DataLabelList(showPercent=True, showVal=False, showCatName=False)
pie.set_categories(pie_cats)
ws_dash.add_chart(pie, f"G{chart_anchor_row}")

# Chart 3: Bar — Hours by Activity Type
bar2 = BarChart()
bar2.type     = "bar"
bar2.grouping = "clustered"
bar2.width    = 17
bar2.height   = 12
bar2.title    = None
bar2.legend   = None
bar2.y_axis.title = "Hours"
bar2_vals = Reference(ws_cd, min_col=2, min_row=row_act+1, max_row=row_act+n_types)
bar2_cats = Reference(ws_cd, min_col=1, min_row=row_act+1, max_row=row_act+n_types)
bar2_s    = Series(bar2_vals, title="Hours")
bar2_s.graphicalProperties.solidFill = "00BFA5"
bar2.series.append(bar2_s)
bar2.set_categories(bar2_cats)
ws_dash.add_chart(bar2, f"K{chart_anchor_row}")

# ══════════════════════════════════════════════════════════════════════════════
# MONTHLY TRAINING LOG SHEETS
# ══════════════════════════════════════════════════════════════════════════════
print("Updating monthly log sheets…")
for month_name in MONTHS_ORDER:
    if month_name not in wb.sheetnames:
        continue
    ws = wb[month_name]
    
    # Clear old logs
    if ws.max_row >= 7:
        ws.delete_rows(7, ws.max_row - 6)
        
    acts = [a for a in all_acts if a.get("month") == month_name]
    
    if acts:
        # Write rows
        for i, act in enumerate(acts):
            row = 7 + i
            is_online = act.get("session","") == "Online"
            bg = "E3F2FD" if is_online else "E8F5E9"
            row_vals = [
                i+1,
                act.get("type",""),
                act.get("details") or act.get("training_details",""),
                act.get("start") or act.get("date",""),
                act.get("end",""),
                act.get("trainer",""),
                act.get("session",""),
                act.get("topic",""),
                to_float(act.get("hours",0)),
            ]
            for j, v in enumerate(row_vals):
                cell = ws.cell(row=row, column=j+1, value=v)
                cell.fill  = fill(bg)
                cell.font  = font("111827", size=10)
                cell.alignment = align("center" if j==0 else ("right" if j==8 else "left"), "center")
                cell.border = thin_border("DDDDDD")
                
        # Total row immediately after data
        total_row = 7 + len(acts)
        ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=8)
        ws.cell(row=total_row, column=1, value="TOTAL MAN HOURS THIS MONTH")
        style_range(ws, f"A{total_row}:H{total_row}", fill_=fill(DARK_BG), font_=font(TEAL, bold=True, size=10), align_=align("left", "center"))
        
        c_val = ws.cell(row=total_row, column=9, value=f"=SUM(I7:I{total_row-1})")
        c_val.fill  = fill(DARK_BG)
        c_val.font  = font("F5A623", bold=True, size=10)
        c_val.alignment = align("right", "center")
        c_val.number_format = "0.00"
    else:
        # Placeholder for empty months (A7:I7 placeholder, row 22 total)
        ws.merge_cells(start_row=7, start_column=1, end_row=7, end_column=9)
        ws.cell(row=7, column=1, value="No training activities recorded yet — add entries below")
        style_range(ws, "A7:I7", fill_=fill("F5F7FA"), font_=font("546E7A", italic=True, size=10), align_=align("center", "center"))
        
        total_row = 22
        ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=8)
        ws.cell(row=total_row, column=1, value="TOTAL MAN HOURS THIS MONTH")
        style_range(ws, f"A{total_row}:H{total_row}", fill_=fill(DARK_BG), font_=font(TEAL, bold=True, size=10), align_=align("left", "center"))
        
        c_val = ws.cell(row=total_row, column=9, value="=SUM(I7:I21)")
        c_val.fill  = fill(DARK_BG)
        c_val.font  = font("F5A623", bold=True, size=10)
        c_val.alignment = align("right", "center")
        c_val.number_format = "0.00"

# ── Save ────────────────────────────────────────────────────────────────────────
wb.save(OUT_PATH)
print(f"\n✅  Saved: {OUT_PATH}")

# Save a sibling copy inside the project folder for local HTTP server downloads
ALT_OUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Zerodha_Training_Dashboard.xlsx")
try:
    wb.save(ALT_OUT_PATH)
    print(f"✅  Saved local server copy: {ALT_OUT_PATH}")
except Exception as e:
    print(f"⚠️  Could not save copy to project folder: {e}")
print(f"   Sheets: Dashboard, CD, Data, and monthly sheets.")
print(f"\nTo use the filter: open the file → Dashboard tab → change cell D6 to a month name")
print("(e.g. 'May 2026') — all KPI cards and charts update automatically.")
