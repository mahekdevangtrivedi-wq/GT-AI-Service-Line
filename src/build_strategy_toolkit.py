# -*- coding: utf-8 -*-
"""GT AI Strategy & Roadmap Toolkit (Excel): use-case scoring, SWOT/TOWS, Gantt roadmap, KPI tracker, AI Ready7 traceability."""
import sys, os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
sys.path.insert(0, os.path.dirname(__file__))
from strategy_data import *

OUT = sys.argv[1]
GT = "4F2D7F"; GT2 = "7B5BA6"; LAV = "EDE7F6"; INPUT = "FFF9E6"
thin = Side(style="thin", color="C9C9C9"); B = Border(left=thin, right=thin, top=thin, bottom=thin)
WT = Alignment(wrap_text=True, vertical="top"); WC = Alignment(wrap_text=True, vertical="center"); CC = Alignment(wrap_text=True, vertical="center", horizontal="center")
wb = openpyxl.Workbook()


def title(ws, t, sub, span):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    c = ws.cell(1, 1, t); c.font = Font(size=16, bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=GT); c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 32
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
    s = ws.cell(2, 1, sub); s.font = Font(italic=True, size=9.5, color=GT); s.alignment = WC; ws.row_dimensions[2].height = 30


def cell(ws, r, c, v, bold=False, fill=None, align=WC, size=9.5, color="000000", fmt=None):
    x = ws.cell(r, c, v); x.font = Font(size=size, bold=bold, color=color); x.alignment = align; x.border = B
    if fill: x.fill = PatternFill("solid", fgColor=fill)
    if fmt: x.number_format = fmt
    return x


def hdr(ws, r, heads, fill=GT):
    for i, h in enumerate(heads, 1): cell(ws, r, i, h, bold=True, fill=fill, align=CC, color="FFFFFF")
    ws.row_dimensions[r].height = 36


# ------------------------------------------------ Guide
g = wb.active; g.title = "Guide"
title(g, "GT AI Strategy & Roadmap Toolkit", "Companion to 'GT Bahrain AI Strategy & Roadmap 2026-2028'. Populated with the GT Bahrain case; reusable as a template for client AI strategy engagements.", 3)
g.column_dimensions["A"].width = 3; g.column_dimensions["B"].width = 30; g.column_dimensions["C"].width = 110
rows = [("1. SWOT & TOWS", "Situational analysis: strengths, weaknesses, opportunities, threats and the resulting strategic options (TOWS)."),
        ("2. Use-Case Scoring", "Score candidate use cases 1-5 on four value and four feasibility criteria (weights editable in row 5). Value, feasibility, quadrant and priority rank calculate automatically."),
        ("3. Roadmap (Gantt)", "Enter start month (1 = Oct 2026) and duration for each activity; the Gantt bars draw automatically. Track owner and status."),
        ("4. KPI Tracker", "Baseline, targets and actuals for strategy KPIs; RAG status calculates from the actual value where numeric."),
        ("5. Ready7 Traceability", "Maps every AI Ready7 item not fully met to a roadmap response, horizon and workstream."),
        ("Colour legend", "Yellow = input  |  Purple = header / calculated")]
for i, (k, v) in enumerate(rows):
    cell(g, 4 + i, 2, k, bold=True, fill=LAV); cell(g, 4 + i, 3, v, align=WT); g.row_dimensions[4 + i].height = 32

# ------------------------------------------------ SWOT & TOWS
sw = wb.create_sheet("1. SWOT & TOWS")
title(sw, "1. SWOT & TOWS Analysis", "Internal factors from the AI Ready7 baseline (Sept 2026); external factors from the Bahrain / GCC landscape scan.", 4)
for c, w in zip("ABCD", [4, 70, 4, 70]): sw.column_dimensions[c].width = w
for j, (k, hname, col) in enumerate([("S", "STRENGTHS", GT), ("W", "WEAKNESSES", GT2)]):
    cell(sw, 4, 2 + j * 2, hname, bold=True, fill=col, color="FFFFFF", align=CC)
    for i, it in enumerate(SWOT[k]): cell(sw, 5 + i, 2 + j * 2, it, fill=INPUT, align=WT); sw.row_dimensions[5 + i].height = 30
for j, (k, hname, col) in enumerate([("O", "OPPORTUNITIES", "00A7B5"), ("T", "THREATS", "E8705F")]):
    cell(sw, 12, 2 + j * 2, hname, bold=True, fill=col, color="FFFFFF", align=CC)
    for i, it in enumerate(SWOT[k]): cell(sw, 13 + i, 2 + j * 2, it, fill=INPUT, align=WT); sw.row_dimensions[13 + i].height = 30
cell(sw, 20, 2, "TOWS - STRATEGIC OPTIONS", bold=True, fill=GT, color="FFFFFF", align=CC); sw.merge_cells("B20:D20")
for i, (h, d) in enumerate(TOWS):
    cell(sw, 21 + i, 2, h, bold=True, fill=LAV); cell(sw, 21 + i, 4, d, align=WT); sw.merge_cells(start_row=21 + i, start_column=2, end_row=21 + i, end_column=3)
    sw.row_dimensions[21 + i].height = 42

# ------------------------------------------------ Use-case scoring
us = wb.create_sheet("2. Use-Case Scoring")
title(us, "2. Use-Case Prioritisation", "Score 1 (low) - 5 (high). Technical simplicity and ease of change: 5 = easiest. Weights (row 5) must total 100% per dimension. Quadrant threshold (cell B3) editable.", 16)
us["A3"] = "Quadrant threshold"; us["A3"].font = Font(bold=True); us["B3"] = 3.5; us["B3"].fill = PatternFill("solid", fgColor=INPUT)
heads = ["#", "Use case", "Area", "Risk tier"] + [c for c, _ in VALUE_CRIT] + [c for c, _ in FEAS_CRIT] + ["VALUE", "FEASIBILITY", "Quadrant", "Priority rank", "Horizon"]
hdr(us, 4, heads)
cell(us, 5, 1, "", fill=LAV); cell(us, 5, 2, "Weights", bold=True, fill=LAV); cell(us, 5, 3, "", fill=LAV); cell(us, 5, 4, "", fill=LAV)
for i, (_, w) in enumerate(VALUE_CRIT + FEAS_CRIT): cell(us, 5, 5 + i, w, fill=INPUT, align=CC, fmt="0%")
cell(us, 5, 13, "=SUM(E5:H5)", fmt="0%", align=CC, fill=LAV); cell(us, 5, 14, "=SUM(I5:L5)", fmt="0%", align=CC, fill=LAV)
for c in (15, 16, 17): cell(us, 5, c, "", fill=LAV)
n = len(USE_CASES)
for k, uc in enumerate(USE_CASES):
    r = 6 + k
    cell(us, r, 1, uc[0], align=CC); cell(us, r, 2, uc[1], bold=True); cell(us, r, 3, uc[6]); cell(us, r, 4, uc[4], align=CC, fill=INPUT)
    vs, fs = SUBSCORES[uc[0]]
    for i, v in enumerate(vs + fs): cell(us, r, 5 + i, v, fill=INPUT, align=CC)
    cell(us, r, 13, f"=ROUND(SUMPRODUCT(E{r}:H{r},$E$5:$H$5),2)", bold=True, align=CC, fmt="0.00")
    cell(us, r, 14, f"=ROUND(SUMPRODUCT(I{r}:L{r},$I$5:$L$5),2)", bold=True, align=CC, fmt="0.00")
    cell(us, r, 15, f'=IF(AND(M{r}>=$B$3,N{r}>=$B$3),"Quick win",IF(M{r}>=$B$3,"Strategic bet",IF(N{r}>=$B$3,"Fill-in","Deprioritise")))', bold=True, align=CC)
    cell(us, r, 16, f"=RANK(M{r}+N{r}+ROW()/10000,$R$6:$R${5+n})", align=CC)
    us.cell(r, 18, f"=M{r}+N{r}+ROW()/10000")
    cell(us, r, 17, uc[5], fill=INPUT, align=CC)
    us.row_dimensions[r].height = 22
us.column_dimensions["R"].hidden = True
dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5", showErrorMessage=True, error="Enter a score from 1 to 5"); us.add_data_validation(dv); dv.add(f"E6:L{5+n}")
dv2 = DataValidation(type="list", formula1='"Low,Medium,High,Critical"'); us.add_data_validation(dv2); dv2.add(f"D6:D{5+n}")
for lab, col in [("Quick win", "C6E0B4"), ("Strategic bet", "D9CCEB"), ("Fill-in", "BDD7EE"), ("Deprioritise", "E7E6E6")]:
    us.conditional_formatting.add(f"O6:O{5+n}", CellIsRule(operator="equal", formula=[f'"{lab}"'], fill=PatternFill("solid", fgColor=col)))
for lab, col in [("Low", "C6E0B4"), ("Medium", "FFE699"), ("High", "F4B183"), ("Critical", "E06666")]:
    us.conditional_formatting.add(f"D6:D{5+n}", CellIsRule(operator="equal", formula=[f'"{lab}"'], fill=PatternFill("solid", fgColor=col)))
for c, w in zip(range(1, 18), [4, 36, 20, 9] + [11] * 8 + [9, 11, 13, 9, 8]): us.column_dimensions[get_column_letter(c)].width = w
us.freeze_panes = "C6"

# ------------------------------------------------ Gantt
gt = wb.create_sheet("3. Roadmap (Gantt)")
title(gt, "3. Implementation Roadmap - Oct 2026 to Mar 2028", "Start month: 1 = Oct 2026. Bars draw automatically from Start and Duration. Dashed milestones: H2 starts Apr 2027 (month 7), H3 starts Jan 2028 (month 16).", 26)
months = ["Oct-26", "Nov-26", "Dec-26", "Jan-27", "Feb-27", "Mar-27", "Apr-27", "May-27", "Jun-27", "Jul-27", "Aug-27", "Sep-27", "Oct-27", "Nov-27", "Dec-27", "Jan-28", "Feb-28", "Mar-28"]
heads = ["Workstream", "Activity", "Start", "Duration", "End", "Owner", "Status"]
hdr(gt, 4, heads + [""] * 18)
for i, m in enumerate(months):
    c = gt.cell(4, 8 + i, m); c.font = Font(size=8, bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=GT if (i // 3) % 2 == 0 else GT2); c.alignment = Alignment(text_rotation=90, horizontal="center", vertical="center"); c.border = B
    gt.cell(3, 8 + i, i + 1).alignment = CC
    gt.column_dimensions[get_column_letter(8 + i)].width = 3.6
OWN = {"WS1": "Risk & Quality / AI CoE", "WS2": "HR / L&D", "WS3": "IT / Cyber", "WS4": "AI CoE + service lines", "WS5": "AI service line lead"}
for k, (ws_, act, st, du) in enumerate(GANTT):
    r = 5 + k
    cell(gt, r, 1, ws_, bold=True, size=8.5); cell(gt, r, 2, act, size=8.5)
    cell(gt, r, 3, st + 1, fill=INPUT, align=CC); cell(gt, r, 4, du, fill=INPUT, align=CC)
    cell(gt, r, 5, f"=INDEX($H$4:$Y$4,MIN(18,C{r}+D{r}-1))", align=CC, size=8.5)
    cell(gt, r, 6, OWN[ws_[:3]], fill=INPUT, size=8.5); cell(gt, r, 7, "Not started", fill=INPUT, size=8.5)
    for i in range(18): gt.cell(r, 8 + i).border = B
    gt.row_dimensions[r].height = 20
last = 4 + len(GANTT)
colors = {"WS1": GT, "WS2": "00A7B5", "WS3": "2E1A47", "WS4": "E8705F", "WS5": GT2}
for wsk, col in colors.items():
    gt.conditional_formatting.add(f"H5:Y{last}", FormulaRule(formula=[f'AND(LEFT($A5,3)="{wsk}",H$3>=$C5,H$3<$C5+$D5)'], fill=PatternFill("solid", fgColor=col)))
dv3 = DataValidation(type="list", formula1='"Not started,Planned,In progress,Completed,Delayed"'); gt.add_data_validation(dv3); dv3.add(f"G5:G{last}")
gt.conditional_formatting.add(f"G5:G{last}", CellIsRule(operator="equal", formula=['"Delayed"'], fill=PatternFill("solid", fgColor="F4B183")))
gt.conditional_formatting.add(f"G5:G{last}", CellIsRule(operator="equal", formula=['"Completed"'], fill=PatternFill("solid", fgColor="C6E0B4")))
for c, w in zip("ABCDEFG", [24, 46, 7, 8, 8, 20, 12]): gt.column_dimensions[c].width = w
for r in range(5, last + 1):
    for col in ("N", "W"):  # month 7 (Apr-27) and month 16 (Jan-28)
        gt[f"{col}{r}"].border = Border(left=Side(style="dashed", color="00A7B5"), right=thin, top=thin, bottom=thin)
gt.freeze_panes = "H5"
gt.page_setup.orientation = "landscape"; gt.page_setup.fitToWidth = 1; gt.page_setup.fitToHeight = 0; gt.sheet_properties.pageSetUpPr.fitToPage = True

# ------------------------------------------------ KPI tracker
kp = wb.create_sheet("4. KPI Tracker")
title(kp, "4. KPI Tracker", "Update 'Actual' each quarter. Status is set manually (On track / At risk / Off track) based on trajectory to target.", 8)
hdr(kp, 4, ["KPI", "Baseline (Sept 2026)", "Target end-2027", "Target end-2028", "Owner", "Actual (latest)", "Status", "Commentary"])
for k, row in enumerate(KPIS):
    r = 5 + k
    for i, v in enumerate(row, 1): cell(kp, r, i, v, bold=(i == 1))
    for c in (6, 7, 8): cell(kp, r, c, None, fill=INPUT)
    kp.row_dimensions[r].height = 30
dv4 = DataValidation(type="list", formula1='"On track,At risk,Off track"'); kp.add_data_validation(dv4); dv4.add(f"G5:G{4+len(KPIS)}")
for lab, col in [("On track", "C6E0B4"), ("At risk", "FFE699"), ("Off track", "F4B183")]:
    kp.conditional_formatting.add(f"G5:G{4+len(KPIS)}", CellIsRule(operator="equal", formula=[f'"{lab}"'], fill=PatternFill("solid", fgColor=col)))
for c, w in zip("ABCDEFGH", [42, 26, 22, 22, 22, 16, 12, 40]): kp.column_dimensions[c].width = w

# ------------------------------------------------ Ready7 traceability
tr = wb.create_sheet("5. Ready7 Traceability")
title(tr, "5. AI Ready7 Gap-to-Roadmap Traceability", f"AI Ready7 self-assessment, GT Bahrain, 17 Sept 2026 - overall {READY7_OVERALL}% (Informed). Pillar scores: " + "; ".join(f"{p} {v:g}%" for p, v in READY7), 6)
hdr(tr, 4, ["Item", "Title", "Result", "Roadmap response", "Horizon", "Workstream"])
for k, row in enumerate(READY7_ITEMS):
    r = 5 + k
    for i, v in enumerate(row, 1): cell(tr, r, i, v, bold=(i == 1), align=WC if i != 2 else WC)
    tr.row_dimensions[r].height = 22
for lab, col in [("Not Met", "F4B183"), ("Some Met", "FFE699"), ("Most Met", "C6E0B4")]:
    tr.conditional_formatting.add(f"C5:C{4+len(READY7_ITEMS)}", CellIsRule(operator="equal", formula=[f'"{lab}"'], fill=PatternFill("solid", fgColor=col)))
for c, w in zip("ABCDEF", [12, 40, 11, 70, 9, 11]): tr.column_dimensions[c].width = w

for ws_ in wb.worksheets:
    ws_.sheet_properties.tabColor = GT
    if ws_.title != "3. Roadmap (Gantt)":
        ws_.page_setup.orientation = "landscape"; ws_.page_setup.fitToWidth = 1; ws_.page_setup.fitToHeight = 0; ws_.sheet_properties.pageSetUpPr.fitToPage = True
wb.save(OUT)
print("saved", OUT)
