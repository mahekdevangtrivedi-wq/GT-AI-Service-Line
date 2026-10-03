# -*- coding: utf-8 -*-
"""GT AI Opportunity Discovery Toolkit (Excel) - where exactly does AI fit in an organisation?

Usage: build_opportunity_toolkit.py OUT.xlsx [example]
"""
import sys, os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.pagebreak import Break
sys.path.insert(0, os.path.dirname(__file__))
import printfix
from opportunity_data import *

OUT = sys.argv[1]
EXAMPLE = len(sys.argv) > 2 and sys.argv[2] == "example"

GT = "4F2D7F"; GT2 = "7B5BA6"; LAV = "EDE7F6"; LAV2 = "D9CCEB"; INPUT = "FFF9E6"; GREY = "F2F2F2"; DARK = "3B2160"
TEAL = "00A7B5"; CORAL = "E8705F"
VFILL = {"AI fit now": "A9D08E", "Automate - AI not needed": "9BC2E6", "AI later - close data gap first": "FFE699",
         "Stretch - partner or later stage": "F4B183", "Redesign / standardise first": "D9D9D9"}
RFILL = {"Low": "C6E0B4", "Medium": "FFE699", "High": "F4B183"}
thin = Side(style="thin", color="C9C9C9"); B = Border(left=thin, right=thin, top=thin, bottom=thin)
WT = Alignment(wrap_text=True, vertical="top"); WC = Alignment(wrap_text=True, vertical="center")
CC = Alignment(wrap_text=True, vertical="center", horizontal="center")
wb = openpyxl.Workbook()
NF = len(FAMILIES)


def title(ws, t, sub, span):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    c = ws.cell(1, 1, t); c.font = Font(size=16, bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=GT)
    c.alignment = Alignment(vertical="center", indent=1); ws.row_dimensions[1].height = 32
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
    s = ws.cell(2, 1, sub); s.font = Font(italic=True, size=9.5, color=GT); s.alignment = WC; ws.row_dimensions[2].height = 32


def cell(ws, r, c, v, bold=False, fill=None, align=WC, size=9.5, color="000000", fmt=None, border=True, italic=False):
    x = ws.cell(r, c, v); x.font = Font(size=size, bold=bold, color=color, italic=italic); x.alignment = align
    if border: x.border = B
    if fill: x.fill = PatternFill("solid", fgColor=fill)
    if fmt: x.number_format = fmt
    return x


def band(ws, r, text, span, fill=GT, c0=1):
    ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + span - 1)
    x = ws.cell(r, c0, text); x.font = Font(bold=True, color="FFFFFF", size=10.5); x.fill = PatternFill("solid", fgColor=fill)
    x.alignment = Alignment(vertical="center", indent=1); ws.row_dimensions[r].height = 22


def merged(ws, r, c1, c2, v, **kw):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    for c in range(c1, c2 + 1): ws.cell(r, c).border = B
    if kw.get("fill"):
        for c in range(c1, c2 + 1): ws.cell(r, c).fill = PatternFill("solid", fgColor=kw["fill"])
    return cell(ws, r, c1, v, **kw)


def dv_list(ws, src, rng, prompt=None, strict=True):
    dv = DataValidation(type="list", formula1=src, allow_blank=True, showErrorMessage=strict)
    if prompt: dv.prompt = prompt; dv.showInputMessage = True
    ws.add_data_validation(dv); dv.add(rng)


def verdict_cf(ws, rng):
    for v, f in VFILL.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{v}"'], fill=PatternFill("solid", fgColor=f)))


def risk_cf(ws, rng):
    for v, f in RFILL.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{v}"'], fill=PatternFill("solid", fgColor=f)))


# ================================================================== Lists (hidden)
L = wb.active; L.title = "Lists"
L["A1"] = "Task nature"; L["B1"] = "Solution family"; L["C1"] = "Level"; L["D1"] = "AI?"; L["E1"] = "Data req"; L["F1"] = "Potential"
L["G1"] = "Typical solutions"; L["H1"] = "KPIs"; L["I1"] = "Design notes"
for i, f in enumerate(FAMILIES, 2):
    for j, v in enumerate(f, 1): L.cell(i, j, v)
FR = f"$2:${NF + 1}"
LA = f"Lists!$A$2:$A${NF + 1}"; LB = f"Lists!$B$2:$B${NF + 1}"


def fam(col):  # family attribute range
    return f"Lists!${col}$2:${col}${NF + 1}"


def put_pairs(col, rows, start=2):
    for i, (a, b) in enumerate(rows, start):
        L.cell(i, col, a); L.cell(i, col + 1, b)
    return f"Lists!${get_column_letter(col)}${start}:${get_column_letter(col)}${start + len(rows) - 1}", \
           f"Lists!${get_column_letter(col + 1)}${start}:${get_column_letter(col + 1)}${start + len(rows) - 1}"


DA_O, DA_V = put_pairs(13, DATA_AVAIL, 2)      # M2:N5
VO_O, VO_V = put_pairs(13, VOLUME, 8)          # M8:N10
ST_O, ST_V = put_pairs(13, STANDARD, 13)       # M13:N15
AU_O, AU_V = put_pairs(13, AUTONOMY, 18)       # M18:N20
for i, v in enumerate(YESNO, 23): L.cell(i, 13, v)
YN = "Lists!$M$23:$M$24"
for i, v in enumerate(FUNCTIONS, 2): L.cell(i, 16, v)
FN = f"Lists!$P$2:$P${len(FUNCTIONS) + 1}"
for i, v in enumerate(SECTORS, 2): L.cell(i, 17, v)
SEC = f"Lists!$Q$2:$Q${len(SECTORS) + 1}"
POS_RANGE = {}
for k, (qid, axis, q, g, opts) in enumerate(POSITION_Q):
    col = 19 + k  # S..
    L.cell(1, col, qid)
    for i, o in enumerate(opts, 2): L.cell(i, col, o)
    cl = get_column_letter(col); POS_RANGE[qid] = f"Lists!${cl}$2:${cl}$5"
for i, (code, name, ceil, head, desc) in enumerate(ARCHETYPES, 2):
    for j, v in enumerate((code, name, ceil, head, desc), 33): L.cell(i, j, v)   # AG..AK
ARC = lambda col: f"Lists!${col}$2:${col}$5"
for i, (lvl, name) in enumerate(LEVELS, 2): L.cell(i, 38, lvl); L.cell(i, 39, name)   # AL, AM
for i, (t, c) in enumerate(CONTROLS, 2): L.cell(i, 41, t); L.cell(i, 42, c)          # AO, AP
for i, (v, m, n) in enumerate(VERDICTS, 2): L.cell(i, 44, v); L.cell(i, 45, m); L.cell(i, 46, n)   # AR, AS, AT
for i, h in enumerate(HORIZONS, 2): L.cell(i, 48, h)   # AV
VER = "Lists!$AR$2:$AR$6"
for i in range(1, 51): L.cell(i + 1, 50, i)   # AX: 1..50
L.sheet_state = "hidden"

# ================================================================== Guide
g = wb.create_sheet("Guide", 0)
title(g, "GT AI Opportunity Discovery Toolkit", "Where exactly does AI fit in this organisation - and where is automation enough? "
      "A facilitated workshop tool of the GT AI Strategy & Roadmap service.", 3)
g.column_dimensions["A"].width = 3; g.column_dimensions["B"].width = 30; g.column_dimensions["C"].width = 108
r = 4
band(g, r, "WHAT THIS TOOLKIT ANSWERS", 2, c0=2); r += 1
for k, v in [("1. Is the organisation positioned for AI?", "Do you have the DATA (digital, joined-up, reliable, with history and outcomes) and the SCALE / CAPACITY (volumes, people, budget, "
                                                     "technology team) for AI - or is automation enough for now? Result: one of four archetypes and a ceiling on solution complexity."),
             ("2. Where exactly does AI fit?", "Process by process: what kind of work is it, which solution family fits (digitise, automate, off-the-shelf AI, configured AI, "
                                               "custom / agentic AI), is the data there, and what is the risk tier?"),
             ("3. What should be done first?", "Value x feasibility, risk-adjusted priority, horizon (H1 / H2 / H3), opportunity map, roadmap lanes and one-page use-case canvases."),
             ("Not a readiness test", "Governance and capability maturity are measured by the GT AI Readiness & Risk Assessment (AI Ready7-based). This toolkit does not score "
                                      "readiness - it locates the opportunities and the right type of solution for each.")]:
    cell(g, r, 2, k, bold=True, fill=LAV); cell(g, r, 3, v, align=WT); g.row_dimensions[r].height = 44; r += 1
r += 1; band(g, r, "HOW TO USE IT (WORKSHOP SEQUENCE)", 2, c0=2); r += 1
steps = [("1. AI Position", "Answer 12 questions with leadership, IT and finance (about 30 minutes). Data Position and Scale & Capacity scores place the organisation in an archetype."),
         ("2. Process Library", "Use the 60+ cross-industry and sector processes (Bahrain priority sectors) as prompts in discovery workshops with each function."),
         ("3. AI Fit Diagnostic", "List up to 50 processes / pain points and answer 10 inputs each (yellow cells). Solution family, data gate, verdict, value, feasibility, risk tier, "
                                  "quadrant, priority rank, horizon and indicative hours saved calculate automatically."),
         ("4. Opportunity Map", "Where AI fits by function, the verdict mix, hours released and the top 10 ranked opportunities - the headline output for the client."),
         ("5. Roadmap", "Opportunities placed automatically in H1 / H2 / H3 lanes plus the enabling foundations they depend on."),
         ("6. Use-Case Canvas", "Pick an opportunity number; a one-page canvas is produced (solution, data, risk tier, minimum controls, KPIs, next step) for sign-off."),
         ("Reference", "Solution families, archetypes, verdicts, minimum controls by risk tier and the standards behind each step."),
         ("Colour legend", "Yellow = input  |  Purple / lavender = header or calculated  |  Verdict colours: green = AI fit now, blue = automate, yellow = data gap, "
                           "orange = stretch, grey = redesign first")]
for k, v in steps:
    cell(g, r, 2, k, bold=True, fill=LAV); cell(g, r, 3, v, align=WT); g.row_dimensions[r].height = 34; r += 1
r += 1; band(g, r, "DECISION LOGIC (SUMMARY)", 2, c0=2); r += 1
logic = [("Solution family", "Determined by the nature of the task (10 task types, ISO/IEC 22989 / OECD-based). Each family has a complexity level 1-5 and a minimum data requirement."),
         ("Data gate", "Data available for the process >= the family's minimum requirement -> 'Data ready', otherwise 'Data gap'."),
         ("Organisational fit", "Family level <= the archetype's ceiling -> 'Within current position', otherwise 'Beyond current position'."),
         ("Verdict (in order)", "Process not standardised -> Redesign first;  family needs no AI -> Automate;  data gap -> AI later;  beyond position -> Stretch;  otherwise -> AI fit now."),
         ("Value (1-5)", "30% volume + 25% effort (hours / month) + 25% strategic importance + 20% customer / quality impact."),
         ("Feasibility (1-5)", "30% data fit + 25% process standardisation + 25% solution simplicity + 20% organisational fit."),
         ("Risk tier", "AIAF-style inherent risk: affects individuals' rights / access (+2), personal or confidential data (+1), autonomy (+0 to +2), customer / quality impact of errors (+1 if 4, +2 if 5). "
                       "4+ = High, 2-3 = Medium, 0-1 = Low; non-AI automation is Low unless it affects individuals."),
         ("Priority", "Value x Feasibility, reduced 15% for High risk and 20% where the process must be redesigned first. Ranked 1..n.")]
for k, v in logic:
    cell(g, r, 2, k, bold=True, fill=LAV); cell(g, r, 3, v, align=WT); g.row_dimensions[r].height = 32; r += 1
g.page_setup.orientation = "portrait"; g.page_setup.fitToWidth = 1; g.page_setup.fitToHeight = 0; g.sheet_properties.pageSetUpPr.fitToPage = True

# ================================================================== 1. AI Position
P = wb.create_sheet("1. AI Position", 1)
title(P, "1. Organisational AI Position", "Does the organisation have the data and the scale / capacity for AI - or is automation enough for now? "
      "Select one answer per question (yellow). This screen does not measure governance maturity.", 6)
for c, w in zip("ABCDEF", [6, 46, 44, 34, 8, 4]): P.column_dimensions[c].width = w
prof = ["Organisation", "Sector", "Scope of discovery", "Assessment date", "Facilitated by"]
for i, k in enumerate(prof):
    cell(P, 4 + i, 1, "", fill=LAV); cell(P, 4 + i, 2, k, bold=True, fill=LAV)
    merged(P, 4 + i, 3, 4, "", fill=INPUT)
dv_list(P, f"={SEC}", "C5")
if EXAMPLE:
    for i, k in enumerate(["Organisation", "Sector", "Scope", "Assessment date", "Facilitated by"]):
        P.cell(4 + i, 3).value = EXAMPLE_PROFILE[k]
r = 10
hdrs = ["#", "Question", "Answer (select)", "Why it matters", "Points"]
for j, h in enumerate(hdrs, 1): cell(P, r, j, h, bold=True, fill=GT, color="FFFFFF", align=CC)
P.row_dimensions[r].height = 24
POS_ROW = {}
r += 1
for axis, label, fillc in [("Data", "A. DATA POSITION - is the data there for AI?", GT2), ("Scale", "B. SCALE & CAPACITY - is there enough volume and capacity to justify and run AI?", GT2)]:
    band(P, r, label, 5, fill=fillc); r += 1
    for qid, ax, q, gd, opts in POSITION_Q:
        if ax != axis: continue
        cell(P, r, 1, qid, bold=True, fill=LAV, align=CC); cell(P, r, 2, q, align=WT)
        a = cell(P, r, 3, "", fill=INPUT, align=WT)
        if EXAMPLE: a.value = opts[EXAMPLE_POSITION[qid]]
        dv_list(P, f"={POS_RANGE[qid]}", f"C{r}")
        cell(P, r, 4, gd, align=WT, size=8.5, color="595959")
        cell(P, r, 5, f'=IFERROR(MATCH(C{r},{POS_RANGE[qid]},0),"")', align=CC, fill=LAV)
        P.row_dimensions[r].height = 42; POS_ROW[qid] = r; r += 1
d_rows = [POS_ROW[q[0]] for q in POSITION_Q if q[1] == "Data"]; s_rows = [POS_ROW[q[0]] for q in POSITION_Q if q[1] == "Scale"]
r += 1; band(P, r, "RESULT - AI POSITION", 5); r += 1
RES = r
d_sum = "+".join(f"N(E{x})" for x in d_rows); s_sum = "+".join(f"N(E{x})" for x in s_rows)
d_cnt = ",".join(f"E{x}" for x in d_rows); s_cnt = ",".join(f"E{x}" for x in s_rows)
rows = [("Answered", f'=COUNT({d_cnt})+COUNT({s_cnt})&" of 12"', None),
        ("Data Position score", f'=IF(COUNT({d_cnt})=6,({d_sum}-6)/18,"")', "0%"),
        ("Scale & Capacity score", f'=IF(COUNT({s_cnt})=6,({s_sum}-6)/18,"")', "0%"),
        ("Archetype", f'=IF(OR(C{RES + 1}="",C{RES + 2}=""),"",IF(C{RES + 1}<0.5,IF(C{RES + 2}<0.5,1,2),IF(C{RES + 2}<0.5,3,4)))', None),
        ("AI position", f'=IF(C{RES + 3}="","Complete all 12 questions",INDEX({ARC("AH")},C{RES + 3}))', None),
        ("Solution ceiling today", f'=IF(C{RES + 3}="","",INDEX({ARC("AI")},C{RES + 3})&" - "&INDEX(Lists!$AM$2:$AM$6,INDEX({ARC("AI")},C{RES + 3})))', None),
        ("What it means", f'=IF(C{RES + 3}="","",INDEX({ARC("AJ")},C{RES + 3}))', None),
        ("Detail", f'=IF(C{RES + 3}="","",INDEX({ARC("AK")},C{RES + 3}))', None),
        ("Foundation flag", f'=IF(E{POS_ROW["D1"]}=1,"Core processes are largely manual / paper-based: digitise them first - AI and automation need digital inputs.",'
                            f'IF(AND(N(E{POS_ROW["D2"]})>0,N(E{POS_ROW["D2"]})<=2,N(E{POS_ROW["D6"]})<=1),"Little digital history and no recorded outcomes: predictive / ML use cases will fail the data gate until data is captured.",'
                            f'"No foundation red flags from the position screen."))', None)]
for i, (k, f, fmt) in enumerate(rows):
    cell(P, RES + i, 2, k, bold=True, fill=LAV)
    x = merged(P, RES + i, 3, 5, f, fill=LAV2 if i in (4, 5) else None, bold=i in (1, 2, 4, 5), size=11 if i == 4 else 9.5, align=WT)
    if fmt: x.number_format = fmt
    P.row_dimensions[RES + i].height = 20 if i < 6 else 44
P.row_dimensions[RES + 7].height = 58
POS_SCORE_D, POS_SCORE_S, POS_ARCH = f"'1. AI Position'!$C${RES + 1}", f"'1. AI Position'!$C${RES + 2}", f"'1. AI Position'!$C${RES + 3}"
POS_NAME = f"'1. AI Position'!$C${RES + 4}"
CEIL = f"IF({POS_ARCH}=\"\",5,INDEX({ARC('AI')},{POS_ARCH}))"
POS_FLAG = f"'1. AI Position'!$C${RES + 8}"
# 2x2 matrix
r = RES + 10; band(P, r, "POSITION MATRIX - the highlighted box is where the organisation sits", 5); r += 1
MX = r
cell(P, MX, 1, "", border=False)
cell(P, MX, 2, "Data Position HIGH  (>= 50%)", bold=True, align=CC, fill=GREY)
cell(P, MX + 1, 2, "Data Position LOW  (< 50%)", bold=True, align=CC, fill=GREY)
cell(P, MX + 2, 3, "Scale & Capacity LOW  (< 50%)", bold=True, align=CC, fill=GREY)
cell(P, MX + 2, 4, "Scale & Capacity HIGH  (>= 50%)", bold=True, align=CC, fill=GREY)
quad = {(MX, 3): 2, (MX, 4): 3, (MX + 1, 3): 0, (MX + 1, 4): 1}
for (rr, cc), ai in quad.items():
    code, name, ceil, head, desc = ARCHETYPES[ai]
    x = cell(P, rr, cc, f"{code}. {name}\n{head}", align=CC, size=9.5)
    P.conditional_formatting.add(f"{get_column_letter(cc)}{rr}", FormulaRule(formula=[f"{POS_ARCH}={ai + 1}"], fill=PatternFill("solid", fgColor="A9D08E"), font=Font(bold=True)))
P.row_dimensions[MX].height = 66; P.row_dimensions[MX + 1].height = 66; P.row_dimensions[MX + 2].height = 22
P.print_area = f"A1:E{MX + 2}"
P.page_setup.orientation = "portrait"; P.page_setup.fitToWidth = 1; P.page_setup.fitToHeight = 1; P.sheet_properties.pageSetUpPr.fitToPage = True

# ================================================================== 2. Process Library
PL = wb.create_sheet("2. Process Library", 2)
title(PL, "2. Process Library - discovery prompts", "Cross-industry processes (APQC-style functions) plus Bahrain priority-sector processes. Use in workshops as prompts; copy relevant "
      "processes into the AI Fit Diagnostic and adapt the pain point. The suggested task nature is typical - confirm it with the process owner.", 6)
for c, w in zip("ABCDEF", [5, 26, 40, 40, 46, 30]): PL.column_dimensions[c].width = w
for j, h in enumerate(["#", "Function", "Process / activity", "Typical pain point", "Typical task nature -> solution family", "Typical sectors"], 1):
    cell(PL, 4, j, h, bold=True, fill=GT, color="FFFFFF", align=CC)
for i, (fn, pr, pain, ni, sec) in enumerate(LIBRARY, 5):
    cell(PL, i, 1, i - 4, align=CC); cell(PL, i, 2, fn, align=WT); cell(PL, i, 3, pr, align=WT, bold=True); cell(PL, i, 4, pain, align=WT)
    cell(PL, i, 5, f"{FAMILIES[ni][0]}  ->  {FAMILIES[ni][1]}", align=WT, size=8.5); cell(PL, i, 6, sec, align=WT, size=8.5)
    PL.row_dimensions[i].height = 30
PL.freeze_panes = "A5"; PL.print_title_rows = "4:4"
PL.page_setup.orientation = "landscape"; PL.page_setup.fitToWidth = 1; PL.page_setup.fitToHeight = 0; PL.sheet_properties.pageSetUpPr.fitToPage = True

# ================================================================== 3. AI Fit Diagnostic
D = wb.create_sheet("3. AI Fit Diagnostic", 3)
NROW = 50; R0 = 8; R1 = R0 + NROW - 1
title(D, "3. AI Fit Diagnostic - process by process", "Enter one process / pain point per row and answer the 10 yellow inputs. Everything to the right calculates. "
      "Effort = staff hours spent on the process per month (all people involved).", 30)
widths = [5, 22, 32, 30, 34, 18, 10, 22, 22, 9, 9, 10, 10, 22,
          26, 7, 7, 11, 15, 22, 8, 9, 9, 13, 9, 7, 19, 11, 6, 6]
for i, w in enumerate(widths, 1): D.column_dimensions[get_column_letter(i)].width = w
cell(D, 3, 1, "", fill=LAV, border=False)
merged(D, 3, 2, 3, "Organisation's AI position (from sheet 1):", bold=True, fill=LAV)
merged(D, 3, 4, 6, f"={POS_NAME}", bold=True, fill=LAV2)
merged(D, 3, 8, 10, "Solution ceiling level:", bold=True, fill=LAV)
cell(D, 3, 11, f"={CEIL}", bold=True, fill=LAV2, align=CC)
merged(D, 4, 2, 3, "Quadrant threshold (value and feasibility, 1-5):", bold=True, fill=LAV)
cell(D, 4, 4, 3.25, fill=INPUT, align=CC)
THR = "$D$4"
merged(D, 6, 1, 14, "INPUTS (yellow)", bold=True, fill=GT, color="FFFFFF", align=CC)
merged(D, 6, 15, 28, "AUTO-CALCULATED", bold=True, fill=DARK, color="FFFFFF", align=CC)
heads = ["#", "Function", "Process / activity", "Pain point / opportunity", "Nature of the task", "Volume", "Effort (hrs / month)",
         "Process standardised?", "Data available for this process", "Strategic importance (1-5)", "Customer / quality impact (1-5)",
         "Affects individuals' rights or access?", "Personal or confidential data?", "Intended autonomy",
         "Solution family", "Level", "AI?", "Data gate", "Organisational fit", "VERDICT", "Value", "Feasib-ility", "Risk tier", "Quadrant",
         "Priority score", "Rank", "Horizon", "Est. hours released / yr", "h", "key"]
for j, h in enumerate(heads, 1):
    cell(D, 7, j, h, bold=True, fill=GT if j <= 14 else DARK, color="FFFFFF", align=CC, size=8.5)
D.row_dimensions[7].height = 48
for r in range(R0, R1 + 1):
    i = r - R0 + 1
    cell(D, r, 1, i, align=CC, fill=LAV)
    for c in range(2, 15):
        cell(D, r, c, None, fill=INPUT, align=WT, size=8.5)
    C, E, F, G, H, I, J, K, Lc, M, N = (f"{x}{r}" for x in "CEFGHIJKLMN")
    full = f'COUNTA(E{r}:N{r})<10'
    f = {
        15: f'=IF(OR(C{r}="",E{r}=""),"",INDEX({LB},MATCH(E{r},{LA},0)))',
        16: f'=IF(O{r}="","",INDEX({fam("C")},MATCH(O{r},{LB},0)))',
        17: f'=IF(O{r}="","",INDEX({fam("D")},MATCH(O{r},{LB},0)))',
        18: f'=IF(OR(O{r}="",I{r}=""),"",IF(INDEX({DA_V},MATCH(I{r},{DA_O},0))>=INDEX({fam("E")},MATCH(O{r},{LB},0)),"Data ready","Data gap"))',
        19: f'=IF(O{r}="","",IF(P{r}<={CEIL},"Within current position","Beyond current position"))',
        20: f'=IF(O{r}="","",IF({full},"Complete inputs",IF(H{r}="{STANDARD[2][0]}","{VERDICTS[4][0]}",IF(Q{r}="No","{VERDICTS[1][0]}",'
            f'IF(R{r}="Data gap","{VERDICTS[2][0]}",IF(S{r}="Beyond current position","{VERDICTS[3][0]}","{VERDICTS[0][0]}"))))))',
        21: f'=IF(OR(T{r}="",T{r}="Complete inputs"),"",ROUND(0.3*INDEX({VO_V},MATCH(F{r},{VO_O},0))+0.25*IF(G{r}<20,1,IF(G{r}<80,2,IF(G{r}<160,3,IF(G{r}<480,4,5))))'
            f'+0.25*J{r}+0.2*K{r},2))',
        22: f'=IF(U{r}="","",ROUND(0.3*IF(R{r}="Data ready",5,IF(INDEX({DA_V},MATCH(I{r},{DA_O},0))=INDEX({fam("E")},MATCH(O{r},{LB},0))-1,3,1))'
            f'+0.25*INDEX({ST_V},MATCH(H{r},{ST_O},0))+0.25*(6-P{r})+0.2*IF(S{r}="Within current position",5,2),2))',
        23: f'=IF(U{r}="","",IF(Q{r}="No",IF(L{r}="Yes","Medium","Low"),IF(IF(L{r}="Yes",2,0)+IF(M{r}="Yes",1,0)+INDEX({AU_V},MATCH(N{r},{AU_O},0))+IF(K{r}=5,2,IF(K{r}>=4,1,0))>=4,"High",'
            f'IF(IF(L{r}="Yes",2,0)+IF(M{r}="Yes",1,0)+INDEX({AU_V},MATCH(N{r},{AU_O},0))+IF(K{r}=5,2,IF(K{r}>=4,1,0))>=2,"Medium","Low"))))',
        24: f'=IF(U{r}="","",IF(AND(U{r}>={THR},V{r}>={THR}),"Quick win",IF(U{r}>={THR},"Strategic bet",IF(V{r}>={THR},"Fill-in","Deprioritise"))))',
        25: f'=IF(U{r}="","",ROUND(U{r}*V{r}*IF(W{r}="High",0.85,1)*IF(T{r}="{VERDICTS[4][0]}",0.8,1),2))',
        26: f'=IF(Y{r}="","",RANK(Y{r},$Y${R0}:$Y${R1})+COUNTIF($Y${R0 - 1}:Y{r - 1},Y{r}))',
        27: f'=IF(AC{r}="","",INDEX(Lists!$AV$2:$AV$5,AC{r}))',
        28: f'=IF(Y{r}="","",ROUND(G{r}*12*INDEX({fam("F")},MATCH(O{r},{LB},0)),0))',
        29: f'=IF(Y{r}="","",IF(X{r}="Deprioritise",4,IF(T{r}="{VERDICTS[2][0]}",3,IF(OR(T{r}="{VERDICTS[4][0]}",T{r}="{VERDICTS[3][0]}",X{r}="Strategic bet"),2,1))))',
        30: f'=IF(Z{r}="","",AC{r}*100+Z{r})',
    }
    for c, fx in f.items():
        x = cell(D, r, c, fx, fill=None, align=CC if c not in (15, 19, 20, 27) else WC, size=8.5, bold=c in (20, 26))
        if c in (21, 22, 25): x.number_format = "0.00"
        if c == 28: x.number_format = "#,##0"
    D.row_dimensions[r].height = 40
rng = lambda col: f"{col}{R0}:{col}{R1}"
dv_list(D, f"={FN}", rng("B"))
dv_list(D, f"={LA}", rng("E"), "What is the work, essentially? This determines the solution family.")
dv_list(D, f"={VO_O}", rng("F")); dv_list(D, f"={ST_O}", rng("H")); dv_list(D, f"={DA_O}", rng("I"))
dv_list(D, '"1,2,3,4,5"', f"J{R0}:K{R1}"); dv_list(D, f"={YN}", f"L{R0}:M{R1}"); dv_list(D, f"={AU_O}", rng("N"))
dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="100000", allow_blank=True); D.add_data_validation(dv); dv.add(rng("G"))
verdict_cf(D, rng("T")); risk_cf(D, rng("W"))
D.conditional_formatting.add(rng("R"), CellIsRule(operator="equal", formula=['"Data gap"'], font=Font(color="C00000", bold=True)))
D.conditional_formatting.add(rng("S"), CellIsRule(operator="equal", formula=['"Beyond current position"'], font=Font(color="C55A11", bold=True)))
D.conditional_formatting.add(rng("X"), CellIsRule(operator="equal", formula=['"Quick win"'], fill=PatternFill("solid", fgColor="A9D08E")))
D.column_dimensions["AC"].hidden = True; D.column_dimensions["AD"].hidden = True
D.freeze_panes = "D8"; D.print_title_rows = "6:7"
D.print_area = f"A1:AB{R1}"
D.page_setup.orientation = "landscape"; D.page_setup.paperSize = D.PAPERSIZE_A3
D.page_setup.fitToWidth = 1; D.page_setup.fitToHeight = 0; D.sheet_properties.pageSetUpPr.fitToPage = True
if EXAMPLE:
    for k, (fi, pr, pain, ni, vi, hrs, si, di, sti, ci, rights, pers, ai) in enumerate(EXAMPLE_PROCESSES):
        r = R0 + k
        vals = [FUNCTIONS[fi], pr, pain, FAMILIES[ni][0], VOLUME[vi][0], hrs, STANDARD[si][0], DATA_AVAIL[di][0], sti, ci, rights, pers, AUTONOMY[ai][0]]
        for j, v in enumerate(vals, 2): D.cell(r, j).value = v
DG = "'3. AI Fit Diagnostic'!"
DR = lambda col: f"{DG}${col}${R0}:${col}${R1}"

# ================================================================== 4. Opportunity Map
M = wb.create_sheet("4. Opportunity Map", 4)
title(M, "4. AI Opportunity Map", "Where AI fits - and where automation is enough - by function, with the top-ranked opportunities. Calculated from sheets 1 and 3.", 11)
for c, w in zip("ABCDEFGHIJK", [5, 30, 11, 11, 11, 11, 11, 11, 13, 13, 22]): M.column_dimensions[c].width = w
r = 4
band(M, r, "A. ORGANISATIONAL AI POSITION", 11); r += 1
for k, f, fmt in [("Organisation", "='1. AI Position'!C4", None), ("AI position (archetype)", f"={POS_NAME}", None),
                  ("Data Position / Scale & Capacity", f'=IF({POS_SCORE_D}="","",TEXT({POS_SCORE_D},"0%")&"  /  "&TEXT({POS_SCORE_S},"0%"))', None),
                  ("What it means", f"='1. AI Position'!C{RES + 6}", None), ("Foundation flag", f"={POS_FLAG}", None)]:
    merged(M, r, 1, 2, k, bold=True, fill=LAV)
    merged(M, r, 3, 11, f, align=WT, bold=k.startswith("AI position"), fill=LAV2 if k.startswith("AI position") else None)
    M.row_dimensions[r].height = 30 if k in ("What it means", "Foundation flag") else 18; r += 1
r += 1; band(M, r, "B. VERDICT SUMMARY - what kind of solution fits the processes assessed", 11); r += 1
for j, h in enumerate(["", "Verdict", "Processes", "% of total", "Hours released / yr", "", "Meaning"], 1):
    if h: cell(M, r, j, h, bold=True, fill=GT, color="FFFFFF", align=CC)
M.merge_cells(start_row=r, start_column=7, end_row=r, end_column=11)
r += 1; VS = r
for i, (v, mean, nxt) in enumerate(VERDICTS):
    cell(M, r, 2, v, bold=True, fill=VFILL[v])
    cell(M, r, 3, f'=COUNTIF({DR("T")},B{r})', align=CC)
    cell(M, r, 4, f'=IF(SUM($C${VS}:$C${VS + 4})=0,"",C{r}/SUM($C${VS}:$C${VS + 4}))', align=CC, fmt="0%")
    cell(M, r, 5, f'=SUMIF({DR("T")},B{r},{DR("AB")})', align=CC, fmt="#,##0")
    merged(M, r, 6, 11, mean, align=WT, size=8.5)
    M.row_dimensions[r].height = 30; r += 1
cell(M, r, 2, "Total assessed", bold=True, fill=LAV); cell(M, r, 3, f"=SUM(C{VS}:C{VS + 4})", bold=True, align=CC, fill=LAV)
cell(M, r, 4, "", fill=LAV); cell(M, r, 5, f"=SUM(E{VS}:E{VS + 4})", bold=True, align=CC, fill=LAV, fmt="#,##0")
r += 1
merged(M, r, 2, 11, f'=IF(C{r - 1}=0,"",TEXT((C{VS}+C{VS + 3})/C{r - 1},"0%")&" of assessed processes are AI opportunities (now or with a partner); "&TEXT(C{VS + 1}/C{r - 1},"0%")&'
                    f'" are better served by automation without AI; "&TEXT((C{VS + 2}+C{VS + 4})/C{r - 1},"0%")&" need data or process foundations first.")',
       italic=True, align=WT, color=GT, bold=True)
M.row_dimensions[r].height = 30; r += 2
band(M, r, "C. WHERE AI FITS BY FUNCTION (number of processes)", 11); r += 1
fh = ["", "Function", "AI fit now", "Stretch", "AI later (data)", "Automate (no AI)", "Redesign first", "Total", "Hours / yr", "High-risk items", "Headline"]
for j, h in enumerate(fh, 1):
    if h: cell(M, r, j, h, bold=True, fill=GT, color="FFFFFF", align=CC, size=8.5)
M.row_dimensions[r].height = 30; r += 1; FS = r
vorder = [VERDICTS[0][0], VERDICTS[3][0], VERDICTS[2][0], VERDICTS[1][0], VERDICTS[4][0]]
for fn in FUNCTIONS:
    cell(M, r, 2, fn, bold=True, fill=LAV, size=8.5)
    for j, v in enumerate(vorder):
        cell(M, r, 3 + j, f'=COUNTIFS({DR("B")},$B{r},{DR("T")},"{v}")', align=CC)
    cell(M, r, 8, f"=SUM(C{r}:G{r})", align=CC, bold=True)
    cell(M, r, 9, f'=SUMIF({DR("B")},$B{r},{DR("AB")})', align=CC, fmt="#,##0")
    cell(M, r, 10, f'=COUNTIFS({DR("B")},$B{r},{DR("W")},"High")', align=CC)
    cell(M, r, 11, f'=IF(H{r}=0,"",IF(C{r}+D{r}>=MAX(E{r},F{r},G{r}),"AI opportunity",IF(F{r}>=MAX(E{r},G{r}),"Automation is enough",IF(E{r}>=G{r},"Fix data first","Redesign first"))))',
         align=CC, size=8.5, bold=True)
    r += 1
FE = r - 1
for j, col in enumerate("CDEFGHIJ"):
    cell(M, r, 3 + j, f"=SUM({col}{FS}:{col}{FE})", bold=True, fill=LAV, align=CC, fmt="#,##0" if col == "I" else None)
cell(M, r, 2, "Total", bold=True, fill=LAV); cell(M, r, 11, "", fill=LAV)
for col, v in zip("CDEFG", vorder):
    M.conditional_formatting.add(f"{col}{FS}:{col}{FE}", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor=VFILL[v])))
M.conditional_formatting.add(f"K{FS}:K{FE}", CellIsRule(operator="equal", formula=['"AI opportunity"'], fill=PatternFill("solid", fgColor="A9D08E")))
M.conditional_formatting.add(f"K{FS}:K{FE}", CellIsRule(operator="equal", formula=['"Automation is enough"'], fill=PatternFill("solid", fgColor="9BC2E6")))
r += 2
ch = BarChart(); ch.type = "bar"; ch.grouping = "stacked"; ch.overlap = 100; ch.title = "Processes by function and verdict"
ch.y_axis.title = None; ch.x_axis.title = None
data = Reference(M, min_col=3, max_col=7, min_row=FS - 1, max_row=FE); cats = Reference(M, min_col=2, min_row=FS, max_row=FE)
ch.add_data(data, titles_from_data=True); ch.set_categories(cats)
for s, v in zip(ch.series, vorder):
    s.graphicalProperties.solidFill = VFILL[v]; s.graphicalProperties.line.solidFill = VFILL[v]
ch.height = 9.5; ch.width = 25; ch.legend.position = "b"
ch.x_axis.scaling.orientation = "maxMin"; ch.x_axis.delete = False; ch.y_axis.delete = False; ch.y_axis.majorGridlines = None
M.add_chart(ch, f"B{r}")
r += 20
M.row_breaks.append(Break(id=r - 1))
band(M, r, "D. TOP 10 OPPORTUNITIES (risk-adjusted priority)", 11); r += 1
th = ["Rank", "Process / activity", "Function", "Solution family", "", "Verdict", "", "Value", "Feasibility", "Risk tier", "Horizon"]
for j, h in enumerate(th, 1):
    if h: cell(M, r, j, h, bold=True, fill=GT, color="FFFFFF", align=CC, size=8.5)
M.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5); M.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
M.row_dimensions[r].height = 24; r += 1; TS = r
for k in range(1, 11):
    mrow = f"MATCH({k},{DR('Z')},0)"
    cell(M, r, 1, k, bold=True, align=CC, fill=LAV)
    g_ = lambda col: f'=IFERROR(INDEX({DR(col)},{mrow}),"")'
    cell(M, r, 2, g_("C"), align=WT, bold=True, size=8.5); cell(M, r, 3, g_("B"), align=WT, size=8)
    merged(M, r, 4, 5, g_("O"), align=WT, size=8.5); merged(M, r, 6, 7, g_("T"), align=WT, size=8.5, bold=True)
    cell(M, r, 8, g_("U"), align=CC, fmt="0.00"); cell(M, r, 9, g_("V"), align=CC, fmt="0.00"); cell(M, r, 10, g_("W"), align=CC)
    cell(M, r, 11, g_("AA"), align=WT, size=8.5)
    M.row_dimensions[r].height = 36; r += 1
verdict_cf(M, f"F{TS}:F{r - 1}"); risk_cf(M, f"J{TS}:J{r - 1}")
r += 1
merged(M, r, 1, 11, "Value and feasibility scored 1-5 (sheet 3). Hours released are indicative (effort x typical reduction for the solution family) and must be validated in "
                    "the business case. Risk tiers are inherent (before controls) - see the Use-Case Canvas for minimum controls.", italic=True, size=8, color="595959", align=WT, border=False)
M.row_dimensions[r].height = 28
M.print_area = f"A1:K{r}"
M.page_setup.orientation = "portrait"; M.page_setup.fitToWidth = 1; M.page_setup.fitToHeight = 0; M.sheet_properties.pageSetUpPr.fitToPage = True

# ================================================================== 5. Roadmap
RM = wb.create_sheet("5. Roadmap", 5)
title(RM, "5. Opportunity Roadmap", "Opportunities placed automatically by horizon (ordered by priority rank), with the foundations they depend on. "
      "Up to 20 / 12 / 10 / 8 items per lane are shown; adjust the horizon by changing inputs on sheet 3.", 7)
for c, w in zip("ABCDEFG", [6, 38, 26, 30, 26, 11, 10]): RM.column_dimensions[c].width = w
r = 4
KEY = DR("AD")
for h, hname in enumerate(HORIZONS, 1):
    fill = [GT, GT2, TEAL, "7F7F7F"][h - 1]
    band(RM, r, hname, 7, fill=fill); r += 1
    for j, hd in enumerate(["Rank", "Process / activity", "Function", "Solution family", "Verdict", "Risk tier", "Hours / yr"], 1):
        cell(RM, r, j, hd, bold=True, fill=LAV, align=CC, size=8.5)
    r += 1; s0 = r
    for k in range(1, [20, 12, 10, 8][h - 1] + 1):
        key = f'IFERROR(SMALL({KEY},COUNTIF({KEY},"<"&{h * 100})+{k}),"")'
        cell(RM, r, 7, None)
        # helper in hidden col H: the key if it belongs to this lane
        RM.cell(r, 8, f'=IF({key}="","",IF({key}<{(h + 1) * 100},{key},""))')
        m = f"MATCH($H{r},{KEY},0)"
        for j, col in enumerate(["Z", "C", "B", "O", "T", "W", "AB"], 1):
            cell(RM, r, j, f'=IF($H{r}="","",INDEX({DR(col)},{m}))', align=CC if j in (1, 6, 7) else WT, size=8.5, bold=j in (1, 2),
                 fmt="#,##0" if j == 7 else None)
        RM.row_dimensions[r].height = 26
        r += 1
    verdict_cf(RM, f"E{s0}:E{r - 1}"); risk_cf(RM, f"F{s0}:F{r - 1}")
    r += 1
RM.column_dimensions["H"].hidden = True
band(RM, r, "ENABLING FOUNDATIONS (generated from the diagnostic and position screen)", 7, fill=DARK); r += 1
T = DR("T"); W = DR("W"); Q = DR("Q")
enablers = [
    ("Governance (always)", f'="Before H1 go-live: AI acceptable-use policy, AI register and approved-tools list; tier-based use-case risk assessment (GT AI Risk Assessment Framework). '
                            f'"&COUNTIF({W},"High")&" High-risk and "&COUNTIF({W},"Medium")&" Medium-risk items need documented assessments."'),
    ("Data foundations", f'=IF(COUNTIF({T},"{VERDICTS[2][0]}")=0,"No processes are blocked by data gaps.",COUNTIF({T},"{VERDICTS[2][0]}")&" process(es) are blocked by data gaps: assign data owners, '
                         f'digitise inputs and start recording outcomes now so they can move to AI in H3.")'),
    ("Process redesign", f'=IF(COUNTIF({T},"{VERDICTS[4][0]}")=0,"No processes need redesign before automation.",COUNTIF({T},"{VERDICTS[4][0]}")&" process(es) vary by person or team: '
                         f'standardise and document them first (Lean / process mapping), then re-diagnose.")'),
    ("Automation platform", f'=IF(COUNTIF({T},"{VERDICTS[1][0]}")=0,"","Automation is enough for "&COUNTIF({T},"{VERDICTS[1][0]}")&" process(es): select one workflow / RPA / BI platform and '
                            f'automate them as a programme rather than one-offs.")'),
    ("Capability & partners", f'=IF(COUNTIF({T},"{VERDICTS[3][0]}")=0,"All AI opportunities are within the current position.",COUNTIF({T},"{VERDICTS[3][0]}")&" opportunity(ies) exceed the '
                              f'current position: use vendor products or a delivery partner and build skills before scaling.")'),
    ("People", '="AI literacy for all users of AI tools and role-based training for process owners (EU AI Act Art. 4-aligned); change management for each roll-out."'),
    ("Readiness check", '="For governance and capability maturity, run the GT AI Readiness & Risk Assessment (AI Ready7-based) alongside this roadmap."'),
]
for k, f in enablers:
    cell(RM, r, 1, "", fill=LAV); cell(RM, r, 2, k, bold=True, fill=LAV)
    merged(RM, r, 3, 7, f, align=WT, size=8.5); RM.row_dimensions[r].height = 32; r += 1
RM.print_area = f"A1:G{r - 1}"
RM.page_setup.orientation = "portrait"; RM.page_setup.fitToWidth = 1; RM.page_setup.fitToHeight = 0; RM.sheet_properties.pageSetUpPr.fitToPage = True

# ================================================================== 6. Use-Case Canvas
CV = wb.create_sheet("6. Use-Case Canvas", 6)
title(CV, "6. AI Use-Case Canvas", "Select an opportunity number from the AI Fit Diagnostic (yellow cell). Calculated fields fill automatically; complete the yellow fields with the owner.", 4)
for c, w in zip("ABCD", [28, 44, 28, 44]): CV.column_dimensions[c].width = w
cell(CV, 4, 1, "Opportunity # (from sheet 3)", bold=True, fill=LAV); cell(CV, 4, 2, 2 if EXAMPLE else 1, fill=INPUT, bold=True, align=CC, size=12)
dv_list(CV, "=Lists!$AX$2:$AX$51", "B4")
cell(CV, 4, 3, "Priority rank", bold=True, fill=LAV)
m = f"MATCH($B$4,{DR('A')},0)"
gx = lambda col: f'=IFERROR(IF(INDEX({DR(col)},{m})="","",INDEX({DR(col)},{m})),"")'
cell(CV, 4, 4, gx("Z"), bold=True, align=CC, size=12)
r = 6
band(CV, r, "1. THE OPPORTUNITY", 4); r += 1
pairs = [("Process / activity", gx("C"), "Function", gx("B")),
         ("Pain point / opportunity", gx("D"), "Nature of the task", gx("E")),
         ("Volume", gx("F"), "Effort today (hrs / month)", gx("G"))]
for a, fa, b, fb in pairs:
    cell(CV, r, 1, a, bold=True, fill=LAV); cell(CV, r, 2, fa, align=WC); cell(CV, r, 3, b, bold=True, fill=LAV); cell(CV, r, 4, fb, align=WC)
    CV.row_dimensions[r].height = 34; r += 1
band(CV, r, "2. THE FIT", 4); r += 1
fam_m = f"MATCH(INDEX({DR('O')},{m}),{LB},0)"
pairs = [("VERDICT", gx("T"), "Solution family", gx("O")),
         ("Data gate", gx("R"), "Organisational fit", gx("S")),
         ("Value / Feasibility", f'=IFERROR(TEXT(INDEX({DR("U")},{m}),"0.00")&" / "&TEXT(INDEX({DR("V")},{m}),"0.00")&"  ("&INDEX({DR("X")},{m})&")","")',
          "Horizon", gx("AA")),
         ("Typical solutions", f'=IFERROR(INDEX({fam("G")},{fam_m}),"")', "Design notes", f'=IFERROR(INDEX({fam("I")},{fam_m}),"")')]
for a, fa, b, fb in pairs:
    cell(CV, r, 1, a, bold=True, fill=LAV); cell(CV, r, 2, fa, align=WC, bold=a == "VERDICT"); cell(CV, r, 3, b, bold=True, fill=LAV); cell(CV, r, 4, fb, align=WC)
    CV.row_dimensions[r].height = 34 if a != "Typical solutions" else 58; r += 1
VR = r - 4
verdict_cf(CV, f"B{VR}")
cell(CV, r, 1, "What this verdict means", bold=True, fill=LAV)
merged(CV, r, 2, 4, f'=IFERROR(INDEX(Lists!$AS$2:$AS$6,MATCH(B{VR},{VER},0)),"")', align=WC); CV.row_dimensions[r].height = 30; r += 1
band(CV, r, "3. RISK & MINIMUM CONTROLS", 4); r += 1
cell(CV, r, 1, "Inherent risk tier", bold=True, fill=LAV); cell(CV, r, 2, gx("W"), bold=True, align=CC); risk_cf(CV, f"B{r}")
cell(CV, r, 3, "Personal data / affects individuals", bold=True, fill=LAV)
cell(CV, r, 4, f'=IFERROR(INDEX({DR("M")},{m})&" / "&INDEX({DR("L")},{m}),"")', align=WC); RT = r; r += 1
cell(CV, r, 1, "Minimum controls before go-live", bold=True, fill=LAV)
merged(CV, r, 2, 4, f'=IFERROR(INDEX(Lists!$AP$2:$AP$4,MATCH(B{RT},Lists!$AO$2:$AO$4,0)),"")', align=WC); CV.row_dimensions[r].height = 66; r += 1
cell(CV, r, 1, "Regulatory screen", bold=True, fill=LAV)
merged(CV, r, 2, 4, f'=IF(B{RT}="","",IF(INDEX({DR("L")},{m})="Yes","Decisions about individuals: check EU AI Act Annex III high-risk categories (e.g., credit, employment, essential services) and Art. 5 prohibitions; '
                    f'PDPL lawful basis and DPIA; CBB requirements if a licensee.",IF(INDEX({DR("M")},{m})="Yes","Personal / confidential data: PDPL lawful basis, minimisation, cross-border transfer and vendor terms '
                    f'(no training on your data); client confidentiality.","No specific regulatory trigger identified - confirm with Compliance.")))', align=WC)
CV.row_dimensions[r].height = 44; r += 1
band(CV, r, "4. MEASURING SUCCESS", 4); r += 1
cell(CV, r, 1, "Suggested KPIs", bold=True, fill=LAV); merged(CV, r, 2, 4, f'=IFERROR(INDEX({fam("H")},{fam_m}),"")', align=WC); r += 1
cell(CV, r, 1, "Indicative hours released / yr", bold=True, fill=LAV); cell(CV, r, 2, gx("AB"), align=CC, fmt="#,##0")
cell(CV, r, 3, "Target / success criterion", bold=True, fill=LAV); cell(CV, r, 4, "", fill=INPUT); r += 1
band(CV, r, "5. NEXT STEP & OWNERSHIP", 4); r += 1
cell(CV, r, 1, "Recommended next step", bold=True, fill=LAV)
merged(CV, r, 2, 4, f'=IFERROR(INDEX(Lists!$AT$2:$AT$6,MATCH(B{VR},{VER},0)),"")', align=WC); CV.row_dimensions[r].height = 34; NX = r; r += 1
for a, b in [("Business owner", "Executive sponsor"), ("Budget / effort estimate", "Decision date")]:
    cell(CV, r, 1, a, bold=True, fill=LAV); cell(CV, r, 2, "", fill=INPUT); cell(CV, r, 3, b, bold=True, fill=LAV); cell(CV, r, 4, "", fill=INPUT); r += 1
cell(CV, r, 1, "Decision", bold=True, fill=LAV); merged(CV, r, 2, 4, "", fill=INPUT); dv_list(CV, '"Proceed to pilot,Proceed to automation,Fix foundations first,Defer,Reject"', f"B{r}"); r += 1
CV.print_area = f"A1:D{r - 1}"
CV.page_setup.orientation = "portrait"; CV.page_setup.fitToWidth = 1; CV.page_setup.fitToHeight = 1; CV.sheet_properties.pageSetUpPr.fitToPage = True

# ================================================================== Reference
RF = wb.create_sheet("Reference", 7)
title(RF, "Reference - solution families, archetypes, verdicts and standards", "The logic behind the toolkit, for facilitators and for the client's appendix.", 6)
for c, w in zip("ABCDEF", [40, 30, 8, 9, 14, 60]): RF.column_dimensions[c].width = w
r = 4; band(RF, r, "SOLUTION FAMILIES (task nature -> solution)", 6); r += 1
for j, h in enumerate(["Nature of the task", "Solution family", "Level", "AI?", "Min. data", "Typical solutions / design notes"], 1):
    cell(RF, r, j, h, bold=True, fill=GT, color="FFFFFF", align=CC)
r += 1
for f in FAMILIES:
    nat, fam_, lvl, ai, dreq, pot, tools, kpi, note = f
    cell(RF, r, 1, nat, align=WT); cell(RF, r, 2, fam_, align=WT, bold=True); cell(RF, r, 3, f"{lvl}", align=CC); cell(RF, r, 4, ai, align=CC)
    cell(RF, r, 5, DATA_AVAIL[dreq][0], align=WT, size=8); cell(RF, r, 6, f"{tools}.  {note}", align=WT, size=8.5)
    RF.row_dimensions[r].height = 52; r += 1
r += 1; band(RF, r, "SOLUTION LEVELS", 6); r += 1
for lvl, name in LEVELS:
    cell(RF, r, 1, f"Level {lvl}", bold=True, fill=LAV); merged(RF, r, 2, 6, name); r += 1
r += 1; band(RF, r, "AI POSITION ARCHETYPES", 6); r += 1
for code, name, ceil, head, desc in ARCHETYPES:
    cell(RF, r, 1, f"{code}. {name}", bold=True, fill=LAV, align=WT); cell(RF, r, 2, f"Ceiling: level {ceil}", align=CC)
    merged(RF, r, 3, 6, f"{head} {desc}", align=WT, size=8.5); RF.row_dimensions[r].height = 58; r += 1
r += 1; band(RF, r, "VERDICTS", 6); r += 1
for v, mean, nxt in VERDICTS:
    cell(RF, r, 1, v, bold=True, fill=VFILL[v]); merged(RF, r, 2, 6, f"{mean}  Next step: {nxt}", align=WT, size=8.5); RF.row_dimensions[r].height = 40; r += 1
r += 1; band(RF, r, "MINIMUM CONTROLS BY RISK TIER", 6); r += 1
for t, c in CONTROLS:
    cell(RF, r, 1, t, bold=True, fill=RFILL[t]); merged(RF, r, 2, 6, c, align=WT, size=8.5); RF.row_dimensions[r].height = 44; r += 1
r += 1; band(RF, r, "STANDARDS AND FRAMEWORKS BEHIND EACH STEP", 6); r += 1
for j, h in enumerate(["Standard / framework", "", "Step", "", "", "How it is used"], 1):
    if h: cell(RF, r, j, h, bold=True, fill=GT, color="FFFFFF", align=CC)
RF.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2); RF.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5); r += 1
for step, std, use in STANDARDS:
    merged(RF, r, 1, 2, std, bold=True, align=WT, size=8.5); merged(RF, r, 3, 5, step, align=CC, size=8.5); cell(RF, r, 6, use, align=WT, size=8.5)
    RF.row_dimensions[r].height = 40; r += 1
RF.print_area = f"A1:F{r - 1}"
RF.page_setup.orientation = "portrait"; RF.page_setup.fitToWidth = 1; RF.page_setup.fitToHeight = 0; RF.sheet_properties.pageSetUpPr.fitToPage = True

# ------------------------------------------------------------------ finish
for ws in wb.worksheets:
    if ws.title != "Lists":
        ws.oddFooter.left.text = "GT AI Opportunity Discovery Toolkit"; ws.oddFooter.left.size = 8
        ws.oddFooter.right.text = "Page &P of &N"; ws.oddFooter.right.size = 8
wb.active = 1 if not EXAMPLE else 4
printfix.harden(wb, keep_a3=("3. AI Fit Diagnostic",))
wb.save(OUT)
print("saved", OUT)
