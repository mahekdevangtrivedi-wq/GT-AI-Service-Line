# -*- coding: utf-8 -*-
"""GT AI Readiness & Risk Assessment Tool (Excel) with auto-generated printable report.
Usage: build_readiness_tool.py OUT.xlsx [short|detailed] [sample]
  short    - free assessment: summary report only (scores, domain radar, observations, regulatory view)
  detailed - full report on request: findings, risks and recommendations, AI privacy & data-quality section,
             priorities, roadmap, action plan"""
import sys, os, math
import openpyxl
import printfix
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import RadarChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.worksheet.pagebreak import Break, RowBreak
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Protection
from openpyxl.formatting.rule import DataBarRule
sys.path.insert(0, os.path.dirname(__file__))
from tool_data import *

OUT = sys.argv[1]
SAMPLE = "sample" in sys.argv[2:]
VARIANT = "short" if "short" in sys.argv[2:] else "detailed"
DET = VARIANT == "detailed"
import datetime
from openpyxl.workbook.defined_name import DefinedName

GT = "4F2D7F"; GT2 = "7B5BA6"; LAV = "EDE7F6"; LAV2 = "D9CCEB"; INPUT = "FFF9E6"; GREY = "F2F2F2"; DARK = "3B2160"
RAGF = {"Low": "C6E0B4", "Moderate": "FFE699", "Medium": "FFE699", "High": "F4B183", "Critical": "E06666", "Very high": "E06666"}
MATF = {"Partial": "F8CBAD", "Informed": "FFE699", "Repeatable": "C6E0B4", "Adaptive": "9BC2E6"}
thin = Side(style="thin", color="C9C9C9")
B = Border(left=thin, right=thin, top=thin, bottom=thin)
WT = Alignment(wrap_text=True, vertical="top"); WC = Alignment(wrap_text=True, vertical="center")
CC = Alignment(wrap_text=True, vertical="center", horizontal="center")
NB = Border()

wb = openpyxl.Workbook()
ALLQ = CONTEXT + READINESS
NQ = len(ALLQ)  # 26 (4 inherent-risk + 22 readiness)


def fill(c, col): c.fill = PatternFill("solid", fgColor=col)


def style(c, size=10, bold=False, color="000000", fillc=None, align=WT, border=True, italic=False):
    c.font = Font(name="Calibri", size=size, bold=bold, color=color, italic=italic)
    c.alignment = align
    if border: c.border = B
    if fillc: fill(c, fillc)


def merge(ws, r1, c1, r2, c2, value=None, **kw):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    c = ws.cell(r1, c1)
    if value is not None: c.value = value
    style(c, **kw)
    if kw.get("border", True):
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                ws.cell(rr, cc).border = B
    return c


def lines(text, chars):
    return sum(max(1, math.ceil(len(t) / chars)) for t in str(text).split("\n"))


# =====================================================================  QBank (hidden)
qb = wb.active; qb.title = "QBank"
for i, h in enumerate(["QID", "Option", "Value", "Status label", "Playback / next step", "Level"], 1): qb.cell(1, i, h)
QROW = {}
for k, q in enumerate(ALLQ):
    a = 2 + k * 5; QROW[q["id"]] = (a, a + 4)
    for j in range(5):
        r = a + j
        qb.cell(r, 1, q["id"])
        if k < 4:
            txt, imp = q["opts"][j]
            qb.cell(r, 2, txt); qb.cell(r, 3, j)
            qb.cell(r, 4, ["Low", "Low", "Medium", "High", "Very high"][j]); qb.cell(r, 5, imp); qb.cell(r, 6, f"Risk level {j} of 4")
        else:
            qb.cell(r, 2, q["opts"][j]); qb.cell(r, 3, j * 0.25)
            qb.cell(r, 4, ["Not met", "Minimal", "Partially met", "Largely met", "Fully met"][j]); qb.cell(r, 5, q["recs"][j]); qb.cell(r, 6, LEVELS[j])
qb.sheet_state = "hidden"

# =====================================================================  Lists (hidden)
ls = wb.create_sheet("Lists")
LISTS = {"A": ["Yes", "No"],
         "B": ["Banking & financial services", "Insurance", "Government & public sector", "Telecommunications", "Healthcare", "Hospitality & tourism",
               "Professional services", "Manufacturing & industry", "Retail & consumer", "Energy & utilities", "Education", "Other"],
         "C": ["1-50 employees", "51-250 employees", "251-1,000 employees", "1,001-5,000 employees", "More than 5,000 employees"],
         "D": ["Whole organisation", "Business unit / function", "Specific AI system or use case"],
         "E": ["Deployer only (uses AI built by others)", "Provider (develops / sells AI systems)", "Both provider and deployer"],
         "F": ["Not started", "Planned", "In progress", "Completed", "Deferred"]}
for col, vals in LISTS.items():
    for i, v in enumerate(vals, 1): ls[f"{col}{i}"] = v
ls["H1"] = "Pasted responses (ID | organisation | name)"
for i in range(2, 502):
    ls[f"H{i}"] = f'=IF(\'0. Paste Response\'!A{i}="","",\'0. Paste Response\'!A{i}&" | "&\'0. Paste Response\'!F{i}&" | "&\'0. Paste Response\'!G{i})'
ls["I1"] = "=IFERROR(MATCH('2. Assessment'!$D$3,$H$2:$H$501,0)+1,0)"
wb.defined_names["ResponseList"] = DefinedName("ResponseList", attr_text="OFFSET(Lists!$H$2,0,0,MAX(1,COUNTA('0. Paste Response'!$A$2:$A$501)),1)")
ls.sheet_state = "hidden"


def dv(ws, rng, src):
    d = DataValidation(type="list", formula1=src, allow_blank=True, showErrorMessage=True, errorTitle="Invalid entry", error="Please choose a value from the list.")
    ws.add_data_validation(d); d.add(rng)


# =====================================================================  0. Paste Response (Forms export, one row per response)
from openpyxl.utils import get_column_letter as _gcl
ps = wb.create_sheet("0. Paste Response")
PASTE_SHORT = {"A1": "AI use profile", "A2": "Decision impact", "A3": "Data sensitivity", "A4": "Oversight & reversibility", "Q5": "Governance structure",
               "Q6": "AI policy & compliance", "Q7": "Responsible-AI guardrails", "Q8": "Incidents & assurance", "Q9": "Risk framework", "Q10": "Risk appetite & KRIs",
               "Q11": "Inventory & impact assessment", "Q12": "Third-party AI risk", "Q13": "Securing AI systems", "Q14": "AI-enabled threats",
               "Q15": "Training & literacy", "Q16": "Strategy & ownership", "Q17": "Use-case portfolio", "Q18": "Data governance",
               "Q19": "Architecture & monitoring", "Q20": "Model risk & validation", "Q21": "Privacy governance & PDPL", "Q22": "Data minimisation & retention",
               "Q23": "Privacy impact & transfers", "Q24": "PIMS (ISO 27701 / 27001)", "Q25": "Data quality", "Q26": "Data provenance & bias"}
PASTE_HEADS = ["ID", "Start time", "Completion time", "Email (Forms)", "Name (Forms)",
               "P1.1 Organisation name", "P1.2 Your name", "P1.3 Email address", "P1.4 Sector", "P1.5 Organisation size", "P1.6 Scope of assessment",
               "P1.7 Scope description", "P1.8 Role / title", "P1.9 Personal data in Bahrain?", "P1.10 Other GCC markets?", "P1.11 Provider or deployer"]
PASTE_COL = {}   # question id -> paste column letter
for q in ALLQ:
    PASTE_HEADS.append(f'{q["id"]} {PASTE_SHORT[q["id"]]}')
    PASTE_COL[q["id"]] = _gcl(len(PASTE_HEADS))
PASTE_HEADS += ["P3.1 AI training", "P3.2 AI / automation review", "P3.3 AI risk assessment", "P3.4 AI strategy & roadmap", "P3.5 Detailed report"]
for j, h in enumerate(PASTE_HEADS, 1):
    c = ps.cell(1, j, h); style(c, size=9, bold=True, color="FFFFFF", fillc=GT, align=CC)
    ps.column_dimensions[_gcl(j)].width = 14 if j <= 5 else 24
ps.row_dimensions[1].height = 48; ps.freeze_panes = "F2"; ps.sheet_properties.tabColor = "00A7B5"

# =====================================================================  1. Profile
pf = wb.create_sheet("1. Profile")
pf.sheet_view.showGridLines = False
for col, w in zip("ABCDE", [2, 46, 62, 2, 40]): pf.column_dimensions[col].width = w
merge(pf, 1, 2, 1, 3, "1. ORGANISATION PROFILE", size=16, bold=True, color="FFFFFF", fillc=GT, align=Alignment(vertical="center", indent=1), border=False)
pf.row_dimensions[1].height = 34
merge(pf, 2, 2, 2, 3, "Complete the yellow cells. The profile personalises the report and determines which regulatory considerations apply. Takes about 2 minutes.", size=10, italic=True, color=GT, border=False)
pf.row_dimensions[2].height = 28
# label, sample, list source, paste column (None = enter manually)
PROFILE = [("Organisation name", "Grant Thornton Bahrain", None, "F"),
           ("Sector", "Professional services", "=Lists!$B$1:$B$12", "I"),
           ("Organisation size", "251-1,000 employees", "=Lists!$C$1:$C$5", "J"),
           ("Scope of assessment", "Whole organisation", "=Lists!$D$1:$D$3", "K"),
           ("Scope description (optional)", "Firm-wide AI adoption across Audit, Tax, Advisory and internal functions", None, "L"),
           ("Completed by", "AI Service Line Lead", None, "G"),
           ("Role / title", "Director, Risk Advisory", None, "M"),
           ("Email address (to send yourself the report)", "ai.serviceline@example.com", None, "H"),
           ("Assessment date (dd/mm/yyyy)", None, None, "C"),
           ("Grant Thornton contact (optional)", "", None, None),
           (None, None, None, None),
           ("REGULATORY CONTEXT", None, None, None),
           ("Is the organisation licensed by the Central Bank of Bahrain (CBB)?", "No", "=Lists!$A$1:$A$2", None),
           ("Do AI systems process personal data of individuals in Bahrain?", "Yes", "=Lists!$A$1:$A$2", "N"),
           ("Are AI-enabled products / services offered in the EU, or AI outputs used in the EU?", "Yes", "=Lists!$A$1:$A$2", None),
           ("Does the organisation operate in KSA, UAE or other GCC markets?", "Yes", "=Lists!$A$1:$A$2", "O"),
           ("Is the organisation a government entity or public-sector body?", "No", "=Lists!$A$1:$A$2", None),
           ("Does the organisation develop AI (provider) or only use it (deployer)?", "Deployer only (uses AI built by others)", "=Lists!$E$1:$E$3", "P")]
P = {}
PSAMPLE = {}   # paste column -> sample value (for the SAMPLE workbook's paste row)
r = 4
for label, sample, src, pcol in PROFILE:
    if label is None:
        r += 1; continue
    if label == "REGULATORY CONTEXT":
        merge(pf, r, 2, r, 3, label, size=11, bold=True, color="FFFFFF", fillc=GT2); r += 1; continue
    style(pf.cell(r, 2, label), bold=True, fillc=LAV, align=WC)
    c = pf.cell(r, 3); style(c, fillc=INPUT, align=WC)
    if pcol == "C":   # assessment date = Forms completion date
        c.value = '=IF(Lists!$I$1=0,"",IF(INDEX(\'0. Paste Response\'!$C:$C,Lists!$I$1)="","",IFERROR(INT(INDEX(\'0. Paste Response\'!$C:$C,Lists!$I$1)),INDEX(\'0. Paste Response\'!$C:$C,Lists!$I$1))))'
        c.number_format = "dd/mm/yyyy"
    elif pcol:
        c.value = f'=IF(Lists!$I$1=0,"",INDEX(\'0. Paste Response\'!${pcol}:${pcol},Lists!$I$1)&"")'
        PSAMPLE[pcol] = sample
    elif SAMPLE:
        c.value = sample
    if src: dv(pf, f"C{r}", src)
    pf.row_dimensions[r].height = 24
    P[label] = f"'1. Profile'!$C${r}"
    r += 1
style(pf.cell(4, 5, "Tip"), bold=True, color=GT, border=False)
tip = pf.cell(5, 5, "Scope the assessment to the whole organisation for a readiness baseline, or to a single AI system to support a use-case decision. For use-case level risk assessment use the GT AI Risk Assessment Toolkit.")
style(tip, size=9, italic=True, color="595959", border=False)
pf.merge_cells("E5:E9")
ORG = P["Organisation name"]; DATE = P["Assessment date (dd/mm/yyyy)"]; EMAIL = P["Email address (to send yourself the report)"]
CBB = P["Is the organisation licensed by the Central Bank of Bahrain (CBB)?"]; PDPL = P["Do AI systems process personal data of individuals in Bahrain?"]
EU = P["Are AI-enabled products / services offered in the EU, or AI outputs used in the EU?"]; GCC = P["Does the organisation operate in KSA, UAE or other GCC markets?"]
GOV = P["Is the organisation a government entity or public-sector body?"]; ROLE = P["Does the organisation develop AI (provider) or only use it (deployer)?"]

# =====================================================================  2. Assessment
asx = wb.create_sheet("2. Assessment")
asx.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGHIJK", [2, 6, 17, 52, 56, 9, 13, 54, 17, 2, 6]): asx.column_dimensions[col].width = w
merge(asx, 1, 2, 1, 9, f"2. ASSESSMENT  -  {NQ} questions  |  approx. 25 minutes", size=16, bold=True, color="FFFFFF", fillc=GT, align=Alignment(vertical="center", indent=1), border=False)
asx.row_dimensions[1].height = 34
merge(asx, 2, 2, 2, 9, "Choose the response to score in cell D3. Each answer in column E then loads from '0. Paste Response' (you can still overwrite a cell manually). Part A profiles AI use and inherent risk; "
      "Part B measures readiness across 6 domains, including AI privacy & data quality (Bahrain PDPL, ISO/IEC 27701 with ISO/IEC 27001). Maturity uses a four-tier scale modelled on the NIST implementation tiers; "
      "questions are aligned to the NIST AI RMF. Results update live below and in the '3. Report' sheet.", size=9.5, italic=True, color=DARK, border=False)
asx.row_dimensions[2].height = 44
style(asx.cell(3, 2, "Response to score"), size=9.5, bold=True, color=GT, fillc=LAV, align=WC); asx.merge_cells("B3:C3")
sel = asx.cell(3, 4, "1 | Grant Thornton Bahrain | AI Service Line Lead" if SAMPLE else None); style(sel, size=10, bold=True, fillc=INPUT, align=WC)
sel.protection = Protection(locked=False)
dsel = DataValidation(type="list", formula1="ResponseList", allow_blank=True, showErrorMessage=False); asx.add_data_validation(dsel); dsel.add("D3")
merge(asx, 3, 5, 3, 9, "<- pick a response pasted in '0. Paste Response' (ID | organisation | name). Profile, Assessment and Report update automatically.", size=9, italic=True, color="595959", border=False, align=WC)
asx.row_dimensions[3].height = 26
KPI_ROW = 4
HDR = 8
heads = ["#", "Area", "Question", "Your response (select from list)", "Score", "Status", "What this means / recommended next step", "Requirement type"]
for i, h in enumerate(heads, 2):
    c = asx.cell(HDR, i, h); style(c, bold=True, color="FFFFFF", fillc=GT, align=CC)
asx.row_dimensions[HDR].height = 30
QR = {}  # qid -> row
SAMPLE_ANS = {"A1": 2, "A2": 2, "A3": 2, "A4": 1, "Q5": 2, "Q6": 3, "Q7": 2, "Q8": 3, "Q9": 1, "Q10": 2, "Q11": 2, "Q12": 2, "Q13": 2, "Q14": 3,
              "Q15": 3, "Q16": 2, "Q17": 2, "Q18": 3, "Q19": 2, "Q20": 2, "Q21": 2, "Q22": 2, "Q23": 1, "Q24": 1, "Q25": 2, "Q26": 1}
r = HDR + 1
for part, qs in [("PART A  -  AI USE & INHERENT RISK PROFILE  (how much could go wrong?)", CONTEXT),
                 ("PART B  -  AI READINESS & CONTROL MATURITY  (6 domains: how well is it managed?)", READINESS)]:
    merge(asx, r, 2, r, 9, part, size=11, bold=True, color="FFFFFF", fillc=GT2, align=Alignment(vertical="center", indent=1))
    asx.row_dimensions[r].height = 22
    r += 1
    for q in qs:
        QR[q["id"]] = r
        a, b = QROW[q["id"]]
        is_ctx = q["id"].startswith("A")
        style(asx.cell(r, 2, q["id"]), bold=True, align=CC, fillc=LAV)
        style(asx.cell(r, 3, q["topic"] if is_ctx else f'{q["pillar"]}\n\n{q["topic"]}'), size=9, bold=True, color=GT, fillc=LAV)
        rt = CellRichText(TextBlock(InlineFont(b=True, sz=10), q["q"]), TextBlock(InlineFont(sz=8.5, i=True, color="6B6B6B"), "\n\n" + q["guide"]))
        c = asx.cell(r, 4); c.value = rt; c.alignment = WT; c.border = B
        ans = asx.cell(r, 5)
        pc = PASTE_COL[q["id"]]
        ans.value = f'=IF(Lists!$I$1=0,"",INDEX(\'0. Paste Response\'!${pc}:${pc},Lists!$I$1)&"")'
        style(ans, size=10, fillc=INPUT, align=WC)
        dv(asx, f"E{r}", f"=QBank!$B${a}:$B${b}")
        asx.cell(r, 11, f'=IF(E{r}="","",IFERROR(MATCH(E{r},QBank!$B${a}:$B${b},0),""))')
        asx.cell(r, 6, f'=IF(K{r}="","",INDEX(QBank!$C${a}:$C${b},K{r}))')
        style(asx.cell(r, 6), bold=True, align=CC)
        asx.cell(r, 6).number_format = '0" / 4"' if is_ctx else "0%"
        asx.cell(r, 7, f'=IF(K{r}="","Not answered",INDEX(QBank!$D${a}:$D${b},K{r}))'); style(asx.cell(r, 7), bold=True, align=CC, size=9)
        asx.cell(r, 8, f'=IF(K{r}="","",INDEX(QBank!$E${a}:$E${b},K{r}))'); style(asx.cell(r, 8), size=9)
        style(asx.cell(r, 9, "Inherent-risk driver" if is_ctx else q["type"]), size=9, align=WC)
        asx.row_dimensions[r].height = 96 if not is_ctx else 84
        if SAMPLE: PSAMPLE[pc] = qb.cell(a + SAMPLE_ANS[q["id"]], 2).value
        r += 1
LASTQ = r - 1
asx.column_dimensions["K"].hidden = True
# status colouring
ctx_rng = f"G{QR['A1']}:G{QR['A4']}"; rd_rng = f"G{QR['Q5']}:G{LASTQ}"
for lab, col in [("Low", "C6E0B4"), ("Medium", "FFE699"), ("High", "F4B183"), ("Very high", "E06666")]:
    asx.conditional_formatting.add(ctx_rng, CellIsRule(operator="equal", formula=[f'"{lab}"'], fill=PatternFill("solid", fgColor=col)))
for lab, col in [("Not met", "E06666"), ("Minimal", "F4B183"), ("Partially met", "FFE699"), ("Largely met", "C6E0B4"), ("Fully met", "9BC2E6")]:
    asx.conditional_formatting.add(rd_rng, CellIsRule(operator="equal", formula=[f'"{lab}"'], fill=PatternFill("solid", fgColor=col)))
asx.conditional_formatting.add(f"G{QR['A1']}:G{LASTQ}", CellIsRule(operator="equal", formula=['"Not answered"'], font=Font(color="A6A6A6", italic=True)))
asx.freeze_panes = f"A{HDR+1}"

# sample paste row (illustrative responses) + the paste sheet's own formatting
if SAMPLE:
    row = {"A": 1, "B": datetime.datetime(2026, 11, 10, 9, 0), "C": datetime.datetime(2026, 11, 10, 9, 24), "D": "anonymous"}
    row.update(PSAMPLE)
    for col, v in row.items():
        if v is not None: ps[f"{col}2"] = v
    ps["B2"].number_format = "dd/mm/yyyy hh:mm"; ps["C2"].number_format = "dd/mm/yyyy hh:mm"


# =====================================================================  Calc (hidden)
cl = wb.create_sheet("Calc")
cl["A1"] = "Per-question calculations"
hdrs = ["QID", "Topic", "Pillar", "Answered", "Value", "Weight/Mult", "Gap", "Priority", "Horizon", "Horizon rank", "Sort key", "Recommendation", "(unused)", "Type", "Response", "Status", "Risk"]
for i, h in enumerate(hdrs, 1): cl.cell(199, i, h)
CR = {}
TIER = "Calc!$C$32"
for k, q in enumerate(ALLQ):
    rr = 200 + k; CR[q["id"]] = rr
    ar = QR[q["id"]]
    cl.cell(rr, 1, q["id"]); cl.cell(rr, 2, q["topic"]); cl.cell(rr, 3, q.get("pillar", "Context"))
    cl.cell(rr, 4, f"=IF('2. Assessment'!K{ar}=\"\",0,1)")
    cl.cell(rr, 5, f"=IF(D{rr}=0,\"\",'2. Assessment'!F{ar})")
    cl.cell(rr, 15, f"='2. Assessment'!E{ar}")
    cl.cell(rr, 16, f"='2. Assessment'!G{ar}")
    if k >= 4:
        mult = "1.5" if ("Mandatory" in q["type"] or "Compliance" in q["type"]) else "1"
        extra = f'+IF(AND(OR({TIER}="High",{TIER}="Critical"),OR(A{rr}="Q7",A{rr}="Q9",A{rr}="Q11",A{rr}="Q19",A{rr}="Q23")),0.5,0)'
        cl.cell(rr, 6, f"={mult}{extra}")
        cl.cell(rr, 7, f'=IF(E{rr}="","",1-E{rr})')
        cl.cell(rr, 8, f'=IF(E{rr}="",-1,G{rr}*F{rr}*100+({NQ}-{k})/100)')
        cl.cell(rr, 9, f'=IF(E{rr}="","",IF(E{rr}<=0.25,"Immediate (0-3 months)",IF(AND(E{rr}=0.5,F{rr}>=1.5),"Immediate (0-3 months)",IF(E{rr}<=0.5,"Short term (3-6 months)",IF(E{rr}<=0.75,"Medium term (6-12 months)","Sustain")))))')
        cl.cell(rr, 10, f'=IF(I{rr}="",0,IF(LEFT(I{rr},3)="Imm",4,IF(LEFT(I{rr},3)="Sho",3,IF(LEFT(I{rr},3)="Med",2,1))))')
        cl.cell(rr, 11, f'=IF(E{rr}="",-1,J{rr}*1000+H{rr})')
        cl.cell(rr, 12, f"='2. Assessment'!H{QR[q['id']]}")
        cl.cell(rr, 14, q["type"]); cl.cell(rr, 17, q["risk"])
QF, QL = 200, 200 + NQ - 1
RF, RL = 204, QL   # readiness rows
# completion
cl["A20"] = "Answered"; cl["B20"] = f"=SUM(D{QF}:D{QL})"
cl["A21"] = "Complete"; cl["B21"] = f"=IF(B20={NQ},1,0)"
# inherent
cl["A25"] = "Inherent risk"
for i, q in enumerate(CONTEXT):
    rr = CR[q["id"]]
    cl.cell(26 + i, 1, q["id"]); cl.cell(26 + i, 2, f'=IF(E{rr}="","",E{rr})'); cl.cell(26 + i, 3, q["weight"])
cl["A30"] = "Inherent %"; cl["C30"] = f'=IF(COUNT(B26:B29)<4,"",SUMPRODUCT(B26:B29,C26:C29)/(4*SUM(C26:C29)))'
cl["A31"] = "Score tier"; cl["C31"] = '=IF(C30="","",IF(C30>=0.75,"Critical",IF(C30>=0.5,"High",IF(C30>=0.25,"Medium","Low"))))'
cl["A32"] = "Final tier"
cl["C32"] = '=IF(C31="","",IF(AND(B27=4,B29>=3),"Critical",IF(AND(B27=4,OR(C31="Low",C31="Medium")),"High",IF(AND(B28=4,C31="Low"),"Medium",C31))))'
cl["A33"] = "Override note"
cl["C33"] = '=IF(C31="","",IF(AND(B27=4,B29>=3),"Override: significant-effect decisions with limited oversight -> Critical",IF(AND(B27=4,OR(C31="Low",C31="Medium")),"Override: significant-effect decisions -> minimum High",IF(AND(B28=4,C31="Low"),"Override: sensitive data -> minimum Medium","No override applied"))))'
# pillars
cl["A36"] = "Pillar"; cl["B36"] = "Score"; cl["C36"] = "Maturity"; cl["D36"] = "Target"; cl["E36"] = "Risk statement"; cl["F36"] = "Service"; cl["G36"] = "Service priority"; cl["H36"] = "Label"
PR = {}
SHORT = {"AI Governance": "Governance", "AI Risk & Security": "Risk & Security", "Training & Awareness": "Training",
         "Strategy & Value": "Strategy", "Data & Technology": "Data & Tech", "AI Privacy & Data Quality": "Privacy & Data"}
for i, p in enumerate(PILLARS):
    rr = 37 + i; PR[p] = rr
    cl.cell(rr, 1, p)
    cl.cell(rr, 2, f'=IFERROR(AVERAGEIF($C${RF}:$C${RL},A{rr},$E${RF}:$E${RL}),"")')
    cl.cell(rr, 3, f'=IF(B{rr}="","",IF(B{rr}>=0.9,"Adaptive",IF(B{rr}>=0.7,"Repeatable",IF(B{rr}>=0.5,"Informed","Partial"))))')
    cl.cell(rr, 4, 0.7)
    cl.cell(rr, 5, PILLAR_RISK[p])
    cl.cell(rr, 8, SHORT[p])
    cl.cell(rr, 9, f'=IF(B{rr}="",0,B{rr}*100)'); cl.cell(rr, 10, 70)
PF_, PL_ = 37, 37 + len(PILLARS) - 1
cl["A45"] = "Overall readiness"; cl["B45"] = f'=IF(COUNT(B{PF_}:B{PL_})=0,"",AVERAGE(B{PF_}:B{PL_}))'
cl["A46"] = "Maturity"; cl["B46"] = '=IF(B45="","",IF(B45>=0.9,"Adaptive",IF(B45>=0.7,"Repeatable",IF(B45>=0.5,"Informed","Partial"))))'
cl["A47"] = "Strongest pillar"; cl["B47"] = f'=IFERROR(INDEX(A{PF_}:A{PL_},MATCH(MAX(B{PF_}:B{PL_}),B{PF_}:B{PL_},0)),"")'; cl["C47"] = f'=IFERROR(MAX(B{PF_}:B{PL_}),"")'
cl["A48"] = "Weakest pillar"; cl["B48"] = f'=IFERROR(INDEX(A{PF_}:A{PL_},MATCH(MIN(B{PF_}:B{PL_}),B{PF_}:B{PL_},0)),"")'; cl["C48"] = f'=IFERROR(MIN(B{PF_}:B{PL_}),"")'
# exposure matrix
cl["A51"] = "Exposure matrix"; bands = ["Adaptive", "Repeatable", "Informed", "Partial"]; tiers = ["Low", "Medium", "High", "Critical"]
for j, bd in enumerate(bands): cl.cell(51, 2 + j, bd)
for i, t in enumerate(tiers):
    cl.cell(52 + i, 1, t)
    for j, bd in enumerate(bands): cl.cell(52 + i, 2 + j, EXPOSURE[(t, bd)])
cl["A57"] = "Exposure"; cl["B57"] = '=IFERROR(INDEX($B$52:$E$55,MATCH(C32,$A$52:$A$55,0),MATCH(B46,$B$51:$E$51,0)),"")'
for i, (ex, ttl, desc) in enumerate(PATHWAY):
    cl.cell(60 + i, 1, ex); cl.cell(60 + i, 2, ttl); cl.cell(60 + i, 3, desc)
cl["A58"] = "Pathway"; cl["B58"] = '=IFERROR(VLOOKUP(B57,$A$60:$C$63,2,FALSE),"")'; cl["C58"] = '=IFERROR(VLOOKUP(B57,$A$60:$C$63,3,FALSE),"")'
for i, (bd, lo, desc) in enumerate(MATURITY):
    cl.cell(66 + i, 1, bd); cl.cell(66 + i, 2, desc)
cl["A70"] = "Maturity description"; cl["B70"] = '=IFERROR(VLOOKUP(B46,$A$66:$B$69,2,FALSE),"")'
# narrative
# AI adoption maturity curve (rows 170-182)
cl["A170"] = "Adoption stage (from A1)"; cl["B170"] = '=IF(B26="","",CHOOSE(B26+1,' + ",".join(str(x) for x in A1_TO_STAGE) + '))'
cl["A171"] = "Governance-supported stage"
cl["B171"] = '=IF(B45="","",' + "".join(f'IF(B45>={t},{st},' for t, st in GOV_STAGE_THRESHOLDS) + "1" + ")" * len(GOV_STAGE_THRESHOLDS) + ")"
cl["A172"] = "Curve position"; cl["B172"] = '=IF(OR(B170="",B171=""),"",IF(AND(B171=5,B170>=3),5,B170))'
for i, (nm, desc, focus) in enumerate(CURVE_STAGES):
    cl.cell(176 + i, 1, i + 1); cl.cell(176 + i, 2, nm); cl.cell(176 + i, 3, desc); cl.cell(176 + i, 4, focus)
CN = lambda ref: f"INDEX($B$176:$B$180,{ref})"; CF_ = lambda ref: f"INDEX($D$176:$D$180,{ref})"
cl["A173"] = "Position label"; cl["B173"] = f'=IF(B172="","","Stage "&B172&" - "&{CN("B172")})'
cl["A174"] = "Curve message"
cl["B174"] = (f'=IF(B172="","",IF(B172=5,"The organisation has reached governance maturity: sustain board oversight, independent assurance and continuous improvement as AI use grows.",'
              f'IF(B170>B171,"Adoption is ahead of governance: AI use is at Stage "&B170&" ({{0}}) but governance currently supports Stage "&B171&" ({{1}}). Before scaling further, put in place "&{CF_("B170")}&", starting with "&B48&".",'
              f'IF(B170=B171,"Governance is keeping pace with adoption at Stage "&B170&" ({{0}}). To move to Stage "&(B170+1)&" ({{2}}), put in place "&{CF_("B170+1")}&".",'
              f'"Governance foundations are ahead of current adoption (Stage "&B171&" vs Stage "&B170&"): the organisation can expand AI use with confidence. Next focus: "&{CF_("B170+1")}&"."))))')
cl["B174"] = cl["B174"].value.replace("{0}", '"&' + CN("B170") + '&"').replace("{1}", '"&' + CN("B171") + '&"').replace("{2}", '"&' + CN("MIN(5,B170+1)") + '&"')
cl["A72"] = "Narrative"
cl["B72"] = ('=IF(B21=0,"The assessment is incomplete ("&B20&" of ' + str(NQ) + ' questions answered). Complete all questions in the Assessment sheet to generate the executive summary.",'
             '"Based on the responses provided, "&IF(' + ORG + '="","the organisation",' + ORG + ')&" has an overall AI readiness score of "&TEXT(B45,"0%")&", corresponding to the '
             '"&CHAR(39)&B46&CHAR(39)&" maturity level: "&LOWER(LEFT(B70,1))&MID(B70,2,300)&" Its AI use profile indicates a "&UPPER(C32)&" inherent risk tier ("&TEXT(C30,"0%")&" of maximum inherent risk), '
             'resulting in an overall AI risk exposure rating of "&UPPER(B57)&". The strongest area is "&B47&" ("&TEXT(C47,"0%")&") and the area requiring most attention is "&B48&" ("&TEXT(C48,"0%")&"). '
             'Recommended pathway: "&B58&". "&C58)')
# principles
cl["A75"] = "Responsible-AI principle coverage"
PRINC = {}
for i, pr in enumerate(AIAF_PRINCIPLES):
    rr = 76 + i; PRINC[pr] = rr
    qids = [q["id"] for q in READINESS if pr in q["principles"]]
    cl.cell(rr, 1, pr)
    cl.cell(rr, 2, "=IFERROR(AVERAGE(" + ",".join(f"E{CR[x]}" for x in qids) + '),"")')
    cl.cell(rr, 3, ", ".join(qids))
    cl.cell(rr, 4, f'=IF(B{rr}="","",IF(B{rr}>=0.75,"Well addressed",IF(B{rr}>=0.5,"Partially addressed","Needs attention")))')
# ranking (top priorities + roadmap order)
cl["A86"] = "Rank"; cl["B86"] = "Priority row"; cl["C86"] = "Roadmap row"
for i in range(len(READINESS)):
    rr = 87 + i
    cl.cell(rr, 1, i + 1)
    cl.cell(rr, 2, f'=IFERROR(IF(LARGE($H${RF}:$H${RL},A{rr})<0.5,"",MATCH(LARGE($H${RF}:$H${RL},A{rr}),$H${RF}:$H${RL},0)+{RF - 1}),"")')
    cl.cell(rr, 3, f'=IFERROR(IF(LARGE($K${RF}:$K${RL},A{rr})<1000,"",IF(INDEX($J${RF}:$J${RL},MATCH(LARGE($K${RF}:$K${RL},A{rr}),$K${RF}:$K${RL},0))<2,"",MATCH(LARGE($K${RF}:$K${RL},A{rr}),$K${RF}:$K${RL},0)+{RF - 1})),"")')
# NIST AI RMF function coverage (rows 129-132)
cl["A128"] = "NIST AI RMF function coverage"
NIST_ROW = {}
for i, (fn, fdesc, qids) in enumerate(NIST_FUNCTIONS):
    rr = 129 + i; NIST_ROW[fn] = rr
    cl.cell(rr, 1, fn)
    cl.cell(rr, 2, "=IFERROR(AVERAGE(" + ",".join(f"E{CR[x]}" for x in qids) + '),"")')
    cl.cell(rr, 3, ", ".join(qids)); cl.cell(rr, 5, fdesc)
    cl.cell(rr, 4, f'=IF(B{rr}="","",IF(B{rr}>=0.75,"Well addressed",IF(B{rr}>=0.5,"Partially addressed","Needs attention")))')
# services
for i, (svc, pil, desc) in enumerate(SERVICES):
    rr = 120 + i
    cl.cell(rr, 1, svc); cl.cell(rr, 2, pil); cl.cell(rr, 3, desc)
    sc_ = f'B{PR[pil]}'
    cl.cell(rr, 4, f'=IFERROR({sc_},"")')
    cl.cell(rr, 5, f'=IF(D{rr}="","",IF(D{rr}<0.5,"Priority",IF(D{rr}<0.7,"Advised",IF(D{rr}<0.9,"Optional","Sustain"))))')
cl.sheet_state = "hidden"


def CV(ref):  # Calc value reference
    return f"Calc!{ref}"


# =====================================================================  KPI band on Assessment
def kpi(ws, r, c1, c2, label, formula, fmt=None, big=16):
    merge(ws, r, c1, r, c2, label, size=9, bold=True, color="FFFFFF", fillc=GT2, align=CC)
    c = merge(ws, r + 1, c1, r + 1, c2, formula, size=big, bold=True, color=DARK, align=CC, fillc=LAV)
    if fmt: c.number_format = fmt
    return c


kpi(asx, KPI_ROW, 2, 3, "ANSWERED", f'={CV("B20")}&" / {NQ}"')
c_ov = kpi(asx, KPI_ROW, 4, 4, "OVERALL AI READINESS", f'=IF({CV("B45")}="","-",{CV("B45")})', "0%")
c_mat = kpi(asx, KPI_ROW, 5, 5, "MATURITY LEVEL", f'=IF({CV("B46")}="","-",{CV("B46")})')
c_tier = kpi(asx, KPI_ROW, 6, 7, "INHERENT RISK TIER", f'=IF({CV("C32")}="","-",{CV("C32")})')
c_exp = kpi(asx, KPI_ROW, 8, 8, "AI RISK EXPOSURE -> PATHWAY", f'=IF({CV("B57")}="","-",{CV("B57")}&"  ->  "&{CV("B58")})', big=12)
asx.row_dimensions[KPI_ROW].height = 20; asx.row_dimensions[KPI_ROW + 1].height = 36
for cref, dct in [(f"E{KPI_ROW+1}", MATF), (f"F{KPI_ROW+1}", RAGF)]:
    for k, v in dct.items():
        asx.conditional_formatting.add(cref, CellIsRule(operator="equal", formula=[f'"{k}"'], fill=PatternFill("solid", fgColor=v)))
for k, v in RAGF.items():
    asx.conditional_formatting.add(f"H{KPI_ROW+1}", FormulaRule(formula=[f'LEFT($H${KPI_ROW+1},{len(k)+2})="{k}  "'], fill=PatternFill("solid", fgColor=v)))

# =====================================================================  3. Report
rp = wb.create_sheet("3. Report")
rp.sheet_view.showGridLines = False
rp.sheet_view.zoomScale = 90
WID = {"A": 1.2, "B": 4.5, "C": 17, "D": 9.6, "E": 13, "F": 13, "G": 13, "H": 13, "I": 13, "J": 13, "K": 13, "L": 2, "M": 60}
for k, v in WID.items(): rp.column_dimensions[k].width = v


def chars(c1, c2):   # approx characters per merged width
    return sum(WID[get_column_letter(i)] for i in range(c1, c2 + 1)) * 1.12


SECTION_ROWS = []
SECNO = [0]


def section(r, text):
    SECTION_ROWS.append(r)
    SECNO[0] += 1
    merge(rp, r, 2, r, 11, f"{SECNO[0]}. {text}", size=12, bold=True, color="FFFFFF", fillc=GT, align=Alignment(vertical="center", indent=1), border=False)
    rp.row_dimensions[r].height = 22
    return r + 1


def th(r, spans, height=26):
    for (c1, c2, t) in spans:
        merge(rp, r, c1, r, c2, t, size=9, bold=True, color="FFFFFF", fillc=GT2, align=CC)
    rp.row_dimensions[r].height = height


def band(avg):   # maturity band text for an average expression
    return f'IF({avg}>=0.9,"Adaptive",IF({avg}>=0.7,"Repeatable",IF({avg}>=0.5,"Informed","Partial")))'


def signal(qids):   # "58% - Informed" from the average of the most relevant questions
    refs = ",".join(CV(f"E{CR[x]}") for x in qids)
    return f'=IF(COUNT({refs})=0,"-",TEXT(AVERAGE({refs}),"0%")&" - "&{band(f"AVERAGE({refs})")})'


MAT_CF = [("Partial", MATF["Partial"]), ("Informed", MATF["Informed"]), ("Repeatable", MATF["Repeatable"]), ("Adaptive", MATF["Adaptive"])]


def mat_cf(rng):   # colour a cell whose text starts with a maturity band
    for k, v in MAT_CF:
        rp.conditional_formatting.add(rng, FormulaRule(formula=[f'ISNUMBER(SEARCH("{k}",{rng.split(":")[0]}))'], fill=PatternFill("solid", fgColor=v)))


r = 1
merge(rp, r, 2, r, 11, "AI READINESS & RISK ASSESSMENT - " + ("RESULTS REPORT" if DET else "SUMMARY REPORT"), size=18, bold=True, color="FFFFFF", fillc=GT, align=Alignment(vertical="center", indent=1), border=False)
rp.row_dimensions[r].height = 40
r = 2
merge(rp, r, 2, r, 11, f'=IF({ORG}="","Organisation: (not specified)",{ORG})&"   |   "&IF({P["Sector"]}="","",{P["Sector"]}&"   |   ")&"Assessment date: "&IF({DATE}="","-",IF(ISNUMBER({DATE}),TEXT({DATE},"dd mmmm yyyy"),{DATE}))',
      size=11, bold=True, color=GT, border=False, align=WC)
rp.row_dimensions[r].height = 22
r = 3
merge(rp, r, 2, r, 11, f'="Scope: "&IF({P["Scope of assessment"]}="","-",{P["Scope of assessment"]})&IF({P["Scope description (optional)"]}="",""," - "&{P["Scope description (optional)"]})&"   |   Completed by: "&IF({P["Completed by"]}="","-",{P["Completed by"]})&IF({P["Role / title"]}="",""," ("&{P["Role / title"]}&")")',
      size=9, color="404040", border=False, align=WC)
rp.row_dimensions[r].height = 26
r = 4
merge(rp, r, 2, r, 11, f'=IF({CV("B21")}=1,"Framework basis: NIST AI RMF, ISO/IEC 42001, ISO/IEC 23894  |  Bahrain PDPL, ISO/IEC 27701 and ISO/IEC 27001  |  EU AI Act","ASSESSMENT INCOMPLETE - "&{CV("B20")}&" of {NQ} questions answered. Results below are provisional.")',
      size=8.5, italic=True, color="595959", border=False, align=WC)
rp.conditional_formatting.add("B4", FormulaRule(formula=[f"{CV('$B$21')}=0"], fill=PatternFill("solid", fgColor="E06666"), font=Font(bold=True, color="FFFFFF")))
r = 6
r = section(r, "EXECUTIVE SUMMARY")
tiles = [(2, 4, "OVERALL AI READINESS", f'=IF({CV("B45")}="","-",{CV("B45")})', "0%"),
         (5, 6, "MATURITY LEVEL", f'=IF({CV("B46")}="","-",{CV("B46")})', None),
         (7, 8, "INHERENT RISK TIER", f'=IF({CV("C32")}="","-",{CV("C32")})', None),
         (9, 11, "OVERALL AI RISK EXPOSURE", f'=IF({CV("B57")}="","-",{CV("B57")})', None)]
for c1, c2, lab, f, fmt in tiles:
    merge(rp, r, c1, r, c2, lab, size=8.5, bold=True, color="FFFFFF", fillc=GT2, align=CC)
    c = merge(rp, r + 1, c1, r + 1, c2, f, size=22, bold=True, color=DARK, fillc=LAV, align=CC)
    if fmt: c.number_format = fmt
rp.row_dimensions[r].height = 18; rp.row_dimensions[r + 1].height = 44
sub = [(2, 4, '="Maturity scale: <50% Partial | 50-70% Informed | 70-90% Repeatable | 90%+ Adaptive (four tiers, modelled on the NIST implementation tiers)"'),
       (5, 6, f'={CV("B70")}'), (7, 8, f'=IF({CV("C30")}="","",TEXT({CV("C30")},"0%")&" of maximum inherent risk. "&{CV("C33")})'),
       (9, 11, '="Combines inherent risk tier and readiness maturity (see Methodology)"')]
for c1, c2, f in sub:
    merge(rp, r + 2, c1, r + 2, c2, f, size=7.5, italic=True, color="404040", fillc=LAV, align=CC)
rp.row_dimensions[r + 2].height = 40
TILE_R = r + 1
for k, v in MATF.items():
    rp.conditional_formatting.add(f"E{TILE_R}", CellIsRule(operator="equal", formula=[f'"{k}"'], fill=PatternFill("solid", fgColor=v)))
for k, v in RAGF.items():
    rp.conditional_formatting.add(f"G{TILE_R}", CellIsRule(operator="equal", formula=[f'"{k}"'], fill=PatternFill("solid", fgColor=v)))
    rp.conditional_formatting.add(f"I{TILE_R}", CellIsRule(operator="equal", formula=[f'"{k}"'], fill=PatternFill("solid", fgColor=v)))
rp.conditional_formatting.add(f"B{TILE_R}", FormulaRule(formula=[f'AND(ISNUMBER($B${TILE_R}),$B${TILE_R}<0.5)'], fill=PatternFill("solid", fgColor=MATF["Partial"])))
rp.conditional_formatting.add(f"B{TILE_R}", FormulaRule(formula=[f'AND(ISNUMBER($B${TILE_R}),$B${TILE_R}>=0.5,$B${TILE_R}<0.7)'], fill=PatternFill("solid", fgColor=MATF["Informed"])))
rp.conditional_formatting.add(f"B{TILE_R}", FormulaRule(formula=[f'AND(ISNUMBER($B${TILE_R}),$B${TILE_R}>=0.7,$B${TILE_R}<0.9)'], fill=PatternFill("solid", fgColor=MATF["Repeatable"])))
rp.conditional_formatting.add(f"B{TILE_R}", FormulaRule(formula=[f'AND(ISNUMBER($B${TILE_R}),$B${TILE_R}>=0.9)'], fill=PatternFill("solid", fgColor=MATF["Adaptive"])))
r += 4
merge(rp, r, 2, r, 3, "RECOMMENDED PATHWAY", size=9, bold=True, color="FFFFFF", fillc=DARK, align=CC)
pw = merge(rp, r, 4, r, 11, f'=IF({CV("B58")}="","-",{CV("B58")})', size=13, bold=True, color=DARK, align=Alignment(vertical="center", indent=1))
for ex, *_ in PATHWAY:
    rp.conditional_formatting.add(f"D{r}", FormulaRule(formula=[f'{CV("$B$57")}="{ex}"'], fill=PatternFill("solid", fgColor=RAGF[ex])))
rp.row_dimensions[r].height = 26
r += 1
merge(rp, r, 2, r, 11, f'={CV("C58")}', size=9.5, color="202020", align=WC)
rp.row_dimensions[r].height = 34
r += 2
merge(rp, r, 2, r, 11, f'={CV("B72")}', size=10, color="202020", align=Alignment(wrap_text=True, vertical="top"))
rp.row_dimensions[r].height = 80
r += 2
# AI adoption maturity curve - where the organisation sits
merge(rp, r, 2, r, 11, "POSITION ON THE AI ADOPTION MATURITY CURVE", size=9, bold=True, color="FFFFFF", fillc=GT2, align=Alignment(vertical="center", indent=1))
rp.row_dimensions[r].height = 18
r += 1
for i, (nm, desc, focus) in enumerate(CURVE_STAGES):
    c1 = 2 + 2 * i
    merge(rp, r, c1, r, c1 + 1, f"{i + 1}  {nm}", size=9, bold=True, color=DARK, fillc=LAV, align=CC)
    col = get_column_letter(c1)
    rp.conditional_formatting.add(f"{col}{r}", FormulaRule(formula=[f"{CV('$B$172')}={i + 1}"], fill=PatternFill("solid", fgColor=GT), font=Font(bold=True, color="FFFFFF")))
    merge(rp, r + 1, c1, r + 1, c1 + 1, f'=IF({CV("$B$170")}={i + 1},"▲ Adoption","")&IF(AND({CV("$B$171")}={i + 1},{CV("$B$170")}<>{i + 1}),"● Governance",IF({CV("$B$171")}={i + 1},"  ● Governance",""))',
          size=8.5, bold=True, color=GT, align=CC)
rp.row_dimensions[r].height = 30; rp.row_dimensions[r + 1].height = 22
r += 2
merge(rp, r, 2, r, 11, f'={CV("B174")}', size=9.5, color="202020", align=WC)
rp.row_dimensions[r].height = 42
r += 2
# ----------------------------------------------------------- readiness by domain
DOM_START = r
r = section(r, "READINESS BY DOMAIN")
th(r, [(2, 3, "Domain"), (4, 4, "Score"), (5, 5, "Maturity"), (6, 6, "Progress")])
PT = r
r += 1
for p in PILLARS:
    pr = PR[p]
    merge(rp, r, 2, r, 3, p, size=9.5, bold=True, color=GT, align=WC)
    c = rp.cell(r, 4, f'=IF({CV(f"B{pr}")}="","-",{CV(f"B{pr}")})'); style(c, bold=True, align=CC); c.number_format = "0%"
    style(rp.cell(r, 5, f'=IF({CV(f"C{pr}")}="","-",{CV(f"C{pr}")})'), size=9, bold=True, align=CC)
    bc = merge(rp, r, 6, r, 6, f'=IF({CV(f"B{pr}")}="",0,{CV(f"B{pr}")})', size=9, bold=True, color=GT, align=Alignment(vertical="center", horizontal="left"))
    bc.number_format = ';;;'
    rp.conditional_formatting.add(f"F{r}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=GT, showValue=False))
    rp.row_dimensions[r].height = 25
    r += 1
for k, v in MATF.items():
    rp.conditional_formatting.add(f"E{PT+1}:E{r-1}", CellIsRule(operator="equal", formula=[f'"{k}"'], fill=PatternFill("solid", fgColor=v)))
merge(rp, r, 2, r, 3, "OVERALL", size=10, bold=True, color="FFFFFF", fillc=DARK, align=WC)
c = rp.cell(r, 4, f'=IF({CV("B45")}="","-",{CV("B45")})'); style(c, bold=True, align=CC, fillc=LAV); c.number_format = "0%"
style(rp.cell(r, 5, f'=IF({CV("B46")}="","-",{CV("B46")})'), bold=True, align=CC, fillc=LAV)
merge(rp, r, 6, r, 6, '="Target 70%"', size=7.5, italic=True, fillc=LAV, align=CC)
rp.row_dimensions[r].height = 25
# radar chart (6 domains)
ch = RadarChart(); ch.type = "marker"; ch.style = 2
ch.title = None
data = Reference(cl, min_col=9, max_col=10, min_row=PF_, max_row=PL_)
ch.add_data(data, titles_from_data=False)
ch.set_categories(Reference(cl, min_col=8, min_row=PF_, max_row=PL_))
ch.series[0].tx = SeriesLabel(v="Current readiness (%)"); ch.series[1].tx = SeriesLabel(v="Target - Repeatable (70%)")
ch.series[0].graphicalProperties.line.solidFill = GT; ch.series[0].graphicalProperties.line.width = 28000
ch.series[1].graphicalProperties.line.solidFill = "00A7B5"; ch.series[1].graphicalProperties.line.dashStyle = "dash"
ch.series[0].marker.symbol = "circle"; ch.series[0].marker.graphicalProperties.solidFill = GT
ch.series[1].marker.symbol = "none"
ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 100; ch.y_axis.majorUnit = 25; ch.y_axis.delete = False; ch.x_axis.delete = False
ch.legend = None
ch.height = 9.0; ch.width = 11.9
rp.add_chart(ch, f"G{PT}")
used = sum((rp.row_dimensions[x].height or 15) for x in range(PT, r + 1))
rp.row_dimensions[r + 1].height = max(15, 262 - used)
KEEP_TOGETHER = [(DOM_START, r + 1)]   # domain table + radar chart never split across pages
r += 2
if DET:
    SECTION_ROWS.append(r)
    RISK_START = r
    th(r, [(2, 3, "Domain"), (4, 11, "Potential risk if the domain is not strengthened (shown where score < 90%)")])
    r += 1
    for p in PILLARS:
        pr = PR[p]
        merge(rp, r, 2, r, 3, p, size=9, bold=True, color=GT, align=WC)
        merge(rp, r, 4, r, 11, f'=IF({CV(f"B{pr}")}="","",IF({CV(f"B{pr}")}>=0.9,"Adaptive - maintain current practices.",{CV(f"E{pr}")}))', size=9, align=WC)
        rp.row_dimensions[r].height = 28
        r += 1
    KEEP_TOGETHER.append((RISK_START, r - 1))
    r += 1
# ----------------------------------------------------------- inherent profile
r = section(r, "AI USE & INHERENT RISK PROFILE")
th(r, [(2, 3, "Factor"), (4, 7, "Your response"), (8, 8, "Risk level"), (9, 11, "Implication")])
r += 1
ctx_start = r
for q in CONTEXT:
    ar = QR[q["id"]]
    merge(rp, r, 2, r, 3, f'{q["id"]}  {q["topic"]}', size=9, bold=True, color=GT, align=WC)
    merge(rp, r, 4, r, 7, f"=IF('2. Assessment'!E{ar}=\"\",\"Not answered\",'2. Assessment'!E{ar})", size=9, align=WC)
    style(rp.cell(r, 8, f"='2. Assessment'!G{ar}"), size=9, bold=True, align=CC)
    merge(rp, r, 9, r, 11, f"='2. Assessment'!H{ar}", size=8.5, align=WC)
    mx = max(max(lines(o[0], chars(4, 7)) for o in q["opts"]), max(lines(o[1], chars(9, 11) * 1.05) for o in q["opts"]))
    rp.row_dimensions[r].height = max(36, mx * 11.5 + 6)
    r += 1
for lab, col in [("Low", "C6E0B4"), ("Medium", "FFE699"), ("High", "F4B183"), ("Very high", "E06666")]:
    rp.conditional_formatting.add(f"H{ctx_start}:H{r-1}", CellIsRule(operator="equal", formula=[f'"{lab}"'], fill=PatternFill("solid", fgColor=col)))
merge(rp, r, 2, r, 7, f'="Inherent risk tier: "&IF({CV("C32")}="","-",UPPER({CV("C32")}))&IF({CV("C30")}="",""," ("&TEXT({CV("C30")},"0%")&" weighted inherent-risk score)")', size=10, bold=True, color=DARK, fillc=LAV, align=WC)
merge(rp, r, 8, r, 11, f'={CV("C33")}', size=8.5, italic=True, fillc=LAV, align=WC)
rp.row_dimensions[r].height = 28
r += 2
# ----------------------------------------------------------- responsible-AI principles (+ NIST AI RMF functions in the detailed report)
r = section(r, "RESPONSIBLE-AI PRINCIPLE COVERAGE" + (" & NIST AI RMF FUNCTIONS" if DET else ""))
th(r, [(2, 4, "Principle"), (5, 5, "Coverage"), (6, 7, "Assessment"), (8, 11, "Informed by questions")])
r += 1
pc_start = r
for pr_ in AIAF_PRINCIPLES:
    rr = PRINC[pr_]
    merge(rp, r, 2, r, 4, pr_, size=9, bold=True, color=GT, align=WC)
    c = rp.cell(r, 5, f'=IF({CV(f"B{rr}")}="","-",{CV(f"B{rr}")})'); style(c, bold=True, align=CC, size=9); c.number_format = "0%"
    merge(rp, r, 6, r, 7, f'={CV(f"D{rr}")}', size=9, bold=True, align=CC)
    merge(rp, r, 8, r, 11, f'={CV(f"C{rr}")}', size=8.5, align=WC)
    rp.row_dimensions[r].height = 18
    r += 1
for lab, col in [("Well addressed", "C6E0B4"), ("Partially addressed", "FFE699"), ("Needs attention", "F4B183")]:
    rp.conditional_formatting.add(f"F{pc_start}:F{r-1}", CellIsRule(operator="equal", formula=[f'"{lab}"'], fill=PatternFill("solid", fgColor=col)))
if DET:
    r += 1
    th(r, [(2, 4, "NIST AI RMF function"), (5, 5, "Coverage"), (6, 7, "Assessment"), (8, 11, "Focus and questions")], 22)
    r += 1
    nf_start = r
    for fn, fdesc, qids in NIST_FUNCTIONS:
        rr = NIST_ROW[fn]
        merge(rp, r, 2, r, 4, fn, size=9, bold=True, color=GT, align=WC)
        c = rp.cell(r, 5, f'=IF({CV(f"B{rr}")}="","-",{CV(f"B{rr}")})'); style(c, bold=True, align=CC, size=9); c.number_format = "0%"
        merge(rp, r, 6, r, 7, f'={CV(f"D{rr}")}', size=9, bold=True, align=CC)
        merge(rp, r, 8, r, 11, f"{fdesc} ({', '.join(qids)})", size=8.5, align=WC)
        rp.row_dimensions[r].height = 22
        r += 1
    for lab, col in [("Well addressed", "C6E0B4"), ("Partially addressed", "FFE699"), ("Needs attention", "F4B183")]:
        rp.conditional_formatting.add(f"F{nf_start}:F{r-1}", CellIsRule(operator="equal", formula=[f'"{lab}"'], fill=PatternFill("solid", fgColor=col)))
r += 1
# ----------------------------------------------------------- findings
if DET:
    r = section(r, "DETAILED FINDINGS & RECOMMENDATIONS")
    th(r, [(2, 3, "Area"), (4, 6, "Your response & score"), (7, 8, "Potential risk"), (9, 11, "Recommended action")], 22)
    r += 1
    fd_start = r
    for q in READINESS:
        ar = QR[q["id"]]
        merge(rp, r, 2, r, 3, f'{q["id"]}  {q["topic"]}\n\n{q["pillar"]}\n{q["type"]}', size=8.5, bold=True, color=GT, align=WT)
        merge(rp, r, 4, r, 6, f"=IF('2. Assessment'!E{ar}=\"\",\"Not answered\",TEXT('2. Assessment'!F{ar},\"0%\")&\" - \"&'2. Assessment'!G{ar}&CHAR(10)&CHAR(10)&'2. Assessment'!E{ar})", size=8.5, align=WT)
        merge(rp, r, 7, r, 8, f"=IF('2. Assessment'!F{ar}=1,\"N/A - reported as fully met.\",\"{q['risk']}\")", size=8.5, align=WT)
        merge(rp, r, 9, r, 11, f"=IF('2. Assessment'!E{ar}=\"\",\"\",'2. Assessment'!H{ar})", size=8.5, align=WT)
        mx = max(max(lines(o, chars(4, 6) * 1.2) for o in q["opts"]) + 2, lines(q["risk"], chars(7, 8) * 1.2), max(lines(x, chars(9, 11) * 1.2) for x in q["recs"]) + 3, 7)
        rp.row_dimensions[r].height = mx * 10.5 + 4
        r += 1
    fd_end = r - 1
    for lab, col in [("Not met", "E06666"), ("Minimal", "F4B183"), ("Partially", "FFE699"), ("Largely", "C6E0B4"), ("Fully", "9BC2E6")]:
        rp.conditional_formatting.add(f"D{fd_start}:D{fd_end}", FormulaRule(formula=[f'ISNUMBER(SEARCH("{lab}",LEFT($D{fd_start},25)))'], fill=PatternFill("solid", fgColor=col)))
    r += 1
else:
    r = section(r, "FINDINGS BY AREA")
    th(r, [(2, 3, "Area"), (4, 5, "Your score"), (6, 11, "Observations")], 22)
    r += 1
    fd_start = r
    for q in READINESS:
        ar = QR[q["id"]]
        merge(rp, r, 2, r, 3, f'{q["id"]}  {q["topic"]}\n\n{q["pillar"]}\n{q["type"]}', size=8.5, bold=True, color=GT, align=WT)
        merge(rp, r, 4, r, 5, f"=IF('2. Assessment'!E{ar}=\"\",\"Not answered\",TEXT('2. Assessment'!F{ar},\"0%\")&\" - \"&'2. Assessment'!G{ar})", size=8.5, bold=True, align=WC)
        merge(rp, r, 6, r, 11, f"=IF('2. Assessment'!E{ar}=\"\",\"\",'2. Assessment'!E{ar})", size=8.5, align=WT)
        mx = max(max(lines(o, chars(6, 11) * 1.2) for o in q["opts"]) + 1, 5)
        rp.row_dimensions[r].height = mx * 10.5 + 6
        r += 1
    fd_end = r - 1
    for lab, col in [("Not met", "E06666"), ("Minimal", "F4B183"), ("Partially", "FFE699"), ("Largely", "C6E0B4"), ("Fully", "9BC2E6")]:
        rp.conditional_formatting.add(f"D{fd_start}:D{fd_end}", FormulaRule(formula=[f'ISNUMBER(SEARCH("{lab}",LEFT($D{fd_start},25)))'], fill=PatternFill("solid", fgColor=col)))
    r += 1
# ----------------------------------------------------------- AI privacy & data quality (detailed report only)
if DET:
    PRIV_START = r
    r = section(r, "AI PRIVACY & DATA QUALITY  (Bahrain PDPL  |  ISO/IEC 27701 PIMS  |  ISO/IEC 27001)")
    merge(rp, r, 2, r, 11, "How personal data is used, minimised, protected and quality-controlled in AI, assessed against the Bahrain Personal Data Protection Law and ISO/IEC 27701 (privacy information management system, PIMS) "
          "operated alongside ISO/IEC 27001 (information security). Scores come from the six AI Privacy & Data Quality questions plus related governance, third-party and security answers. Not legal advice.",
          size=8.5, italic=True, color="595959", border=False, align=WC)
    rp.row_dimensions[r].height = 38
    r += 1
    PRV = PR["AI Privacy & Data Quality"]
    merge(rp, r, 2, r, 4, "PRIVACY & DATA-QUALITY DOMAIN", size=9, bold=True, color="FFFFFF", fillc=DARK, align=CC)
    c = merge(rp, r, 5, r, 5, f'=IF({CV(f"B{PRV}")}="","-",{CV(f"B{PRV}")})', size=14, bold=True, color=DARK, fillc=LAV, align=CC); c.number_format = "0%"
    merge(rp, r, 6, r, 6, f'=IF({CV(f"C{PRV}")}="","-",{CV(f"C{PRV}")})', size=10, bold=True, color=DARK, fillc=LAV, align=CC)
    merge(rp, r, 7, r, 11, f'=IF({CV(f"B{PRV}")}="","",IF({CV(f"B{PRV}")}>=0.9,"Adaptive - maintain current practices.",{CV(f"E{PRV}")}))', size=8.5, align=WC)
    for k, v in MATF.items():
        rp.conditional_formatting.add(f"F{r}", CellIsRule(operator="equal", formula=[f'"{k}"'], fill=PatternFill("solid", fgColor=v)))
    rp.row_dimensions[r].height = 46
    r += 2
    th(r, [(2, 4, "Privacy & data-quality dimension"), (5, 5, "Score"), (6, 6, "Maturity"), (7, 11, "What good looks like")], 22)
    r += 1
    dm_start = r
    for q in [x for x in READINESS if x["pillar"] == "AI Privacy & Data Quality"]:
        crow = CR[q["id"]]
        merge(rp, r, 2, r, 4, f'{q["id"]}  {q["topic"]}', size=9, bold=True, color=GT, align=WC)
        c = rp.cell(r, 5, f'=IF({CV(f"E{crow}")}="","-",{CV(f"E{crow}")})'); style(c, bold=True, align=CC, size=9); c.number_format = "0%"
        style(rp.cell(r, 6, f'=IF({CV(f"E{crow}")}="","-",{band(CV(f"E{crow}"))})'), size=9, bold=True, align=CC)
        merge(rp, r, 7, r, 11, q["opts"][3], size=8, align=WC)
        rp.row_dimensions[r].height = max(30, lines(q["opts"][3], chars(7, 11)) * 10.5 + 6)
        r += 1
    for k, v in MATF.items():
        rp.conditional_formatting.add(f"F{dm_start}:F{r-1}", CellIsRule(operator="equal", formula=[f'"{k}"'], fill=PatternFill("solid", fgColor=v)))
    r += 1
    th(r, [(2, 4, "Requirement area"), (5, 7, "Bahrain PDPL"), (8, 10, "ISO/IEC 27701 and ISO/IEC 27001"), (11, 11, "Readiness")], 24)
    r += 1
    al_start = r
    ALIGN = [
        ("Lawfulness, transparency & data-subject rights",
         "A lawful basis (or consent) for each purpose; clear privacy notices; rights of access, rectification, blocking / erasure and objection, including to automated processing.",
         "ISO/IEC 27701: purposes and lawful basis identified and documented; information and rights handling for individuals (PII principals).", ["Q21", "Q6"]),
        ("Minimisation, purpose limitation & retention",
         "Process only data that is adequate, relevant and not excessive for the stated purpose; keep it no longer than necessary.",
         "ISO/IEC 27701: limit collection and processing; retention and disposal; de-identification (see ISO/IEC 20889 techniques).", ["Q22", "Q18"]),
        ("Data quality & accuracy",
         "Personal data must be accurate, complete and kept up to date (data-quality control requirements, Art. 3).",
         "ISO/IEC 27701: accuracy and quality of personal data; ISO/IEC 25012 quality model and the ISO/IEC 5259 series for data used in AI.", ["Q25", "Q26", "Q18"]),
        ("Impact assessment, processors & cross-border transfers",
         "Impact assessment for high-risk AI processing (sensitive data, profiling); processor contracts; transfers outside Bahrain only to adequate jurisdictions or with safeguards or a specific derogation.",
         "ISO/IEC 27701: privacy impact assessment, processor and third-party controls, transfer controls; ISO/IEC 42005 for AI impact assessment.", ["Q23", "Q12"]),
        ("Security of processing & breach handling",
         "Technical and organisational measures to protect personal data; incident and breach handling and notification.",
         "ISO/IEC 27001 security controls and ISO/IEC 27701 breach handling; AI-specific security and threat controls (Q13, Q14).", ["Q13", "Q14", "Q8"]),
        ("Privacy governance & management system",
         "Accountability: records of processing, a privacy lead or registered Data Protection Guardian, engagement with the regulator.",
         "ISO/IEC 27701 PIMS (standalone since the 2025 edition), integrated with the ISO/IEC 27001 ISMS and the ISO/IEC 42001 AI management system.", ["Q24", "Q21", "Q5"]),
    ]
    for area, pd_, iso, qids in ALIGN:
        merge(rp, r, 2, r, 4, area, size=8.5, bold=True, color=GT, align=WC)
        merge(rp, r, 5, r, 7, pd_, size=8, align=WC)
        merge(rp, r, 8, r, 10, iso, size=8, align=WC)
        style(rp.cell(r, 11, signal(qids)), size=8, bold=True, align=CC)
        rp.row_dimensions[r].height = max(40, max(lines(pd_, chars(5, 7)), lines(iso, chars(8, 10))) * 10.5 + 6)
        r += 1
    mat_cf(f"K{al_start}:K{r-1}")
    KEEP_TOGETHER.append((PRIV_START, r - 1))
    r += 1
# ----------------------------------------------------------- top priorities + roadmap (detailed only)
if DET:
    T5_START = r
    r = section(r, "TOP 5 PRIORITY ACTIONS")
    merge(rp, r, 2, r, 11, "Ranked by size of gap, weighted for compliance / mandatory requirements and the organisation's inherent risk tier.", size=8.5, italic=True, color="595959", border=False)
    r += 1
    th(r, [(2, 2, "#"), (3, 3, "Area"), (4, 4, "Score"), (5, 9, "Action"), (10, 11, "Timing")])
    r += 1
    for i in range(5):
        rk = 87 + i
        idx = f'{CV(f"B{rk}")}'
        style(rp.cell(r, 2, i + 1), bold=True, align=CC, fillc=LAV)
        style(rp.cell(r, 3, f'=IF({idx}="","",INDEX(Calc!$B:$B,{idx}))'), size=9, bold=True, color=GT, align=WC)
        c = rp.cell(r, 4, f'=IF({idx}="","",INDEX(Calc!$E:$E,{idx}))'); style(c, size=9, bold=True, align=CC); c.number_format = "0%"
        merge(rp, r, 5, r, 9, f'=IF({idx}="",IF({i}=0,"No gaps identified - maintain current practices.",""),INDEX(Calc!$L:$L,{idx}))', size=9, align=WC)
        merge(rp, r, 10, r, 11, f'=IF({idx}="","",INDEX(Calc!$I:$I,{idx}))', size=8.5, bold=True, align=CC)
        rp.row_dimensions[r].height = 40
        r += 1
    KEEP_TOGETHER.append((T5_START, r - 1))
    r += 1
    r = section(r, "SUGGESTED IMPLEMENTATION ROADMAP")
    th(r, [(2, 3, "Horizon"), (4, 5, "Area"), (6, 11, "Action")])
    r += 1
    rm_start = r
    for i in range(len(READINESS)):
        rk = 87 + i
        idx = f'{CV(f"C{rk}")}'
        merge(rp, r, 2, r, 3, f'=IF({idx}="","",INDEX(Calc!$I:$I,{idx}))', size=8.5, bold=True, align=WC)
        merge(rp, r, 4, r, 5, f'=IF({idx}="","",INDEX(Calc!$B:$B,{idx}))', size=8.5, bold=True, color=GT, align=WC)
        merge(rp, r, 6, r, 11, f'=IF({idx}="","",INDEX(Calc!$L:$L,{idx}))', size=8.5, align=WC)
        rp.row_dimensions[r].height = 36
        r += 1
    for lab, col in [("Immediate", "F4B183"), ("Short", "FFE699"), ("Medium", "C6E0B4")]:
        rp.conditional_formatting.add(f"B{rm_start}:B{r-1}", FormulaRule(formula=[f'LEFT($B{rm_start},{len(lab)})="{lab}"'], fill=PatternFill("solid", fgColor=col)))
    r += 1
# ----------------------------------------------------------- regulatory & standards
r = section(r, "REGULATORY & STANDARDS CONSIDERATIONS")
merge(rp, r, 2, r, 11, "Applicability is based on the Profile answers. Maturity is the average score of the questions most relevant to each item (Partial <50%, Informed 50-69%, Repeatable 70-89%, Adaptive 90%+). Not legal advice - confirm obligations with legal counsel.", size=8.5, italic=True, color="595959", border=False)
rp.row_dimensions[r].height = 26
r += 1
th(r, [(2, 4, "Law / regulation / standard"), (5, 5, "Applies?"), (6, 6, "Maturity level"), (7, 11, "Key considerations")])
r += 1
REGS = [("Bahrain Personal Data Protection Law (Law No. 30 of 2018)", f'=IF(OR({PDPL}="Yes",N({CV("B28")})>=3),"Yes","Check")', ["Q18", "Q6", "Q12", "Q21", "Q22", "Q23"],
         "Lawful basis and notice for AI processing; data minimisation and retention; impact assessments; security; data-subject rights incl. objection to automated processing; cross-border transfer rules for AI services hosted outside Bahrain."),
        ("ISO/IEC 27701 (PIMS) and ISO/IEC 27001 (ISMS)", f'=IF({PDPL}="Yes","Recommended","Reference")', ["Q24", "Q21", "Q23", "Q13"],
         "Privacy information management system (standalone since the 2025 edition), normally operated alongside an ISO/IEC 27001 security management system; evidences PDPL-aligned privacy controls for AI and supports certification or client assurance."),
        ("iGA General Policy for the Use of AI (2025) & GCC Guiding Manual on the Ethics of AI Use", f'=IF({GOV}="Yes","Yes","Reference")', ["Q6", "Q7", "Q15"],
         "National AI policy pillars (compliance, adoption, awareness, cooperation) and GCC ethics principles - mandatory reference for government entities and good practice for all."),
        ("Central Bank of Bahrain Rulebook (HC, RM, OM, BC modules)", f'=IF({CBB}="Yes","Yes","No")', ["Q5", "Q9", "Q12", "Q13", "Q20"],
         "Board oversight, risk management, cyber security, outsourcing and consumer-protection requirements apply to AI; expect model validation for material models and notification of material outsourcing."),
        ("EU AI Act (Regulation 2024/1689) - general obligations", f'=IF({EU}="Yes","Yes","No")', ["Q11", "Q15", "Q7"],
         "Extraterritorial reach. Prohibited practices (Art. 5) and AI literacy (Art. 4) apply since Feb 2025; GPAI rules since Aug 2025; transparency (Art. 50) from Aug 2026."),
        ("EU AI Act - high-risk AI systems (Annex III)", f'=IF(AND({EU}="Yes",N({CV("B27")})>=3),"Likely",IF({EU}="Yes","Check","No"))', ["Q9", "Q18", "Q19", "Q7", "Q11", "Q26"],
         "Decisions on credit, employment, insurance, essential services etc. may be high-risk: risk management, data governance, logging, human oversight, FRIA and registration (timelines subject to the EU Digital Omnibus outcome)."),
        ("KSA / UAE / GCC regimes (SDAIA AI Ethics, KSA PDPL, UAE AI Charter, DIFC Regulation 10, QCB AI Guideline)", f'=IF({GCC}="Yes","Yes","No")', ["Q6", "Q18", "Q7", "Q23"],
         "Include in the AI obligations register; align data-transfer, ethics and sector AI requirements across GCC operations."),
        ("ISO/IEC 42001:2023 AI management system", f'=IF({CV("B45")}="","-",IF({CV("B45")}>=0.7,"Ready to plan","Build first"))', ["Q5", "Q6", "Q9", "Q11", "Q8"],
         "International certifiable standard for AI governance; demonstrates trustworthy AI to clients and regulators. Recommended once readiness reaches 'Repeatable'."),
        ("Bahrain draft AI Regulation Law (Shura Council, 2024)", '="Monitor"', ["Q6", "Q10"],
         "Proposed standalone AI law (38 articles) with a new regulator and penalties - not yet in force. Maintain horizon scanning.")]
rg_start = r
for name, ap_, qids, txt in REGS:
    merge(rp, r, 2, r, 4, name, size=8.5, bold=True, color=GT, align=WC)
    style(rp.cell(r, 5, ap_), size=8.5, bold=True, align=CC)
    style(rp.cell(r, 6, signal(qids)), size=8, bold=True, align=CC)
    merge(rp, r, 7, r, 11, txt, size=8, align=WC)
    rp.row_dimensions[r].height = max(34, lines(txt, chars(7, 11)) * 10.5 + 6, lines(name, chars(2, 4)) * 11 + 6)
    r += 1
mat_cf(f"F{rg_start}:F{r-1}")
rp.conditional_formatting.add(f"E{rg_start}:E{r-1}", CellIsRule(operator="equal", formula=['"Yes"'], fill=PatternFill("solid", fgColor=LAV2), font=Font(bold=True, color=DARK)))
r += 1
# ----------------------------------------------------------- GT services
r = section(r, "HOW GRANT THORNTON CAN HELP")
if not DET:
    merge(rp, r, 2, r, 11, "Want the full picture? The Detailed AI Readiness Report adds findings, potential risks and recommendations for every question, an AI privacy & data-quality assessment "
          "(Bahrain PDPL, ISO/IEC 27701 with ISO/IEC 27001), the NIST AI RMF view, top-5 priorities, a phased roadmap and an action plan. Ask your Grant Thornton contact.",
          size=9.5, bold=True, color=DARK, fillc=LAV, align=WC)
    rp.row_dimensions[r].height = 46
    r += 1
th(r, [(2, 4, "GT AI service line"), (5, 5, "Relevance"), (6, 11, "How we help")])
r += 1
sv_start = r
for i, (svc, pil, desc) in enumerate(SERVICES):
    rr = 120 + i
    merge(rp, r, 2, r, 4, svc, size=9, bold=True, color=GT, align=WC)
    style(rp.cell(r, 5, f'={CV(f"E{rr}")}'), size=8.5, bold=True, align=CC)
    merge(rp, r, 6, r, 11, desc, size=8.5, align=WC)
    rp.row_dimensions[r].height = max(32, lines(desc, chars(6, 11)) * 11 + 8)
    r += 1
for lab, col in [("Priority", "F4B183"), ("Advised", "FFE699"), ("Optional", "C6E0B4"), ("Sustain", "9BC2E6")]:
    rp.conditional_formatting.add(f"E{sv_start}:E{r-1}", CellIsRule(operator="equal", formula=[f'"{lab}"'], fill=PatternFill("solid", fgColor=col)))
r += 1
# ----------------------------------------------------------- next steps & sign-off (detailed) / disclaimer
if DET:
    NS_START = r
    r = section(r, "NEXT STEPS & SIGN-OFF")
    steps = ("1. Review this report with the executive sponsor, the AI / risk committee and the privacy lead.\n2. Validate responses with evidence (the level selected should be demonstrable).\n"
             "3. Assign owners and dates in the '4. Action Plan' sheet for all Immediate and Short-term actions.\n4. For individual AI systems rated High or Critical, or any AI using personal data, perform a use-case risk and privacy impact assessment (GT AI Risk Assessment Toolkit).\n"
             "5. Re-run this assessment in 6-12 months or after significant AI adoption or regulatory change to track progress.")
    merge(rp, r, 2, r, 11, steps, size=9, align=WT)
    rp.row_dimensions[r].height = 72
    r += 2
    th(r, [(2, 4, "Role"), (5, 7, "Name"), (8, 10, "Signature"), (11, 11, "Date")], 20)
    r += 1
    for role in ["Assessment completed by", "Reviewed by (risk / compliance)", "Executive sponsor"]:
        merge(rp, r, 2, r, 4, role, size=9, bold=True, align=WC)
        merge(rp, r, 5, r, 7, f'=IF({P["Completed by"]}="","",{P["Completed by"]})' if role.startswith("Assessment") else "", size=9, align=WC)
        for cc_ in range(5, 12): rp.cell(r, cc_).protection = Protection(locked=False)
        merge(rp, r, 8, r, 10, "", size=9); style(rp.cell(r, 11), size=9)
        rp.row_dimensions[r].height = 28
        r += 1
    r += 1
    disc = (f"Methodology: Part A (4 questions) derives an inherent-risk tier from AI use, decision impact, data sensitivity and oversight & reversibility. Part B ({len(READINESS)} questions) scores readiness 0-100% across 6 domains "
            "(AI Governance; AI Risk & Security; Training & Awareness; Strategy & Value; Data & Technology; AI Privacy & Data Quality); domain scores are averaged equally into the overall score and mapped to a four-tier maturity scale modelled on the NIST implementation tiers. "
            "Questions are aligned to the NIST AI RMF functions (Govern, Map, Measure, Manage). Exposure combines tier and maturity.  Disclaimer: this is a self-assessment based on the responses provided and has not been independently verified by Grant Thornton. "
            "It is not an audit, certification or legal opinion.  (c) 2026 Grant Thornton Bahrain. All rights reserved.")
else:
    disc = ("Disclaimer: this summary report is a self-assessment based on the responses provided and has not been independently verified by Grant Thornton. Maturity uses a four-tier scale modelled on the NIST implementation tiers. "
            "It is not an audit, certification or legal opinion.  (c) 2026 Grant Thornton Bahrain. All rights reserved.")
merge(rp, r, 2, r, 11, disc, size=7.5, italic=True, color="595959", border=False, align=WT)
rp.row_dimensions[r].height = 58 if DET else 30
LAST = r
if DET: KEEP_TOGETHER.append((NS_START, LAST))
# ---- pagination: break before a row that would overflow the A4 page (fit-to-width scale ~0.78);
#      never leave a section header (or its table header) orphaned at the bottom of a page
PAGE_BUDGET = 925.0
acc = 0.0; prev_break = 1
for rr in range(1, LAST + 1):
    h = rp.row_dimensions[rr].height or 15
    if acc + h > PAGE_BUDGET and rr > prev_break + 1:
        brk = rr
        for sr in SECTION_ROWS:
            if rr - 3 <= sr < rr and sr > prev_break + 1:
                brk = sr
        for ks, ke in KEEP_TOGETHER:
            if ks < brk <= ke and ks > prev_break + 1:
                brk = ks
        rp.row_breaks.append(Break(id=brk - 1))
        acc = sum((rp.row_dimensions[x].height or 15) for x in range(brk, rr + 1)); prev_break = brk
    else:
        acc += h
# side panel (outside print area)
style(rp.cell(1, 13, "HOW TO PRINT, SAVE OR EMAIL THIS REPORT"), size=11, bold=True, color="FFFFFF", fillc=GT, align=WC)
rp.row_dimensions[1].height = 40
help_txt = ("SAVE AS PDF (recommended): File > Save As > choose 'PDF' as the file type > Options > 'Active sheet(s)' > Save. Or File > Export > Create PDF/XPS. This uses Excel's built-in PDF engine and does not depend on any printer driver.\n\n"
            "PRINT: File > Print > 'Print Active Sheets' (A4 portrait, fitted to one page wide, page breaks between sections). Do not choose 'Print Entire Workbook'.\n\n"
            "IF 'MICROSOFT PRINT TO PDF' SHOWS AN ERROR: the Windows PDF printer driver is failing, not the workbook. Use Save As > PDF above, or select another printer. To repair the driver: Windows 'Turn Windows features on or off' > untick 'Microsoft Print to PDF' > OK > re-tick > OK, then restart Excel.\n\n"
            "EMAIL: save the PDF, then click the link below to open a pre-filled email with your headline results and attach the PDF.\n\n"
            "This panel is outside the print area and will not appear on the printed report.")
style(rp.cell(2, 13, help_txt), size=9, fillc=LAV, align=WT)
rp.merge_cells("M2:M4")
rp.column_dimensions["M"].width = 70
body = f'"AI readiness "&TEXT({CV("B45")},"0%")&" ("&{CV("B46")}&"); risk tier "&{CV("C32")}&"; exposure "&{CV("B57")}&". Report attached."'
link = f'"mailto:"&{EMAIL}&"?subject=AI%20Readiness%20Results&body="&SUBSTITUTE(SUBSTITUTE({body},"%","%25")," ","%20")'
rp["M5"] = f'=IF({CV("B21")}=0,"Complete the assessment to enable email",IF({EMAIL}="","Enter your email address in 1. Profile to enable the email link",HYPERLINK({link},">> Click to email the results summary to "&{EMAIL})))'
style(rp["M5"], size=10, bold=True, color="0563C1", align=WC)
rp.row_dimensions[5].height = 36
rp.print_area = f"A1:K{LAST}"
rp.page_setup.orientation = "portrait"; rp.page_setup.paperSize = rp.PAPERSIZE_A4
rp.page_setup.fitToWidth = 1; rp.page_setup.fitToHeight = 0; rp.sheet_properties.pageSetUpPr.fitToPage = True  # 1 page wide; manual section breaks kept
rp.page_margins.left = rp.page_margins.right = 0.45; rp.page_margins.top = 0.6; rp.page_margins.bottom = 0.6
rp.oddFooter.left.text = "GT AI Readiness && Risk Assessment - Confidential"; rp.oddFooter.left.size = 8
rp.oddFooter.right.text = "Page &P"; rp.oddFooter.right.size = 8
rp.page_setup.firstPageNumber = 1; rp.page_setup.useFirstPageNumber = True
rp.print_options.horizontalCentered = False

# =====================================================================  4. Action plan (detailed report only)
if DET:
    ap = wb.create_sheet("4. Action Plan")
    ap.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHIJKL", [2, 6, 26, 9, 10, 22, 58, 22, 13, 12, 30, 2]): ap.column_dimensions[col].width = w
    merge(ap, 1, 2, 1, 11, "4. ACTION PLAN & RISK TREATMENT TRACKER", size=16, bold=True, color="FFFFFF", fillc=GT, align=Alignment(vertical="center", indent=1), border=False)
    ap.row_dimensions[1].height = 34
    merge(ap, 2, 2, 2, 11, "Columns B-G update automatically from the assessment. Assign an owner, target date and status (yellow cells) to track remediation. Priority rank 1 = most urgent.", size=9.5, italic=True, color=DARK, border=False)
    for i, h in enumerate(["#", "Area", "Current score", "Priority rank", "Timing", "Recommended action", "Owner", "Target date", "Status", "Notes / evidence"], 2):
        style(ap.cell(4, i, h), bold=True, color="FFFFFF", fillc=GT, align=CC)
    ap.row_dimensions[4].height = 30
    for k, q in enumerate(READINESS):
        rr = 5 + k; crow = CR[q["id"]]
        style(ap.cell(rr, 2, q["id"]), bold=True, align=CC, fillc=LAV)
        style(ap.cell(rr, 3, f'{q["topic"]}\n({q["pillar"]})'), size=9, bold=True, color=GT, align=WC)
        c = ap.cell(rr, 4, f'=IF(Calc!E{crow}="","-",Calc!E{crow})'); style(c, bold=True, align=CC); c.number_format = "0%"
        style(ap.cell(rr, 5, f'=IFERROR(MATCH({crow},Calc!$B$87:$B${86 + len(READINESS)},0),"-")'), bold=True, align=CC)
        style(ap.cell(rr, 6, f'=Calc!I{crow}'), size=9, align=WC)
        style(ap.cell(rr, 7, f'=Calc!L{crow}'), size=9, align=WC)
        for col in (8, 9, 10, 11): style(ap.cell(rr, col), size=9, fillc=INPUT, align=WC)
        ap.cell(rr, 9).number_format = "dd/mm/yyyy"
        ap.row_dimensions[rr].height = 48
    dv(ap, f"J5:J{4+len(READINESS)}", "=Lists!$F$1:$F$5")
    ap.freeze_panes = "D5"
    ap.page_setup.orientation = "landscape"; ap.page_setup.fitToWidth = 1; ap.page_setup.fitToHeight = 0; ap.sheet_properties.pageSetUpPr.fitToPage = True

# =====================================================================  Methodology & mapping
mt = wb.create_sheet("Methodology")
mt.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGHI", [2, 7, 34, 26, 22, 34, 14, 46, 2]): mt.column_dimensions[col].width = w
merge(mt, 1, 2, 1, 8, "METHODOLOGY & FRAMEWORK MAPPING", size=16, bold=True, color="FFFFFF", fillc=GT, align=Alignment(vertical="center", indent=1), border=False)
mt.row_dimensions[1].height = 34
blocks = [("How the tool is built",
           "1) Inherent-risk profile: four questions on AI use, decision impact, data sensitivity and autonomy / reversibility, with Low-Critical risk bands, pattern overrides and oversight pathways.\n"
           f"2) Readiness: {len(READINESS)} questions across six domains - AI Governance; AI Risk & Security; Training & Awareness; Strategy & Value; Data & Technology; AI Privacy & Data Quality - scored on a four-tier maturity scale (Partial, Informed, Repeatable, Adaptive) modelled on the NIST implementation tiers.\n"
           "3) Standards alignment: questions and recommendations reference the NIST AI RMF (Govern, Map, Measure, Manage), ISO/IEC 42001, ISO/IEC 23894, ISO/IEC 27701 and ISO/IEC 27001, ISO/IEC 25012 and the ISO/IEC 5259 series (data quality), the EU AI Act and the Bahrain PDPL."),
          ("Scoring - Part A (inherent risk)",
           "Each answer scores 0-4. Weights: A1 AI use x2, A2 decision impact x3, A3 data sensitivity x2, A4 oversight & reversibility x3 (max 40). Weighted % bands: <25% Low, 25-49% Medium, 50-74% High, >=75% Critical.\n"
           "Overrides (pattern risk): significant-effect decisions (A2=4) -> minimum High; A2=4 with oversight A4>=3 -> Critical; sensitive data (A3=4) -> minimum Medium."),
          ("Scoring - Part B (readiness)",
           "Each answer maps to a maturity level 0-4 scored 0%, 25%, 50%, 75%, 100%. Domain score = average of its questions; overall readiness = equal-weighted average of the 6 domain scores. Maturity: <50% Partial, 50-69% Informed, 70-89% Repeatable, >=90% Adaptive (modelled on the four NIST implementation tiers: Partial, Risk Informed, Repeatable, Adaptive)."),
          ("AI Privacy & Data Quality domain",
           "Six questions (Q21-Q26) cover privacy governance and the Bahrain PDPL, data minimisation and retention, privacy impact assessment and cross-border data, a privacy information management system (ISO/IEC 27701) operated alongside ISO/IEC 27001, data-quality management (ISO/IEC 25012, ISO/IEC 5259) and data provenance and bias. "
           "The short report shows the domain score and the PDPL / ISO/IEC 27701 view in the regulatory table; the detailed report adds a full privacy and data-quality section."),
          ("Exposure & pathway",
           "Exposure = f(inherent tier, maturity) per the matrix below. Pathways: Low -> Proceed (self-managed); Moderate -> Proceed with conditions (SME review, 90-day plan); High -> Remediate before scaling (committee oversight); Critical -> Pause high-impact AI and escalate (board / ExCo)."),
          ("Prioritisation",
           "Priority = gap (1 - score) x weight. Weight 1.5 for compliance obligations / mandatory requirements, +0.5 for inventory, guardrails, security and privacy impact assessment when the inherent tier is High or Critical. Timing: score <=25% Immediate; 50% on a weighted item Immediate; 50% otherwise Short term; 75% Medium term; 100% Sustain.")]
r = 3
for ttl, txt in blocks:
    merge(mt, r, 2, r, 8, ttl, size=11, bold=True, color="FFFFFF", fillc=GT2, align=WC); r += 1
    merge(mt, r, 2, r, 8, txt, size=9.5, align=WT); mt.row_dimensions[r].height = max(30, lines(txt, 175) * 13 + 6); r += 2
merge(mt, r, 2, r, 8, "Exposure matrix", size=11, bold=True, color="FFFFFF", fillc=GT2, align=WC); r += 1
style(mt.cell(r, 3, "Inherent tier \\ Maturity"), bold=True, fillc=LAV, align=CC)
for j, bd in enumerate(["Adaptive", "Repeatable", "Informed", "Partial"]): style(mt.cell(r, 4 + j, bd), bold=True, fillc=MATF[bd], align=CC)
r += 1
for t in ["Low", "Medium", "High", "Critical"]:
    style(mt.cell(r, 3, t), bold=True, fillc=RAGF[t], align=CC)
    for j, bd in enumerate(["Adaptive", "Repeatable", "Informed", "Partial"]):
        ex = EXPOSURE[(t, bd)]; style(mt.cell(r, 4 + j, ex), align=CC, fillc=RAGF[ex])
    r += 1
r += 1
merge(mt, r, 2, r, 8, "Question-to-framework mapping", size=11, bold=True, color="FFFFFF", fillc=GT2, align=WC); r += 1
for i, h in enumerate(["#", "Topic", "Domain", "Type", "Responsible-AI principles", "NIST AI RMF function", "Key references"], 2):
    style(mt.cell(r, i, h), bold=True, color="FFFFFF", fillc=GT, align=CC)
r += 1
REFS = {"A1": "EU AI Act Art. 3", "A2": "EU AI Act Annex III", "A3": "Bahrain PDPL", "A4": "EU AI Act Art. 14",
        "Q5": "ISO/IEC 42001 cl.5; NIST AI RMF Govern 1-2; IIA Three Lines Model", "Q6": "ISO/IEC 42001 cl.5.2, 6.2; NIST AI RMF Govern 1; EU AI Act; Bahrain PDPL",
        "Q7": "ISO/IEC 42001 A.3, A.9; NIST AI RMF Govern 3, Measure 2; EU AI Act Art. 14", "Q8": "ISO/IEC 42001 cl.9-10; NIST AI RMF Manage 4; EU AI Act Art. 72-73",
        "Q9": "ISO/IEC 23894; ISO/IEC 42001 6.1; NIST AI RMF Map / Measure / Manage", "Q10": "NIST AI RMF Govern 1.3; ISO 31000; CBB risk management requirements",
        "Q11": "ISO/IEC 42001 6.1, A.5; ISO/IEC 42005; NIST AI RMF Map 1; EU AI Act Art. 9, 27", "Q12": "ISO/IEC 42001 A.10; NIST AI RMF Govern 6; EU AI Act Art. 25; CBB outsourcing",
        "Q13": "ISO/IEC 27001; OWASP LLM Top 10; MITRE ATLAS; NIST AI RMF Manage 2", "Q14": "MITRE ATLAS; FinCEN deepfake alert; NIST AI RMF Manage",
        "Q15": "EU AI Act Art. 4; ISO/IEC 42001 7.2; NIST AI RMF Govern 2.2", "Q16": "ISO/IEC 42001 cl.4-5; NIST AI RMF Govern 1",
        "Q17": "ISO/IEC 42001 A.6; NIST AI RMF Map 1, Measure 1", "Q18": "ISO/IEC 42001 A.7; NIST AI RMF Measure 2; Bahrain PDPL",
        "Q19": "OWASP LLM Top 10; NIST AI RMF Measure 2-3; EU AI Act Art. 15", "Q20": "SR 11-7; PRA SS1/23; NIST AI RMF Measure 1-2",
        "Q21": "Bahrain PDPL; ISO/IEC 27701", "Q22": "Bahrain PDPL; ISO/IEC 27701; ISO/IEC 20889 (de-identification)",
        "Q23": "Bahrain PDPL (processors, transfers); ISO/IEC 27701; ISO/IEC 42005", "Q24": "ISO/IEC 27701; ISO/IEC 27001; ISO/IEC 42001",
        "Q25": "ISO/IEC 25012; ISO/IEC 5259 series; NIST AI RMF Measure 2", "Q26": "ISO/IEC 5259 series; NIST AI RMF Map 2-4, Measure 2.11; EU AI Act Art. 10"}
NIST_OF = {q: fn for fn, _d, qs in NIST_FUNCTIONS for q in qs}
for q in ALLQ:
    is_ctx = q["id"].startswith("A")
    vals = [q["id"], q["topic"], "Inherent risk (Part A)" if is_ctx else q["pillar"], "Inherent-risk driver" if is_ctx else q["type"],
            "All (risk exposure)" if is_ctx else ", ".join(q["principles"]), "Map" if is_ctx else NIST_OF[q["id"]], REFS[q["id"]]]
    for i, v in enumerate(vals, 2): style(mt.cell(r, i, v), size=8.5, align=WC)
    mt.row_dimensions[r].height = 30
    r += 1

# =====================================================================  Welcome
wl = wb.create_sheet("Welcome", 0)
wl.sheet_view.showGridLines = False
for col, w in zip("ABCDEFG", [2, 44, 3, 44, 3, 44, 2]): wl.column_dimensions[col].width = w
merge(wl, 1, 2, 1, 6, "GT AI READINESS & RISK ASSESSMENT TOOL" + ("" if DET else "  -  SHORT REPORT"), size=22, bold=True, color="FFFFFF", fillc=GT, align=Alignment(vertical="center", indent=1), border=False)
wl.row_dimensions[1].height = 52
merge(wl, 2, 2, 2, 6, "Grant Thornton Bahrain  |  AI Governance, Risk & Compliance", size=10.5, italic=True, color=GT, border=False)
wl.row_dimensions[2].height = 22
how = ("0. PASTE RESPONSE - paste the Forms responses (one per row) and pick one in '2. Assessment' cell D3\n\n1. PROFILE - organisation details and regulatory context\n\n"
       f"2. ASSESSMENT - {NQ} questions\n   Part A: 4 questions on AI use and inherent risk\n   Part B: {len(READINESS)} questions on readiness across 6 domains (incl. AI privacy & data quality)\n\n"
       "3. REPORT - generated automatically, print-ready (A4)" + ("\n\n4. ACTION PLAN - assign owners and dates to recommended actions" if DET else ""))
get_ = ("- Overall AI readiness score and maturity level\n- Inherent AI risk tier (Low / Medium / High / Critical)\n- Overall AI risk exposure and recommended oversight pathway\n- Domain radar chart and responsible-AI principle coverage\n"
        + ("- Findings, potential risks and recommendations per area\n- AI privacy & data-quality assessment (Bahrain PDPL, ISO/IEC 27701 with ISO/IEC 27001)\n- Top-5 priorities and a phased roadmap\n" if DET else "- Score and observation for each area\n")
        + "- Applicable regulations (Bahrain PDPL, CBB, EU AI Act, GCC)\n- Printable report you can save as PDF and email")
cols3 = [
    ("WHEN TO USE THIS TOOL", "- Before or early in your AI journey, to baseline readiness\n- Before scaling AI or deploying AI that affects customers or decisions\n- When AI is discovered in use without assessment (shadow AI)\n"
     "- When risk or context changes: new data, model, vendor, purpose or less human oversight\n- Annually, to track maturity progress\n\nPublic GenAI tools used for low-risk drafting may be governed by your Acceptable Use Policy instead."),
    ("HOW IT WORKS  (approx. 25-30 minutes)", how),
    ("WHAT YOU GET", get_)]
for i, (ttl, txt) in enumerate(cols3):
    c = 2 + i * 2
    style(wl.cell(4, c, ttl), size=11, bold=True, color="FFFFFF", fillc=GT2, align=WC)
    style(wl.cell(5, c, txt), size=9.5, fillc=LAV, align=WT)
wl.row_dimensions[4].height = 24; wl.row_dimensions[5].height = 220
style(wl.cell(7, 2, "COLOUR LEGEND"), size=10, bold=True, color=GT, border=False)
legend = [("Input cell - select or type", INPUT), ("Low / Largely met", "C6E0B4"), ("Medium / Moderate / Partially met", "FFE699"), ("High / Minimal", "F4B183"), ("Critical / Very high / Not met", "E06666")]
for i, (t, col) in enumerate(legend):
    style(wl.cell(8 + i, 2, t), size=9, fillc=col, align=WC)
style(wl.cell(7, 4, "REQUIREMENT TYPES"), size=10, bold=True, color=GT, border=False)
for i, t in enumerate(["⟳ Compliance Obligation - linked to laws / regulation", "⚠ Mandatory Requirement - baseline control expected for any AI use", "⚖ Ethical Consideration - responsible-AI principle", "☆ Recommended - good practice"]):
    style(wl.cell(8 + i, 4, t), size=9, align=WC)
style(wl.cell(7, 6, "START HERE"), size=10, bold=True, color=GT, border=False)
st_ = wl.cell(8, 6, "Go to '0. Paste Response'  >>"); style(st_, size=12, bold=True, color="FFFFFF", fillc=GT, align=CC); st_.hyperlink = "#'0. Paste Response'!A1"
st2 = wl.cell(10, 6, "Go to '2. Assessment'  >>"); style(st2, size=11, bold=True, color=GT, fillc=LAV, align=CC); st2.hyperlink = "#'2. Assessment'!A1"
st3 = wl.cell(12, 6, "View '3. Report'  >>"); style(st3, size=11, bold=True, color=GT, fillc=LAV, align=CC); st3.hyperlink = "#'3. Report'!A1"
merge(wl, 15, 2, 15, 6, "This is a self-assessment tool. Results depend on the accuracy of responses and do not constitute an audit, certification or legal advice. Version 2.0 - October 2026. GT internal scoring engine - do not distribute; clients complete the GT AI Readiness Questionnaire." + (" SAMPLE: pre-completed with illustrative responses for Grant Thornton Bahrain." if SAMPLE else ""),
      size=8.5, italic=True, color="595959", border=False)
wl.row_dimensions[15].height = 30

# =====================================================================  protection, order, finish
for ws_ in (asx, rp):
    ws_.protection.sheet = True
    ws_.protection.formatColumns = False; ws_.protection.formatRows = False
    ws_.protection.selectLockedCells = False
for q in ALLQ:
    asx.cell(QR[q["id"]], 5).protection = Protection(locked=False)
wl.sheet_properties.tabColor = GT; pf.sheet_properties.tabColor = GT; asx.sheet_properties.tabColor = GT; rp.sheet_properties.tabColor = "00A7B5"; mt.sheet_properties.tabColor = "A6A6A6"
order = [wl, ps, pf, asx, rp] + ([ap] if DET else []) + [mt, qb, cl, ls]
if DET: ap.sheet_properties.tabColor = GT2
wb._sheets = order
wb.active = 0
for ws_ in [wl, pf, asx, mt] + ([ap] if DET else []):
    ws_.page_setup.fitToWidth = 1; ws_.page_setup.fitToHeight = 0; ws_.sheet_properties.pageSetUpPr.fitToPage = True
    ws_.page_setup.orientation = "landscape"
wl.page_setup.fitToHeight = 1
ps.page_setup.orientation = "landscape"; ps.page_setup.fitToWidth = 1; ps.page_setup.fitToHeight = 0; ps.sheet_properties.pageSetUpPr.fitToPage = True
printfix.harden(wb)
wb.save(OUT)
print("saved", OUT, VARIANT, "report rows", LAST)
