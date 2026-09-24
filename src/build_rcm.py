# -*- coding: utf-8 -*-
"""Builds AI RCM v2.0 (2026 update) from the original RCM + rcm_data."""
import math, sys, os, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
sys.path.insert(0, os.path.dirname(__file__))
from rcm_data import DOMAINS, EXISTING, NEW, LIBRARY

SRC = sys.argv[1]
OUT = sys.argv[2]

PURPLE = "7030A0"   # header colour used in the original RCM
GT = "4F2D7F"       # Grant Thornton purple
LAV = "EDE7F6"
GREY = "D9D9D9"
LGREY = "F2F2F2"
NEWFILL = "E2EFDA"
ENHFILL = "FFF2CC"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")


def hdr(cell, fill=PURPLE, size=11):
    cell.font = Font(name="Calibri", size=size, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.alignment = CENTER
    cell.border = BORDER


def est_lines(text, width_chars):
    if text is None:
        return 1
    n = 0
    for line in str(text).split("\n"):
        n += max(1, math.ceil(len(line) / max(1.0, width_chars)))
    return n


# ---------------------------------------------------------------- read original
wb = openpyxl.load_workbook(SRC)
orig = wb.active
orig.title = "RCM v1.0 (Original)"
original = []  # list of dicts
domain = None
for r in range(3, orig.max_row + 1):
    a = orig.cell(r, 1).value
    b = orig.cell(r, 2).value
    if a and not b and orig.cell(r, 3).value is None:
        continue
    if a:
        domain = a
    if b:
        original.append(dict(domain=domain, id=b, topic=orig.cell(r, 3).value, statement=orig.cell(r, 4).value,
                             refs=orig.cell(r, 5).value, activities=orig.cell(r, 6).value,
                             evidence=orig.cell(r, 7).value, tests=orig.cell(r, 8).value))
assert len(original) == 44, len(original)
for c in original:  # fix domain for rows where column A merged value only on first row
    pass

# domain of original controls by prefix
PREFIX_DOMAIN = {}
for c in original:
    if c["domain"]:
        PREFIX_DOMAIN[c["id"].split("-")[0]] = c["domain"]
for c in original:
    c["domain"] = PREFIX_DOMAIN[c["id"].split("-")[0]]

# ---------------------------------------------------------------- build merged list
rows = []
for c in original:
    iso, add, gcc, r7, app, typ, freq, notes = EXISTING[c["id"]]
    refs = c["refs"].replace("NIST RMF:", "NIST AI RMF:")
    rows.append(dict(domain=c["domain"], id=c["id"], topic=c["topic"], statement=c["statement"], refs=refs,
                     activities=c["activities"], evidence=c["evidence"], tests=c["tests"], iso=iso, add=add,
                     gcc=gcc, r7=r7, app=app, typ=typ, freq=freq, status="Enhanced",
                     notes=("References extended (2026 update). " + notes).strip()))
for c in NEW:
    rows.append(dict(domain=c["domain"], id=c["id"], topic=c["topic"], statement=c["statement"], refs=c["refs"],
                     activities=c["activities_txt"], evidence=c["evidence_txt"], tests=c["tests_txt"], iso=c["iso"],
                     add=c["add"], gcc=c["gcc"], r7=c["r7"], app=c["app"], typ=c["typ"], freq=c["freq"],
                     status="New", notes="New control. Gap addressed: " + c["why"]))


def id_key(cid):
    p, n = cid.split("-")
    return int(n)


rows.sort(key=lambda x: (DOMAINS.index(x["domain"]), id_key(x["id"])))

# ---------------------------------------------------------------- RCM v2.0 sheet
ws = wb.create_sheet("RCM v2.0 (2026)", 0)
cols = [
    ("Domain", 20), ("Control ID", 9), ("Topic", 22), ("Control Statement", 55), ("Control References", 30),
    ("Key control activities", 55), ("Required Evidence", 45), ("Control test plan and procedures", 55),
    ("ISO/IEC 27001:2022 Mapping (remapped)", 22), ("Additional Framework References (2026 update)", 40),
    ("Bahrain / GCC Regulatory References", 36), ("AI Ready7 Mapping", 11), ("Applicability", 11),
    ("Control Type", 11), ("Frequency", 18), ("Update Status", 10), ("Update Notes / Gap Rationale", 45),
    ("Control Owner", 16), ("Implementation Status", 15), ("Test Result", 12), ("Comments", 25),
]
for i, (name, w) in enumerate(cols, 1):
    c = ws.cell(1, i, name)
    hdr(c, PURPLE if i <= 8 else GT)
    ws.column_dimensions[get_column_letter(i)].width = w
ws.row_dimensions[1].height = 45

r = 2
current = None
FIELDS = ["domain", "id", "topic", "statement", "refs", "activities", "evidence", "tests", "iso", "add", "gcc",
          "r7", "app", "typ", "freq", "status", "notes"]
for row in rows:
    if row["domain"] != current:
        current = row["domain"]
        ws.cell(r, 1, current)
        for ci in range(1, len(cols) + 1):
            cc = ws.cell(r, ci)
            cc.fill = PatternFill("solid", fgColor=GREY)
            cc.font = Font(name="Calibri", size=12, bold=True)
            cc.border = BORDER
        if current == "Protection from AI-Enabled Threats":
            ws.cell(r, 1).value = current + "  (NEW DOMAIN - aligned to AI Ready7 Pillar 7)"
        ws.row_dimensions[r].height = 22
        r += 1
    maxlines = 1
    for ci, f in enumerate(FIELDS, 1):
        v = row[f]
        cell = ws.cell(r, ci, v)
        cell.alignment = WRAP_TOP
        cell.border = BORDER
        size = 8 if f in ("refs", "iso", "add", "gcc") else 9
        cell.font = Font(name="Calibri", size=size, bold=(f == "id"))
        maxlines = max(maxlines, est_lines(v, cols[ci - 1][1] * (1.25 if size == 8 else 1.1)))
    ws.cell(r, 1).fill = PatternFill("solid", fgColor=LGREY)
    ws.cell(r, 1).font = Font(name="Calibri", size=10, bold=True)
    for ci in range(18, 22):
        cell = ws.cell(r, ci)
        cell.border = BORDER
        cell.alignment = WRAP_TOP
        cell.font = Font(name="Calibri", size=9)
    ws.row_dimensions[r].height = min(409, max(30, maxlines * 11.5 + 6))
    r += 1
last = r - 1
ws.freeze_panes = "D2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{last}"
st_col = "P"
ws.conditional_formatting.add(f"{st_col}2:{st_col}{last}", CellIsRule(operator="equal", formula=['"New"'], fill=PatternFill("solid", fgColor=NEWFILL), font=Font(bold=True, color="375623")))
ws.conditional_formatting.add(f"{st_col}2:{st_col}{last}", CellIsRule(operator="equal", formula=['"Enhanced"'], fill=PatternFill("solid", fgColor=ENHFILL)))
ws.conditional_formatting.add(f"B2:B{last}", FormulaRule(formula=[f'$P2="New"'], fill=PatternFill("solid", fgColor=NEWFILL)))
dv1 = DataValidation(type="list", formula1='"Not started,Planned,Partially implemented,Implemented,Not applicable"', allow_blank=True)
dv2 = DataValidation(type="list", formula1='"Effective,Partially effective,Ineffective,Not tested"', allow_blank=True)
ws.add_data_validation(dv1); ws.add_data_validation(dv2)
dv1.add(f"S2:S{last}"); dv2.add(f"T2:T{last}")
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A3
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "1:1"
ws.sheet_properties.tabColor = GT

# ---------------------------------------------------------------- Cover sheet
cv = wb.create_sheet("Cover & Guide", 0)
cv.sheet_view.showGridLines = False
cv.column_dimensions["A"].width = 3
cv.column_dimensions["B"].width = 34
cv.column_dimensions["C"].width = 110
cv.merge_cells("B2:C2")
cv["B2"] = "AI Risk & Control Matrix (RCM) - Version 2.0 (2026 Update)"
cv["B2"].font = Font(name="Calibri", size=20, bold=True, color="FFFFFF")
cv["B2"].fill = PatternFill("solid", fgColor=GT)
cv["B2"].alignment = Alignment(vertical="center", indent=1)
cv.row_dimensions[2].height = 42
cv.merge_cells("B3:C3")
cv["B3"] = "Grant Thornton Bahrain | AI Governance, Risk & Compliance Service Line"
cv["B3"].font = Font(size=12, italic=True, color=GT)
n_new = sum(1 for x in rows if x["status"] == "New")
n_enh = sum(1 for x in rows if x["status"] == "Enhanced")
info = [
    ("Version / date", f"v2.0 - {datetime.date(2026, 9, 24).strftime('%d %B %Y')} (supersedes v1.0 'RCM Frameworks')"),
    ("Purpose", "Unified AI risk & control matrix used by the GT Bahrain AI GRC service line for AI governance framework design, readiness / gap assessments, ISO/IEC 42001 certification readiness, EU AI Act and Bahrain regulatory compliance reviews, and internal-audit style control testing."),
    ("What changed", f"{len(rows)} controls in {len(DOMAINS)} domains: {n_new} NEW controls (incl. new domain 'Protection from AI-Enabled Threats') and all {n_enh} original controls ENHANCED with 2026 references, Bahrain/GCC regulatory mapping, ISO/IEC 27001:2022 remapping, AI Ready7 linkage and control attributes. See 'Gap Analysis' and 'Change Log'."),
    ("Sources used for the update", "EU AI Act (Reg. 2024/1689) & GPAI Code of Practice; ISO/IEC 42001, 23894, 42005, 42006, 5338, 22989, 24027/24028/24368, 25059, 5259; ISO/IEC 27001/27002:2022, 27701:2025; NIST AI RMF, AI 600-1, AI 100-2 E2025; OWASP LLM & Agentic Top 10; MITRE ATLAS; ETSI EN 304 223; OECD/UNESCO/IEEE/CoE; Bahrain PDPL, iGA General Policy for the Use of AI (2025), GCC AI Ethics Manual, Bahrain draft AI law, CBB Rulebook, NCSC; KSA/UAE/Qatar regimes; NSW AIAF; SR 11-7 / PRA SS1/23; GT AI Ready7 (7 pillars, 33 items)."),
    ("Sheet guide", "RCM v2.0 (2026) - the working matrix (filterable; columns R-U for engagement use)\nGap Analysis - every gap identified and how it was resolved\nAI Ready7 Mapping - coverage of each AI Ready7 item by RCM controls\nRegulatory Library - laws, standards and frameworks with status and key dates\nChange Log - change summary per control and mapping corrections\nRCM v1.0 (Original) - unchanged original for traceability"),
    ("Column guide (new columns I-Q)", "I  ISO/IEC 27001:2022 Annex A / clause mapping (original references mixed 2013 and 2022 numbering)\nJ  Additional framework references added in this update\nK  Bahrain / GCC regulatory references\nL  AI Ready7 item(s) the control supports (e.g., 2.4 = Governance & Ethics - Risk Assessment)\nM  Applicability: Provider (develops/places AI on market), Deployer (uses AI), Both\nN  Control type: Directive / Preventive / Detective / Corrective\nO  Recommended operating frequency\nP  Update status: New / Enhanced\nQ  Update notes / gap rationale"),
    ("How to use on engagements", "1. Filter by Applicability (Provider / Deployer) and by client jurisdiction.\n2. Scope controls by the client's AI risk tier (RM-5) - Low-tier use cases may apply a reduced set.\n3. Record Control Owner, Implementation Status and Test Result in columns R-T.\n4. Use the test plan (column H) and evidence list (column G) as the audit programme."),
    ("Important note", "Regulatory dates and statuses reflect information available at September 2026. The EU AI Act high-risk timelines are subject to the Digital Omnibus legislative process, and the Bahrain draft AI law is not yet in force. Items marked 'confirm' in the Regulatory Library should be verified against primary sources before client reliance."),
]
rr = 5
for k, v in info:
    cv.cell(rr, 2, k).font = Font(bold=True, color=GT, size=11)
    cv.cell(rr, 2).alignment = Alignment(vertical="top", wrap_text=True)
    c = cv.cell(rr, 3, v)
    c.alignment = Alignment(vertical="top", wrap_text=True)
    c.font = Font(size=10)
    cv.row_dimensions[rr].height = max(20, est_lines(v, 118) * 14 + 4)
    for col in (2, 3):
        cv.cell(rr, col).border = Border(bottom=Side(style="thin", color="D0C4E4"))
    rr += 1
# domain summary table
rr += 1
cv.cell(rr, 2, "Domain summary").font = Font(bold=True, size=12, color=GT)
rr += 1
for i, h in enumerate(["Domain", "Controls (Total / New / Enhanced)"], 2):
    hdr(cv.cell(rr, i, h), GT)
rr += 1
for d in DOMAINS:
    tot = sum(1 for x in rows if x["domain"] == d)
    nw = sum(1 for x in rows if x["domain"] == d and x["status"] == "New")
    cv.cell(rr, 2, d).border = BORDER
    cv.cell(rr, 3, f"{tot}  /  {nw} new  /  {tot - nw} enhanced").border = BORDER
    rr += 1
cv.cell(rr, 2, "TOTAL").font = Font(bold=True)
cv.cell(rr, 3, f"{len(rows)}  /  {n_new} new  /  {n_enh} enhanced").font = Font(bold=True)
cv.sheet_properties.tabColor = GT

# ---------------------------------------------------------------- Gap analysis
ga = wb.create_sheet("Gap Analysis", 2)
gcols = [("#", 5), ("Gap Category", 18), ("Source (law / standard / framework)", 36), ("Gap identified in RCM v1.0", 60),
         ("Resolution in RCM v2.0", 16), ("Priority", 10), ("Regulatory driver date", 22)]
for i, (h, w) in enumerate(gcols, 1):
    hdr(ga.cell(1, i, h), GT)
    ga.column_dimensions[get_column_letter(i)].width = w
PRIMARY = {
    "GL-4": ("Regulatory obligation", "EU AI Act Art. 4; iGA AI Policy Pillar 3; AI Ready7 Pillar 3", "High", "2 Feb 2025 (in effect)"),
    "GL-5": ("AI Ready7 coverage", "AI Ready7 1.2, 2.1; NSW AIAF", "High", "-"),
    "GL-6": ("AI Ready7 coverage", "AI Ready7 1.2, 1.3; ISO/IEC 42001 6.2", "Medium", "-"),
    "GL-7": ("AI Ready7 coverage", "AI Ready7 2.3; NIST AI RMF Govern 2.1; NSW AIAF (AIRC)", "Medium", "-"),
    "GL-8": ("Good practice / regulatory", "AI Ready7 2.5; ISO/IEC 42001 A.4; EU AI Act Art. 49", "High", "-"),
    "RM-5": ("Regulatory obligation", "EU AI Act Art. 5, 6, Annex III; NSW AIAF risk bands", "High", "2 Feb 2025 (Art. 5 in effect)"),
    "RM-6": ("AI Ready7 coverage", "AI Ready7 1.1, 4.4; NIST AI RMF Govern 1.3", "Medium", "-"),
    "RO-5": ("Regulatory obligation", "Multi-jurisdiction (Bahrain, GCC, EU, APAC)", "High", "Continuous"),
    "RO-6": ("Regulatory obligation", "EU AI Act Art. 26, 27", "High", "2 Aug 2026 (subject to Digital Omnibus)"),
    "LC-6": ("Regulatory obligation", "EU AI Act Art. 53-55; GPAI Code of Practice", "High", "2 Aug 2025 (in effect)"),
    "LC-7": ("Regulatory obligation", "EU AI Act Art. 50(2), 50(4); NIST AI 600-1", "High", "2 Aug 2026"),
    "LC-8": ("Emerging risk", "OWASP Agentic Top 10; Singapore Agentic AI framework; AI Ready7 4.2", "High", "-"),
    "LC-9": ("Sector requirement", "SR 11-7; PRA SS1/23; CBB licensees", "Medium", "-"),
    "LC-10": ("AI Ready7 coverage", "AI Ready7 6.4; ISO/IEC 5338", "Low", "-"),
    "LC-11": ("Regulatory obligation", "EU AI Act Art. 53(1)(c)-(d); Bahrain Copyright Law", "Medium", "2 Aug 2025"),
    "SE-6": ("Emerging risk", "OWASP LLM Top 10 2025; MITRE ATLAS; EU AI Act Art. 15(5)", "High", "-"),
    "SE-7": ("Emerging risk", "OWASP LLM03; CISA AI Data Security", "High", "-"),
    "RS-7": ("Regulatory obligation", "EU AI Act Art. 16(l); EU Accessibility Act", "Low", "28 Jun 2025"),
    "RS-8": ("Good practice", "ISO/IEC 42001 impact considerations; ISO/IEC TR 20226", "Low", "-"),
    "RS-9": ("Regulatory obligation", "EU AI Act Art. 26(7); AI Ready7 3.4", "Medium", "2 Aug 2026"),
    "PR-5": ("Regulatory obligation", "Bahrain PDPL data-subject rights; GDPR Art. 22", "High", "In effect"),
    "PR-6": ("Regulatory obligation", "Bahrain PDPL transfers; Decree No. 56/2018; CBB outsourcing", "High", "In effect"),
    "TP-3": ("AI Ready7 coverage", "AI Ready7 6.2; EU MCC-AI", "Medium", "-"),
    "CO-3": ("Regulatory obligation", "EU AI Act Art. 86; NSW AIAF contestability", "Medium", "2 Aug 2026"),
    "AT-1": ("AI Ready7 coverage", "AI Ready7 7.1", "Medium", "-"),
    "AT-2": ("AI Ready7 coverage", "AI Ready7 7.2, 7.4", "Medium", "-"),
    "AT-3": ("AI Ready7 coverage", "AI Ready7 7.3; FinCEN deepfake alert", "High", "-"),
    "AT-4": ("AI Ready7 coverage", "AI Ready7 7.5", "Medium", "-"),
}
gaps = []
_orig_ids = {o["id"] for o in original}
_all_items = [f"{p}.{k}" for p, k in [(1,1),(1,2),(1,3),(1,4),(2,1),(2,2),(2,3),(2,4),(2,5),(3,1),(3,2),(3,3),(3,4),(4,1),(4,2),(4,3),(4,4),(4,5),(5,1),(5,2),(5,3),(5,4),(5,5),(6,1),(6,2),(6,3),(6,4),(6,5),(7,1),(7,2),(7,3),(7,4),(7,5)]]
V1_GAPS = sum(1 for it in _all_items if not any(it in [t.strip() for t in x["r7"].split(",")] for x in rows if x["id"] in _orig_ids))
for c in NEW:
    cat, src, pri, dt = PRIMARY[c["id"]]
    gaps.append((cat, src, c["why"], f"New control {c['id']} - {c['topic']}", pri, dt))
gaps += [
    ("Mapping accuracy", "ISO/IEC 27001:2022", "Annex A references mixed 2013 numbering (e.g., A.9, A.12, A.14, A.18) and 2022 numbering (e.g., A.5.2, A.8.1-A.8.7). The 2013 transition deadline passed on 31 Oct 2025.", "Column I: curated 2022 remapping for all 72 controls", "High", "31 Oct 2025"),
    ("Mapping accuracy", "EU AI Act Art. 4", "GL-1 and CO-2 referenced 'EU AI ACT 4.1' - Art. 4 is AI literacy, not executive commitment or stakeholder engagement.", "Noted in GL-1 / CO-2 update notes; Art. 4 mapped to new GL-4", "Medium", "-"),
    ("Mapping accuracy", "EU AI Act Art. 13 / 86", "RS-4 (Explainability) mapped only to Art. 50.2-50.5 (GenAI transparency).", "Art. 13 and Art. 86 added in column J", "Medium", "-"),
    ("Mapping accuracy", "EU AI Act Art. 10", "RS-5 (Fairness) mapped to Art. 5(1) (prohibitions) rather than bias-related data-governance duties.", "Art. 10(2)(f)-(g), 10(5) added in column J", "Medium", "-"),
    ("Terminology", "NIST AI RMF", "'NIST RMF' is ambiguous with NIST SP 800-37 Risk Management Framework.", "Relabelled 'NIST AI RMF' throughout", "Low", "-"),
    ("Mapping currency", "ISO/IEC 27701", "Clause references are to the 2019 edition; the 2025 edition is a standalone PIMS standard.", "Noted in PR-1; remap clauses at next review", "Low", "-"),
    ("Coverage", "Bahrain / GCC regulation", "No Bahrain or GCC legal/regulatory references in v1.0 (e.g., PDPL, iGA AI Policy, CBB Rulebook, NCSC, GCC AI Ethics Manual).", "Column K added for all controls; RO-5 obligations register", "High", "-"),
    ("Coverage", "Deployer vs provider", "v1.0 was largely provider-centric; applicability not indicated.", "Column M 'Applicability' + RO-6 deployer obligations", "High", "-"),
    ("Usability", "Audit methodology", "Controls lacked type, frequency, owner and testing-result fields; 3-row merged cells prevented filtering/sorting.", "Columns N, O, R-U added; one row per control with auto-filter", "Medium", "-"),
    ("Coverage", "GT AI Ready7", f"{V1_GAPS} of 33 AI Ready7 items had no mapped RCM control (incl. all of Pillar 7 - Protection from AI Threats).", "All 33 items now mapped - see 'AI Ready7 Mapping'", "High", "-"),
]
for i, g in enumerate(gaps, 1):
    vals = (i,) + g
    for j, v in enumerate(vals, 1):
        c = ga.cell(i + 1, j, v)
        c.alignment = WRAP_TOP
        c.border = BORDER
        c.font = Font(size=9)
    ga.row_dimensions[i + 1].height = max(30, est_lines(g[2], 66) * 12 + 4)
ga.conditional_formatting.add(f"F2:F{len(gaps)+1}", CellIsRule(operator="equal", formula=['"High"'], fill=PatternFill("solid", fgColor="F8CBAD")))
ga.conditional_formatting.add(f"F2:F{len(gaps)+1}", CellIsRule(operator="equal", formula=['"Medium"'], fill=PatternFill("solid", fgColor="FFE699")))
ga.conditional_formatting.add(f"F2:F{len(gaps)+1}", CellIsRule(operator="equal", formula=['"Low"'], fill=PatternFill("solid", fgColor="C6E0B4")))
ga.freeze_panes = "A2"
ga.auto_filter.ref = f"A1:G{len(gaps)+1}"
ga.page_setup.orientation = "landscape"; ga.page_setup.fitToWidth = 1; ga.page_setup.fitToHeight = 0
ga.sheet_properties.pageSetUpPr.fitToPage = True

# ---------------------------------------------------------------- AI Ready7 mapping
READY7 = [
    ("1", "Business Alignment", [("1.1", "Strategy, Culture & Risk Appetite"), ("1.2", "Portfolio Governance"), ("1.3", "Value Realisation"), ("1.4", "Client & Stakeholder Assurance")]),
    ("2", "Governance & Ethics", [("2.1", "AI Policy & Acceptable Use"), ("2.2", "Regulatory Compliance & Audit Readiness"), ("2.3", "Adoption Guardrails & Ethical Standards"), ("2.4", "Risk Assessment & Governance Reviews"), ("2.5", "Inventory & Ownership")]),
    ("3", "Training & Development", [("3.1", "Enterprise AI Literacy & Learning Pathways"), ("3.2", "Ethical Use, Validation & Quality Awareness"), ("3.3", "Security & Risk Awareness"), ("3.4", "Role-Based Capability Development")]),
    ("4", "Architecture & Technology", [("4.1", "Enterprise AI Reference Architecture"), ("4.2", "Standardised AI Integration & Execution Authority"), ("4.3", "API, Access & Trust Enforcement"), ("4.4", "Observability, Monitoring & Performance Controls"), ("4.5", "Cost, Compute & Platform Governance")]),
    ("5", "Data & Quality", [("5.1", "Data Classification, Access & Handling"), ("5.2", "Human Validation & Oversight"), ("5.3", "Secure Data Pipelines"), ("5.4", "Data Lineage, Quality & Observability"), ("5.5", "Privacy & Anonymisation Safeguards")]),
    ("6", "Protect AI Systems", [("6.1", "Identity, Access & Privilege Management"), ("6.2", "Data & Model Protection Controls"), ("6.3", "Secure AI Development & ModelOps Pipeline"), ("6.4", "AI Security Testing & Adversarial Resilience"), ("6.5", "Runtime Protection, Logging & Auditability")]),
    ("7", "Protection from AI Threats", [("7.1", "AI-Specific Threat Intelligence"), ("7.2", "Augmented Detection & Security Operations"), ("7.3", "Identity, Behavioural Analytics & Deepfake Defence"), ("7.4", "Automated Incident Response & Vulnerability Management"), ("7.5", "Adversarial Testing & Exposure Management")]),
]
mp = wb.create_sheet("AI Ready7 Mapping", 3)
mcols = [("AI Ready7 Pillar", 26), ("Item", 7), ("Item Title", 44), ("RCM v1.0 controls", 30), ("RCM v2.0 controls (all)", 40), ("New controls added", 22), ("Coverage v1.0", 13), ("Coverage v2.0", 13)]
for i, (h, w) in enumerate(mcols, 1):
    hdr(mp.cell(1, i, h), GT)
    mp.column_dimensions[get_column_letter(i)].width = w
rr = 2
uncovered_v1 = 0
for pn, pname, items in READY7:
    for iid, title in items:
        allc = [x["id"] for x in rows if iid in [s.strip() for s in x["r7"].split(",")]]
        v1 = [x for x in allc if any(o["id"] == x for o in original)]
        nw = [x for x in allc if x not in v1]
        uncovered_v1 += (0 if v1 else 1)
        vals = [f"{pn}. {pname}", iid, title, ", ".join(v1) or "-", ", ".join(allc), ", ".join(nw) or "-",
                "Covered" if v1 else "Gap", "Covered" if allc else "Gap"]
        for j, v in enumerate(vals, 1):
            c = mp.cell(rr, j, v)
            c.alignment = WRAP_TOP; c.border = BORDER; c.font = Font(size=9)
        rr += 1
for col in ("G", "H"):
    mp.conditional_formatting.add(f"{col}2:{col}{rr-1}", CellIsRule(operator="equal", formula=['"Gap"'], fill=PatternFill("solid", fgColor="F8CBAD")))
    mp.conditional_formatting.add(f"{col}2:{col}{rr-1}", CellIsRule(operator="equal", formula=['"Covered"'], fill=PatternFill("solid", fgColor="C6E0B4")))
mp.freeze_panes = "A2"
print("Ready7 items without v1 coverage:", uncovered_v1)

# ---------------------------------------------------------------- Regulatory library
lb = wb.create_sheet("Regulatory Library", 4)
lcols = [("#", 4), ("Law / Standard / Framework", 44), ("Issuer", 24), ("Jurisdiction", 16), ("Type", 16), ("Status & key dates", 50), ("Relevance for GT Bahrain clients", 46), ("Primary RCM controls", 22)]
for i, (h, w) in enumerate(lcols, 1):
    hdr(lb.cell(1, i, h), GT)
    lb.column_dimensions[get_column_letter(i)].width = w
for i, item in enumerate(LIBRARY, 1):
    vals = (i,) + item
    ml = 1
    for j, v in enumerate(vals, 1):
        c = lb.cell(i + 1, j, v)
        c.alignment = WRAP_TOP; c.border = BORDER; c.font = Font(size=9)
        ml = max(ml, est_lines(v, lcols[j - 1][1] * 1.1))
    lb.row_dimensions[i + 1].height = ml * 12 + 4
    if "Bahrain" in item[2] or "GCC" in item[2]:
        lb.cell(i + 1, 4).fill = PatternFill("solid", fgColor=LAV)
lb.freeze_panes = "C2"
lb.auto_filter.ref = f"A1:H{len(LIBRARY)+1}"

# ---------------------------------------------------------------- Change log
cl = wb.create_sheet("Change Log", 5)
ccols = [("Control ID", 10), ("Topic", 34), ("Domain", 26), ("Status", 10), ("Summary of change", 90)]
for i, (h, w) in enumerate(ccols, 1):
    hdr(cl.cell(1, i, h), GT)
    cl.column_dimensions[get_column_letter(i)].width = w
for i, x in enumerate(rows, 2):
    for j, v in enumerate([x["id"], x["topic"], x["domain"], x["status"], x["notes"]], 1):
        c = cl.cell(i, j, v)
        c.alignment = WRAP_TOP; c.border = BORDER; c.font = Font(size=9)
    cl.row_dimensions[i].height = max(15, est_lines(x["notes"], 98) * 12 + 3)
cl.conditional_formatting.add(f"D2:D{len(rows)+1}", CellIsRule(operator="equal", formula=['"New"'], fill=PatternFill("solid", fgColor=NEWFILL)))
cl.conditional_formatting.add(f"D2:D{len(rows)+1}", CellIsRule(operator="equal", formula=['"Enhanced"'], fill=PatternFill("solid", fgColor=ENHFILL)))
cl.freeze_panes = "A2"
cl.auto_filter.ref = f"A1:E{len(rows)+1}"
vh = len(rows) + 3
cl.cell(vh, 1, "Version history").font = Font(bold=True, color=GT, size=12)
cl.cell(vh + 1, 1, "v1.0"); cl.cell(vh + 1, 2, "RCM Frameworks (original)"); cl.cell(vh + 1, 5, "44 controls, 12 domains; ISO 42001 / 27001 / 27701, EU AI Act, NIST AI RMF, SOC 2.")
cl.cell(vh + 2, 1, "v2.0"); cl.cell(vh + 2, 2, "2026 update (this version)"); cl.cell(vh + 2, 5, f"{len(rows)} controls, {len(DOMAINS)} domains; {n_new} new controls, all originals enhanced; see Gap Analysis.")

orig.sheet_properties.tabColor = "A6A6A6"
for _ws in wb.worksheets:
    _ws.page_setup.fitToWidth = 1
    _ws.page_setup.fitToHeight = 0
    _ws.sheet_properties.pageSetUpPr.fitToPage = True
    if _ws.title != "Cover & Guide":
        _ws.page_setup.orientation = "landscape"
    _ws.page_margins.left = _ws.page_margins.right = 0.4
wb.active = 0
wb.save(OUT)
print("rows", len(rows), "new", n_new, "enhanced", n_enh)
