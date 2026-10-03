# -*- coding: utf-8 -*-
"""GT AI Strategy & Roadmap - proposal deck (approach & methodology), built on the GT AI Service Lines template.

Usage: build_strategy_proposal.py TEMPLATE.pptx OUT.pptx IMGDIR
IMGDIR holds rendered pages of the Opportunity Discovery Toolkit worked example (tk_map.png, tk_position.png, tk_canvas.png, tk_diag_crop.png).
"""
import sys, os
from pptx.util import Emu, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.dml.color import RGBColor
sys.path.insert(0, os.path.dirname(__file__))
import deck_common as dc
from deck_common import txt, box, table, content, divider, duplicate, find, set_text, GT, GT2, LAV, LAV2, TEAL, CORAL, DARK, GREY, WHITE, BLACK, AMBER, GREEN, L, R, CW, TOP
from strategy_data import READY7, READY7_OVERALL, READY7_TARGET_2027, APPS
from opportunity_data import ARCHETYPES, FAMILIES, VERDICTS, LEVELS, STANDARDS
from tool_data import DOMAIN_MAP, PILLARS as DOMAINS

TEMPLATE, OUT, IMG = sys.argv[1], sys.argv[2], sys.argv[3]
dc.init(TEMPLATE)
add = dc.NEW.append
VCOL = {"AI fit now": RGBColor(0xA9, 0xD0, 0x8E), "Automate - AI not needed": RGBColor(0x9B, 0xC2, 0xE6), "AI later - close data gap first": RGBColor(0xFF, 0xE6, 0x99),
        "Stretch - partner or later stage": RGBColor(0xF4, 0xB1, 0x83), "Redesign / standardise first": RGBColor(0xD9, 0xD9, 0xD9)}


def pic(s, name, x, y, w=None, h=None):
    p = os.path.join(IMG, name)
    return s.shapes.add_picture(p, Emu(x), Emu(y), width=Emu(w) if w else None, height=Emu(h) if h else None)


def chevrons(s, y, items, h=620000, size=12, fills=None):
    n = len(items); gap = -60000; w = (CW - gap * (n - 1)) // n
    for i, t in enumerate(items):
        shp = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
        box(s, L + i * (w + gap), y, w, h, fill=(fills or [GT, GT2, TEAL, DARK, CORAL, GT])[i], text=t, size=size, bold=True, color=WHITE, shape=shp)
    return w


def agenda(items):
    ag = duplicate(dc.S_AGENDA)
    for old, new in zip(["Market Opportunities", "AI Service Lines for", "System Setup"], items):
        sh = find(ag, old); set_text(sh, new)
    for sh in ag.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() in items:
            sh.left = Emu(6795700); sh.width = Emu(4485523); sh.text_frame.word_wrap = True
            for p_ in sh.text_frame.paragraphs: p_.alignment = PP_ALIGN.LEFT
            for r_ in [r for p_ in sh.text_frame.paragraphs for r in p_.runs]: r_.font.size = Pt(18)
        if sh.has_text_frame and "Grant Thornton Bahrain" in sh.text_frame.text and "©" in sh.text_frame.text:
            set_text(sh, "© 2026 Grant Thornton Bahrain. All rights reserved")
    return ag


# GT Ready7 (Sept 2026) consolidated into the five GT readiness domains
cur, tgt = {}, {}
for (p, v), t in zip(READY7, READY7_TARGET_2027):
    d = DOMAIN_MAP[p]; cur.setdefault(d, []).append(v); tgt.setdefault(d, []).append(t)
CUR = [int(sum(cur[d]) / len(cur[d]) + 0.5) for d in DOMAINS]; TGT = [int(sum(tgt[d]) / len(tgt[d]) + 0.5) for d in DOMAINS]
SRC = {d: " + ".join(p for p, _ in READY7 if DOMAIN_MAP[p] == d) for d in DOMAINS}

# =========================================================== 1 cover
cov = duplicate(dc.S_COVER)
set_text(find(cov, "Artificial Intelligence"), "AI Strategy &\nRoadmap")
set_text(find(cov, "We go beyond"), "Proposal  |  October 2026")
add(cov)

# =========================================================== 2 agenda
add(agenda(["Context & Our Starting Point", "Approach & Methodology", "Engagement Plan & Deliverables"]))

# =========================================================== 3 the questions clients ask
s = content("The Questions Every Leadership Team Is Asking", "An AI strategy is only useful if it says precisely where AI creates value in this organisation - and where it does not.")
qs = [("Where exactly does AI fit?", "Which processes, functions and decisions would genuinely benefit - beyond generic 'use ChatGPT' advice?"),
      ("Do we have the data?", "Is our data digital, joined-up, reliable and deep enough for AI - or do we need foundations first?"),
      ("Is automation enough?", "For many organisations - especially smaller ones - workflow automation, RPA and dashboards deliver most of the value at lower cost and risk."),
      ("Build, buy or partner?", "What can we realistically run with our scale, budget and technology team?"),
      ("What are the risks?", "Which uses affect customers, employees or regulators (PDPL, CBB, EU AI Act) and what controls are required?"),
      ("What do we do first?", "A prioritised, costed roadmap with quick wins, measurable KPIs and clear ownership.")]
cw = (CW - 2 * 150000) // 3; chh = 1900000
for i, (q, a) in enumerate(qs):
    x = L + (i % 3) * (cw + 150000); y = TOP + (i // 3) * (chh + 150000)
    box(s, x, y, cw, chh, fill=LAV)
    box(s, x, y, cw, 90000, fill=[GT, GT2, TEAL, CORAL, DARK, GREEN][i])
    txt(s, x + 100000, y + 200000, cw - 200000, 500000, [[(q, True)]], size=16, color=GT)
    txt(s, x + 100000, y + 750000, cw - 200000, 1100000, a, size=12.5)
box(s, L, TOP + 2 * chh + 400000, CW, 520000, fill=DARK, text=[[("Our approach answers these with evidence: ", True),
    ("an organisational AI position screen, a process-by-process AI fit diagnostic and a risk-tiered roadmap - built on international standards.", False)]], size=12, color=WHITE)
add(s)

# =========================================================== 4 our starting point - Ready7 5 domains
s = content("Our Starting Point - GT Bahrain AI Ready7 Position", "We apply our methods to ourselves first. GT Bahrain's AI Ready7 self-assessment (Sept 2026), consolidated into the five GT readiness domains, against our end-2027 targets.")
cd = CategoryChartData(); cd.categories = DOMAINS
cd.add_series("Current position (Sept 2026)", CUR); cd.add_series("Target (end-2027)", TGT)
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Emu(L), Emu(TOP - 50000), Emu(int(CW * 0.58)), Emu(4300000), cd)
ch = gf.chart; ch.has_title = False; ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.BOTTOM; ch.legend.include_in_layout = False
ch.legend.font.size = Pt(10)
pl = ch.plots[0]; pl.gap_width = 70; pl.overlap = -10; pl.has_data_labels = True
pl.data_labels.number_format = '0"%"'; pl.data_labels.number_format_is_linked = False; pl.data_labels.font.size = Pt(10); pl.data_labels.font.bold = True
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
for ser, col in zip(pl.series, [GT, LAV2]):
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb = col
ch.value_axis.maximum_scale = 100; ch.value_axis.minimum_scale = 0; ch.value_axis.major_unit = 25; ch.value_axis.tick_labels.font.size = Pt(9)
ch.value_axis.has_major_gridlines = True; ch.value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
ch.category_axis.tick_labels.font.size = Pt(10)
x2 = L + int(CW * 0.60); w2 = R - x2
rows = [["Domain", "Now", "Target", "Gap"]] + [[d, f"{c:.0f}%", f"{t:.0f}%", f"{t - c:+.0f}"] for d, c, t in zip(DOMAINS, CUR, TGT)] + \
       [["Overall readiness", f"{READY7_OVERALL:.1f}%", ">= 75%", ""]]
fills = {}
for i, c in enumerate(CUR, 1): fills[(i, 1)] = RGBColor(0xF8, 0xCB, 0xAD) if c < 60 else RGBColor(0xFF, 0xE6, 0x99)
table(s, x2, TOP - 50000, w2, rows, [3.3, 0.9, 1, 0.8], size=10, rowh=360000, fills=fills)
txt(s, x2, TOP + 2550000, w2, 1700000, [[("What it means for our AI strategy work: ", True), ("we are 'Informed' (61.4%) - strongest in People & Skills and Data; weakest in Strategy & Value (50%) and Governance (60%). "
     "Our own priority is therefore exactly what this proposal offers clients: knowing where AI fits and governing it.", False)],
     [("Readiness vs fit: ", True), ("Ready7 tells us HOW READY we are; the Opportunity Discovery method tells us WHERE AI FITS. A complete strategy needs both.", False)]], size=10.5, spacing=8)
txt(s, L, TOP + 4300000, CW, 300000, "Domain scores are the averages of the AI Ready7 pillars they consolidate: " + "; ".join(f"{d} = {SRC[d]}" for d in DOMAINS[3:]) + ". Maturity: <50% Partial | 50-70% Informed | 70-90% Repeatable | 90%+ Adaptive.", size=8, color=GREY)
add(s)

# =========================================================== 5 our position applied - opportunity view
s = content("Our Starting Point - Where AI Fits in GT Bahrain", "We ran the Opportunity Discovery Toolkit on our own firm (21 processes): the result shapes our own roadmap and is the worked example clients see.")
kp = [("Buy & configure AI", "AI position (archetype C)\nData 56%  |  Scale & capacity 44%"), ("52%", "of processes: AI fit now\n(GenAI drafting, IDP, knowledge assistants)"),
      ("24%", "automation is enough\n(VAT data, billing, KYC checks, KPI packs)"), ("24%", "need a partner, data or process\nfoundations first")]
kw = (CW - 3 * 120000) // 4
for i, (big, small) in enumerate(kp):
    x = L + i * (kw + 120000)
    box(s, x, TOP - 50000, kw, 1250000, fill=[GT, RGBColor(0x5B, 0xA8, 0x5B), RGBColor(0x2E, 0x75, 0xB6), CORAL][i])
    txt(s, x + 60000, TOP + 30000, kw - 120000, 520000, [[(big, True)]], size=22 if i else 18, color=WHITE, align=PP_ALIGN.CENTER)
    txt(s, x + 60000, TOP + 600000, kw - 120000, 600000, small.split("\n"), size=10.5, color=WHITE, align=PP_ALIGN.CENTER, spacing=0)
x2 = L; w2 = CW
txt(s, x2, TOP + 1300000, w2, 400000, [[("Top-ranked opportunities (risk-adjusted priority)", True)]], size=13, color=GT)
top = [["#", "Process", "Function", "Solution family", "Verdict", "Risk"],
       ["1", "VAT return data preparation", "Tax", "Workflow automation / RPA", "Automate - AI not needed", "Low"],
       ["2", "Timesheets, WIP & billing", "Finance", "Workflow automation / RPA", "Automate - AI not needed", "Low"],
       ["3", "Meeting minutes & action tracking", "Firm-wide", "GenAI assistant (off-the-shelf)", "AI fit now", "Low"],
       ["4", "Audit working-paper & memo drafting", "Audit", "GenAI assistant (off-the-shelf)", "AI fit now", "Medium"],
       ["5", "Client acceptance, KYC & independence checks", "Risk & Quality", "Rules engine / workflow", "Automate - AI not needed", "Low"],
       ["6", "Bank & confirmation letters", "Audit", "Intelligent document processing", "AI fit now", "Medium"],
       ["7", "Methodology & knowledge Q&A", "Firm-wide", "GenAI knowledge assistant (RAG)", "AI fit now", "Low"]]
fills = {(i, 4): VCOL[r[4]] for i, r in enumerate(top[1:], 1)}
fills.update({(i, 5): RGBColor(0xC6, 0xE0, 0xB4) if r[5] == "Low" else RGBColor(0xFF, 0xE6, 0x99) for i, r in enumerate(top[1:], 1)})
table(s, x2, TOP + 1700000, w2, top, [0.3, 3.6, 1.4, 3, 2.4, 0.9], size=10, rowh=330000, fills=fills)
txt(s, x2, TOP + 4420000, w2, 400000, "Insight: our two highest-ranked opportunities need no AI at all - exactly the discipline we bring to clients. Journal-entry analytics and the AI Financial Statements tool "
    "rank as 'stretch' (custom ML beyond our position) - delivered with partners; the AI-powered BCP and Financial Statements tools remain in preparation.", size=9.5, color=GREY)
add(s)

# =========================================================== 6 divider
add(divider("Approach &\nMethodology", 2))

# =========================================================== 7 methodology overview
s = content("Our Methodology at a Glance", "Five phases over 6-8 weeks. Phase 3 - AI Opportunity Discovery - is what makes the strategy specific to the organisation.")
ph = ["1  Mobilise", "2  Situational analysis", "3  AI opportunity discovery", "4  Strategy formulation", "5  Roadmap & business case"]
w = chevrons(s, TOP - 50000, ph, h=650000, size=12.5)
detail = [(["Scope, sponsor and stakeholders", "Document request", "Interview plan", "Kick-off workshop"], "Project charter"),
          (["External scan (PESTLE: regulation, market, competitors, technology)", "Internal scan: strategy, operating model, systems", "Stakeholder interviews", "SWOT -> TOWS options",
            "Optional: AI Readiness & Risk Assessment baseline"], "Situational report, SWOT / TOWS"),
          (["Organisational AI position (data x scale)", "Process discovery workshops per function", "AI fit diagnostic per process", "Value, feasibility & risk scoring", "Opportunity map & use-case canvases"],
           "Opportunity map, top-10 canvases"),
          (["AI vision & ambition", "Objectives & KPIs", "Priority domains & portfolio", "Build / buy / partner", "Risk appetite & guardrails", "Operating model (owner, CoE, decision rights)"],
           "AI strategy document"),
          (["H1 / H2 / H3 roadmap", "Enablers: data, platforms, people, governance", "Indicative budget & benefits", "100-day plan", "KPI dashboard"], "Roadmap, business case, 100-day plan")]
for i, (items, out) in enumerate(detail):
    x = L + i * (w - 60000) + 30000; ww = w - 120000
    box(s, x, TOP + 700000, ww, 3150000, fill=LAV)
    txt(s, x + 30000, TOP + 780000, ww - 60000, 3050000, items, size=10.5, bullet=True, spacing=6)
    box(s, x, TOP + 3900000, ww, 560000, fill=WHITE, line=GT, text=out, size=9.5, bold=True, color=GT)
txt(s, L, TOP + 4520000, CW, 300000, "Week:   1   |   1-2   |   2-5   |   5-7   |   7-8        Steering checkpoints at the end of phases 2, 3 and 5.", size=9.5, color=GREY, bold=True)
add(s)

# =========================================================== 8 principles
s = content("Principles Behind Our Approach", "What makes a GT AI strategy different from a technology wish-list.")
pr = [("Problem first, not technology first", "We start from processes, pain points and decisions - then choose the technology."),
      ("Automation before AI - where it is enough", "If rules, workflow or a dashboard solve it, we say so. AI is recommended only where it adds value that automation cannot."),
      ("Data decides", "Every AI recommendation passes a data gate: is the data digital, available, reliable and (for prediction) labelled?"),
      ("Right-sized to the organisation", "A solution ceiling based on scale, budget and technology capacity: an SME and a bank get different answers."),
      ("Proportionate risk", "Each opportunity is risk-tiered (NSW AIAF / EU AI Act / PDPL) and carries minimum controls before go-live."),
      ("Standards-based and vendor-agnostic", "ISO/IEC 42001, 22989, 5259, 42005, NIST AI RMF, OECD - and no vendor commissions.")]
pw = (CW - 150000) // 2; phh = 1050000
for i, (h, d) in enumerate(pr):
    x = L + (i % 2) * (pw + 150000); y = TOP - 50000 + (i // 2) * (phh + 120000)
    box(s, x, y, 650000, phh, fill=[GT, GT2, TEAL, CORAL, DARK, GREEN][i], text=str(i + 1), size=24, bold=True, color=WHITE)
    box(s, x + 650000, y, pw - 650000, phh, fill=LAV)
    txt(s, x + 750000, y + 90000, pw - 820000, 380000, [[(h, True)]], size=13.5, color=GT)
    txt(s, x + 750000, y + 470000, pw - 820000, 560000, d, size=11)
add(s)

# =========================================================== 9 phase 1-2
s = content("Phases 1-2: Mobilise and Situational Analysis", "Understand the organisation, its environment and its starting point before discussing solutions.")
cols = [("ACTIVITIES", GT, ["Kick-off with sponsor; confirm scope, objectives and success criteria", "Review strategy, operating model, IT landscape, data assets and current AI use (incl. shadow AI)",
                            "6-10 stakeholder interviews (leadership, business units, IT, risk, HR)", "External scan: Bahrain Vision 2030, iGA AI Policy, PDPL, CBB, EU AI Act reach; sector and competitor AI use",
                            "SWOT and TOWS to derive strategic options"]),
        ("TOOLS & INPUTS", GT2, ["Document request list and interview guides", "PESTLE and SWOT / TOWS templates", "AI use inventory template (incl. embedded AI in SaaS)",
                                 "Optional baseline: GT AI Readiness & Risk Assessment (20 questions, 5 domains, auto-generated report)"]),
        ("OUTPUTS", TEAL, ["Situational analysis report", "SWOT / TOWS and strategic options", "Current AI use inventory", "Readiness baseline (if selected)"])]
cw = (CW - 2 * 150000) // 3
for i, (h, col, items) in enumerate(cols):
    x = L + i * (cw + 150000)
    box(s, x, TOP - 50000, cw, 420000, fill=col, text=h, size=13, bold=True, color=WHITE)
    box(s, x, TOP + 370000, cw, 3900000, fill=LAV)
    txt(s, x + 70000, TOP + 470000, cw - 140000, 3800000, items, size=12, bullet=True, spacing=9)
add(s)

# =========================================================== 10 toolkit overview
s = content("Phase 3: The GT AI Opportunity Discovery Toolkit", "A structured, facilitated method that answers 'where exactly does AI fit?' - process by process - and tells the organisation where automation is enough.")
steps = [("1", "Position", "Is the organisation positioned for AI? 12 questions on DATA (digitisation, history, quality, documents, integration, outcomes) and SCALE & CAPACITY (size, volumes, IT, budget, digital channels, change capacity).",
          "Archetype + solution ceiling"),
         ("2", "Discover", "Function-by-function workshops using a 60+ process library (cross-industry and Bahrain priority sectors: banking, insurance, government, healthcare, hospitality, manufacturing).",
          "Process & pain-point inventory"),
         ("3", "Diagnose", "For every process: nature of the task -> solution family (10 families, 5 levels), data gate, organisational fit and AIAF-style risk tier.",
          "Verdict per process"),
         ("4", "Prioritise", "Value (volume, effort, strategic importance, customer impact) x feasibility (data fit, standardisation, simplicity, organisational fit), risk-adjusted.",
          "Ranked opportunities"),
         ("5", "Plan", "Opportunity map, H1 / H2 / H3 roadmap lanes with the enabling foundations, and a one-page canvas per priority use case.",
          "Opportunity map, roadmap, canvases")]
sw_ = (CW - 4 * 100000) // 5
for i, (n, h, d, o) in enumerate(steps):
    x = L + i * (sw_ + 100000)
    box(s, x, TOP - 50000, sw_, 700000, fill=[GT, GT2, TEAL, DARK, CORAL][i], text=[[(n + "  " + h, True)]], size=16, color=WHITE)
    box(s, x, TOP + 650000, sw_, 2900000, fill=LAV, text=d, size=11, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
    box(s, x, TOP + 3600000, sw_, 520000, fill=WHITE, line=GT, text=o, size=10, bold=True, color=GT)
box(s, L, TOP + 4250000, CW, 520000, fill=DARK, text=[[("Not a readiness test: ", True), ("the toolkit locates opportunities and the right type of solution. Governance maturity is measured separately by the GT AI Readiness & Risk Assessment.", False)]],
    size=11, color=WHITE)
add(s)

# =========================================================== 11 step 1 position
s = content("Step 1: Organisational AI Position - Is Automation Enough?", "Two scores place the organisation in one of four archetypes. Each archetype sets a ceiling on the complexity of solution it can realistically adopt today.")
mx = L + 1200000; my = TOP + 250000; qw = 2350000; qh = 1650000
pos = {(0, 0): 2, (0, 1): 3, (1, 0): 0, (1, 1): 1}
for (rr, cc), ai in pos.items():
    code, name, ceil, head, desc = ARCHETYPES[ai]
    col = [RGBColor(0x9B, 0xC2, 0xE6), RGBColor(0xFF, 0xE6, 0x99), RGBColor(0xC6, 0xE0, 0xB4), RGBColor(0xA9, 0xD0, 0x8E)][ai]
    box(s, mx + cc * (qw + 60000), my + rr * (qh + 60000), qw, qh, fill=col,
        text=[[(f"{code}. {name}", True)], [(head, False)], [(f"Ceiling: level {ceil} - {LEVELS[ceil - 1][1]}", True)]], size=10.5, color=BLACK)
txt(s, L, my + 300000, 1150000, qh, [[("DATA", True)], [("HIGH", True)]], size=11, color=GT, align=PP_ALIGN.CENTER)
txt(s, L, my + qh + 360000, 1150000, qh, [[("DATA", True)], [("LOW", True)]], size=11, color=GT, align=PP_ALIGN.CENTER)
txt(s, mx, my + 2 * qh + 150000, qw, 300000, [[("SCALE & CAPACITY LOW", True)]], size=11, color=GT, align=PP_ALIGN.CENTER)
txt(s, mx + qw + 60000, my + 2 * qh + 150000, qw, 300000, [[("SCALE & CAPACITY HIGH", True)]], size=11, color=GT, align=PP_ALIGN.CENTER)
x2 = mx + 2 * qw + 300000; w2 = R - x2
box(s, x2, TOP - 50000, w2, 380000, fill=GT, text="12 position questions", size=12, bold=True, color=WHITE)
txt(s, x2, TOP + 380000, w2, 2000000, [[("Data position: ", True), ("how processes run (paper -> integrated systems); depth of digital history; data quality & ownership; where documents live; integration; recorded outcomes ('labels').", False)],
                                       [("Scale & capacity: ", True), ("employees; transaction volumes; in-house technology skills; budget; digital interactions; capacity for change.", False)]], size=10.5, spacing=8)
box(s, x2, TOP + 2350000, w2, 380000, fill=TEAL, text="Solution levels", size=12, bold=True, color=WHITE)
txt(s, x2, TOP + 2750000, w2, 1700000, [f"{n}. {t}" for n, t in LEVELS], size=11, spacing=4)
txt(s, L, TOP + 4450000, CW, 300000, "Basis: OECD Framework for the Classification of AI Systems (Economic Context; Data & Input), ISO/IEC 42001 cl. 4.1-4.2, NIST AI RMF MAP 1.3-1.6. "
    "A foundation flag is raised where core processes are still paper-based or no outcomes are recorded.", size=8.5, color=GREY)
add(s)

# =========================================================== 12 step 2-3 diagnose
s = content("Steps 2-3: Discover and Diagnose - From Task to Solution", "The nature of the work - not the hype - determines the solution. Each process gets one of five verdicts.")
rows = [["Nature of the task", "Solution family", "Level", "AI?", "Min. data"]]
for f in FAMILIES:
    rows.append([f[0], f[1], str(f[2]), f[3], ["None", "Partial", "Digital history", "With outcomes"][f[4]]])
fills = {}
for i, f in enumerate(FAMILIES, 1):
    fills[(i, 3)] = RGBColor(0x9B, 0xC2, 0xE6) if f[3] == "No" else RGBColor(0xC6, 0xE0, 0xB4) if f[3] == "Light" else RGBColor(0xA9, 0xD0, 0x8E)
table(s, L, TOP - 100000, int(CW * 0.62), rows, [4.2, 3.1, 0.6, 0.6, 1.3], size=8.5, rowh=355000, fills=fills)
x2 = L + int(CW * 0.62) + 150000; w2 = R - x2
box(s, x2, TOP - 100000, w2, 360000, fill=GT, text="Verdict logic (in order)", size=12, bold=True, color=WHITE)
flow = [("Process varies by person?", "Redesign / standardise first"), ("Family needs no AI?", "Automate - AI not needed"), ("Data below the family's minimum?", "AI later - close data gap first"),
        ("Level above the organisation's ceiling?", "Stretch - partner or later stage"), ("Otherwise", "AI fit now")]
for i, (q, v) in enumerate(flow):
    y = TOP + 330000 + i * 640000
    box(s, x2, y, int(w2 * 0.48), 560000, fill=LAV, text=q, size=9.5, bold=True, color=GT)
    box(s, x2 + int(w2 * 0.5), y, int(w2 * 0.5), 560000, fill=VCOL[v], text=v, size=9.5, bold=True)
txt(s, x2, TOP + 3600000, w2, 900000, "Inputs per process: task nature, volume, effort (hours / month), standardisation, data availability, strategic importance, customer / quality impact, effect on "
    "individuals' rights, personal data, intended autonomy. Basis: ISO/IEC 22989, ISO/IEC 23053, OECD AI classification, ISO/IEC 5259, ISO/IEC 42001 Annex A.7.", size=9, color=GREY)
add(s)

# =========================================================== 13 step 4 prioritise
s = content("Step 4: Prioritise - Value, Feasibility and Risk", "Transparent scoring that leadership can challenge - and that separates quick wins from strategic bets.")
cards = [("VALUE (1-5)", GT, ["30%  Volume of the process", "25%  Effort today (staff hours / month)", "25%  Strategic importance", "20%  Customer / quality impact"]),
         ("FEASIBILITY (1-5)", GT2, ["30%  Data fit (vs. family minimum)", "25%  Process standardisation", "25%  Solution simplicity (level 1-5)", "20%  Organisational fit (vs. ceiling)"]),
         ("INHERENT RISK TIER", CORAL, ["+2  Affects individuals' rights or access", "+1  Personal or confidential data", "+0-2  Autonomy (assist / approve / automated)", "+1-2  Impact of errors",
                                         "4+ High | 2-3 Medium | 0-1 Low"])]
cw = (CW - 2 * 150000) // 3
for i, (h, col, items) in enumerate(cards):
    x = L + i * (cw + 150000)
    box(s, x, TOP - 50000, cw, 420000, fill=col, text=h, size=13, bold=True, color=WHITE)
    box(s, x, TOP + 370000, cw, 1850000, fill=LAV)
    txt(s, x + 80000, TOP + 470000, cw - 160000, 1750000, items, size=12, spacing=9)
qx = L; qy = TOP + 2400000; qw2 = 1800000; qh2 = 700000
for i, (lab, col, d) in enumerate([("Strategic bets", GT2, "High value, harder"), ("Quick wins", GREEN, "High value, feasible"), ("Deprioritise", GREY, "Low value, harder"), ("Fill-ins", TEAL, "Low value, feasible")]):
    box(s, qx + (i % 2) * (qw2 + 50000), qy + (i // 2) * (qh2 + 50000), qw2, qh2, fill=col, text=[[(lab, True)], [(d, False)]], size=11, color=WHITE)
x2 = qx + 2 * qw2 + 250000; w2 = R - x2
txt(s, x2, qy, w2, 1500000, [[("Priority score ", True), ("= value x feasibility, reduced for High risk (-15%) and where redesign is needed first (-20%).", False)],
                             [("Horizons: ", True), ("H1 Now (0-6 months): quick wins and fill-ins that fit now or need only automation.  H2 Next (6-12 months): strategic bets, stretch items with partners, redesigned processes.  "
                                                      "H3 Later (12-24 months): items waiting on data foundations.", False)],
                             [("Minimum controls by tier ", True), ("(GT AI Risk Assessment Framework, ISO/IEC 42005, PDPL, EU AI Act screen) are attached to every canvas.", False)]], size=11, spacing=8)
txt(s, L, TOP + 4450000, CW, 300000, "Basis: NIST AI RMF MAP 3.1-3.2 (benefits and costs), ISO/IEC 42001 cl. 6.2 (AI objectives), NSW AI Assessment Framework, ISACA Securing AI Agents (2026) for autonomy.", size=8.5, color=GREY)
add(s)

# =========================================================== 14 outputs
s = content("Step 5: What the Client Receives from Phase 3", "Illustrative pages from the worked example (GT Bahrain).")
ih = 4500000
pic(s, "tk_position.png", L, TOP - 100000, h=ih)
p2 = pic(s, "tk_map.png", L + 3350000, TOP - 100000, h=ih)
pic(s, "tk_canvas.png", L + 6700000, TOP - 100000, h=ih)
for x, t in [(L, "Organisational AI position"), (L + 3350000, "AI opportunity map"), (L + 6700000, "Use-case canvas (one per priority)")]:
    txt(s, x, TOP + ih - 50000, 3200000, 300000, [[(t, True)]], size=10.5, color=GT)
x2 = L + 9900000
txt(s, x2, TOP - 100000, R - x2, ih, [[("Also delivered:", True)], "Full AI fit diagnostic (up to 50 processes)", "Roadmap lanes with enabling foundations", "Process library used in workshops",
                                       "Standards basis appendix"], size=10.5, spacing=8)
add(s)

# =========================================================== 15 phase 4
s = content("Phase 4: AI Strategy Formulation", "The opportunity evidence becomes a strategy leadership can approve, fund and govern.")
blocks = [("Vision & ambition", "Where AI will - and will not - be used; ambition level consistent with the AI position archetype."),
          ("Objectives & KPIs", "3-5 measurable objectives (ISO/IEC 42001 cl. 6.2), e.g., hours released, cycle time, quality, customer outcomes."),
          ("Priority portfolio", "The ranked opportunities grouped into themes; automation programme and AI programme distinguished."),
          ("Build / buy / partner", "Decided per family against the solution ceiling; vendor-selection criteria (security, data residency, no training on client data)."),
          ("Risk appetite & guardrails", "AI risk appetite statement, acceptable-use policy, approval by risk tier, human oversight."),
          ("Operating model", "Executive sponsor, AI owner / CoE, decision rights, intake and stage-gates; links to existing risk and IT governance."),
          ("People & change", "AI literacy for all users, role-based skills, change and communications plan."),
          ("Data & platforms", "Data foundations required by H2 / H3 opportunities; platform choices (automation, GenAI, data).")]
bw = (CW - 3 * 120000) // 4; bh = 1950000
for i, (h, d) in enumerate(blocks):
    x = L + (i % 4) * (bw + 120000); y = TOP - 50000 + (i // 4) * (bh + 150000)
    box(s, x, y, bw, 450000, fill=[GT, GT2, TEAL, DARK][i % 4], text=h, size=12.5, bold=True, color=WHITE)
    box(s, x, y + 450000, bw, bh - 450000, fill=LAV, text=d, size=11, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
add(s)

# =========================================================== 16 phase 5
s = content("Phase 5: Roadmap and Business Case", "A sequenced plan the organisation can execute - with the foundations each opportunity depends on.")
lanes = [("H1  Now  (0-6 months)", GT, "Quick wins and fill-ins: automation programme, off-the-shelf GenAI under policy, first configured AI pilots."),
         ("H2  Next  (6-12 months)", GT2, "Strategic bets, stretch items with partners, redesigned processes, scaling of proven pilots."),
         ("H3  Later  (12-24 months)", TEAL, "AI use cases unlocked by data foundations; advanced analytics / supervised agentic AI where the position allows.")]
for i, (h, col, d) in enumerate(lanes):
    y = TOP - 50000 + i * 760000
    box(s, L, y, 2600000, 680000, fill=col, text=h, size=13, bold=True, color=WHITE)
    box(s, L + 2600000, y, CW - 2600000, 680000, fill=LAV, text=d, size=12, align=PP_ALIGN.LEFT)
box(s, L, TOP + 2300000, CW, 380000, fill=DARK, text="ENABLERS ACROSS ALL HORIZONS", size=11.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
en = [("Governance", "AI policy, AI register, tier-based risk assessment, GT AI RCM v2.1 controls"), ("Data", "owners, quality, integration, capture of outcomes"),
      ("Platforms", "one automation platform; enterprise GenAI; data platform as needed"), ("People", "AI literacy, champions, change management"),
      ("Value tracking", "KPI dashboard, benefits realisation, quarterly portfolio review")]
ew = (CW - 4 * 100000) // 5
for i, (h, d) in enumerate(en):
    x = L + i * (ew + 100000)
    box(s, x, TOP + 2760000, ew, 1250000, fill=LAV, text=[[(h, True)], [(d, False)]], size=10.5)
txt(s, L, TOP + 4100000, CW, 600000, [[("Also delivered: ", True), ("indicative budget and benefits per horizon, 100-day mobilisation plan, KPI dashboard and governance cadence.", False)]], size=11.5)
add(s)

# =========================================================== 17 standards
s = content("Standards and Frameworks Behind the Method", "Every step is traceable to recognised international guidance - giving boards and regulators confidence in the result.")
std = [["Phase / step", "Standard or framework", "How it is used"]]
for st, name, use in STANDARDS:
    std.append([st, name, use])
table(s, L, TOP - 250000, CW, std, [1.1, 4.2, 6.6], size=7.5, rowh=232000)
add(s)

# =========================================================== 18 divider
add(divider("Engagement Plan &\nDeliverables", 3))

# =========================================================== 19 timeline
s = content("Engagement Timeline (Standard Scope: 8 Weeks)", "Indicative; adjusted to the number of functions in scope and stakeholder availability.")
weeks = 8; gx = L + 3000000; gw = CW - 3000000; cwk = gw // weeks
for k in range(weeks):
    box(s, gx + k * cwk, TOP - 100000, cwk - 20000, 340000, fill=GT, text=f"W{k + 1}", size=11, bold=True, color=WHITE)
gantt = [("1  Mobilise", 0, 1, GT), ("2  Situational analysis", 0.5, 1.5, GT2), ("3a Position screen", 1.5, 0.5, TEAL), ("3b Discovery workshops", 2, 2, TEAL), ("3c Diagnose & prioritise", 3, 1.5, TEAL),
         ("4  Strategy formulation", 4.5, 1.5, DARK), ("5  Roadmap & business case", 5.5, 2, CORAL), ("Steering checkpoints", None, None, AMBER)]
for i, (name, st, du, col) in enumerate(gantt):
    y = TOP + 330000 + i * 450000
    box(s, L, y, 2950000, 380000, fill=LAV, text=name, size=11, bold=True, color=GT, align=PP_ALIGN.LEFT)
    if st is not None:
        box(s, gx + int(st * cwk), y + 40000, int(du * cwk) - 20000, 300000, fill=col)
    else:
        for wk in (1.9, 4.4, 7.8):
            box(s, gx + int(wk * cwk), y + 60000, 260000, 260000, fill=col, shape=MSO_SHAPE.DIAMOND)
txt(s, L, TOP + 4050000, CW, 700000, [[("Client time: ", True), ("sponsor 1 hour / week; 6-10 interviews (1 hour); one 2-3 hour discovery workshop per function; position screen with leadership, IT and finance (30-45 minutes); "
                                         "two strategy workshops with the leadership team.", False)]], size=11)
add(s)

# =========================================================== 20 deliverables & options
s = content("Deliverables and Engagement Options")
rows = [["Deliverable", "Format"], ["Situational analysis, SWOT / TOWS", "Report (Word / PDF)"], ["Organisational AI position and opportunity map", "Toolkit output (PDF) + workshop deck"],
        ["AI fit diagnostic of in-scope processes", "Excel (GT AI Opportunity Discovery Toolkit)"], ["Use-case canvases for the top 10 opportunities", "One-page canvases (PDF)"],
        ["AI strategy (vision, objectives, portfolio, risk appetite, operating model)", "Strategy document + board deck"], ["Roadmap, business case, 100-day plan, KPI dashboard", "Excel + deck"]]
table(s, L, TOP - 250000, int(CW * 0.55), rows, [3.6, 2.4], size=10.5, rowh=470000)
x2 = L + int(CW * 0.55) + 200000; w2 = R - x2
opts = [("AUTOMATION & AI SCAN", TEAL, "3 weeks - SMEs", "Position screen, 1-2 discovery workshops, diagnostic of up to 15 processes, opportunity map and roadmap lanes. Answers: is automation enough?"),
        ("AI STRATEGY & ROADMAP", GT, "6-8 weeks - standard", "All five phases for the whole organisation (up to 50 processes), full strategy, roadmap and business case."),
        ("STRATEGY + READINESS", DARK, "8-10 weeks - extended", "Standard scope plus the GT AI Readiness & Risk Assessment baseline and an AI governance blueprint (policy, register, risk framework).")]
for i, (h, col, sub, d) in enumerate(opts):
    y = TOP - 250000 + i * 1450000
    box(s, x2, y, w2, 420000, fill=col, text=[[(h + "   ", True), (sub, False)]], size=11.5, color=WHITE)
    box(s, x2, y + 420000, w2, 950000, fill=LAV, text=d, size=10.5, align=PP_ALIGN.LEFT)
txt(s, L, TOP + 4100000, int(CW * 0.55), 500000, "Fees are proposed per option once scope (entities, functions, locations) is confirmed.", size=10, color=GREY)
add(s)

# =========================================================== 21 why GT
s = content("Why Grant Thornton Bahrain", "Strategy, governance, risk, cyber and assurance under one roof - with proprietary tools and the GT global network.")
why = [("Practise what we advise", "We assessed our own readiness (AI Ready7) and ran the Opportunity Discovery Toolkit on our own firm before offering it."),
       ("Trusted and regulation-ready", "Audit, GRC and cyber heritage; Bahrain PDPL, CBB, iGA AI Policy and EU AI Act mapped into our AI RCM v2.1 (81 controls)."),
       ("Honest about automation", "We recommend AI only where it beats automation - protecting budgets and reducing risk."),
       ("End-to-end AI service lines", "AI Strategy & Roadmap | AI Governance, Risk & Compliance | AI-Driven Automation & Apps | Training & Awareness.")]
half = (CW - 150000) // 2
for i, (h, d) in enumerate(why):
    y = TOP - 50000 + i * 760000
    box(s, L, y, half, 690000, fill=LAV, text=[[(h, True)], [(d, False)]], size=10.5, align=PP_ALIGN.LEFT)
x2 = L + half + 150000
box(s, x2, TOP - 50000, half, 380000, fill=GT, text="GT AI tools and apps", size=12, bold=True, color=WHITE)
rows = [["Tool / app", "Status"]] + [[a[0], a[3]] for a in APPS if a[0] != "GT AI Readiness & Risk Assessment Tool"]
rows.insert(1, ["GT AI Opportunity Discovery Toolkit", "Available (v1.0)"])
rows.insert(2, ["GT AI Readiness & Risk Assessment (20 questions, 5 domains)", "Available (v2.0)"])
fills = {(i, 1): RGBColor(0xFF, 0xE6, 0x99) for i, r in enumerate(rows) if i and r[1] in ("In preparation", "Planned")}
table(s, x2, TOP + 330000, half, rows, [4.2, 1.6], size=9.5, rowh=380000, fills=fills)
txt(s, x2, TOP + 3400000, half, 500000, "The AI-powered BCP tool and AI-powered Financial Statements tool are in preparation and are not yet available to clients.", size=9, color=GREY)
add(s)

# =========================================================== 22 thank you
ty = duplicate(dc.S_THANKS)
t = find(ty, "Grant Thornton Bahrain")
if t: set_text(t, "© 2026 Grant Thornton Bahrain. All rights reserved")
add(ty)

dc.finish(OUT)
