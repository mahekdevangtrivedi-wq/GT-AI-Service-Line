# -*- coding: utf-8 -*-
"""GT AI Readiness & Risk Assessment - proposal deck (approach & methodology), built on the GT AI Service Lines template.

Usage: build_readiness_proposal.py TEMPLATE.pptx OUT.pptx IMGDIR
IMGDIR holds rendered pages of SAMPLE_Report_GT_Bahrain.pdf (rep0.png, rep2.png, rep4.png ...).
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
from strategy_data import READY7, READY7_OVERALL
from tool_data import DOMAIN_MAP, PILLARS as DOMAINS, READINESS, CONTEXT, MATURITY, PATHWAY, PROFILE_Q

TEMPLATE, OUT, IMG = sys.argv[1], sys.argv[2], sys.argv[3]
dc.init(TEMPLATE)
add = dc.NEW.append
DCOL = [GT, GT2, TEAL, DARK, CORAL]
# GT Bahrain sample assessment with the v2 tool (SAMPLE_Report_GT_Bahrain.pdf)
SAMPLE = {"Strategy & Value": 50, "Governance, Risk & Compliance": 53, "People & Skills": 75, "Data & Technology": 58, "AI Security & Threat Risk": 63}
SAMPLE_OVERALL = 60
cur = {}
for p, v in READY7: cur.setdefault(DOMAIN_MAP[p], []).append(v)
R7 = {d: int(sum(cur[d]) / len(cur[d]) + 0.5) for d in DOMAINS}
NQ = {d: [q for q in READINESS if q["pillar"] == d] for d in DOMAINS}
NPROF = len(PROFILE_Q); NCTX = len(CONTEXT); NREAD = len(READINESS)


def pic(s, name, x, y, w=None, h=None, border=True):
    p = s.shapes.add_picture(os.path.join(IMG, name), Emu(x), Emu(y), width=Emu(w) if w else None, height=Emu(h) if h else None)
    if border: p.line.color.rgb = LAV2; p.line.width = Pt(0.75)
    return p


def chevrons(s, y, items, h=620000, size=12):
    n = len(items); gap = -60000; w = (CW - gap * (n - 1)) // n
    for i, t in enumerate(items):
        box(s, L + i * (w + gap), y, w, h, fill=[GT, GT2, TEAL, DARK, CORAL][i], text=t, size=size, bold=True, color=WHITE,
            shape=MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON)
    return w


def agenda(items):
    ag = duplicate(dc.S_AGENDA)
    for old, new in zip(["Market Opportunities", "AI Service Lines for", "System Setup"], items):
        set_text(find(ag, old), new)
    for sh in ag.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() in items:
            sh.left = Emu(6795700); sh.width = Emu(4485523); sh.text_frame.word_wrap = True
            for p_ in sh.text_frame.paragraphs: p_.alignment = PP_ALIGN.LEFT
            for r_ in [r for p_ in sh.text_frame.paragraphs for r in p_.runs]: r_.font.size = Pt(18)
        if sh.has_text_frame and "Grant Thornton Bahrain" in sh.text_frame.text and "©" in sh.text_frame.text:
            set_text(sh, "© 2026 Grant Thornton Bahrain. All rights reserved")
    return ag


# =========================================================== 1 cover
cov = duplicate(dc.S_COVER)
set_text(find(cov, "Artificial Intelligence"), "AI Readiness &\nRisk Assessment")
set_text(find(cov, "We go beyond"), "Proposal  |  October 2026")
add(cov)
add(agenda(["Why Assess Now", "Approach & Methodology", "Engagement Plan & Deliverables"]))

# =========================================================== 3 why now
s = content("Why Assess AI Readiness and Risk Now", "AI is already in the organisation - in staff tools, in software and in suppliers' services. Boards and regulators now expect it to be governed.")
why = [("Regulation is arriving", "EU AI Act obligations phase in to Aug 2026-27 (incl. AI literacy since Feb 2025); Bahrain PDPL applies to AI processing personal data; iGA AI Policy (2025); CBB expectations for licensees; draft Bahrain AI law."),
       ("Shadow AI and embedded AI", "Staff use public GenAI tools; vendors switch on AI features in existing software. Without an inventory, risk is invisible."),
       ("Agentic AI raises the stakes", "AI that takes actions through tools and APIs introduces prompt injection, excessive agency and new fraud paths (ISACA, 2026)."),
       ("AI-enabled threats", "Deepfake payment fraud, AI-crafted phishing and automated attacks target every organisation - whether or not it uses AI itself."),
       ("Boards are asking", "Who owns AI? What is our risk appetite? Which AI do we use, with what data, and who checks it?"),
       ("Investment needs a baseline", "A measured starting point focuses budgets on the gaps that matter and shows progress over time.")]
cw = (CW - 2 * 150000) // 3; chh = 1950000
for i, (h, d) in enumerate(why):
    x = L + (i % 3) * (cw + 150000); y = TOP - 50000 + (i // 3) * (chh + 150000)
    box(s, x, y, cw, chh, fill=LAV); box(s, x, y, cw, 90000, fill=[GT, GT2, TEAL, CORAL, DARK, GREEN][i])
    txt(s, x + 100000, y + 180000, cw - 200000, 450000, [[(h, True)]], size=15, color=GT)
    txt(s, x + 100000, y + 650000, cw - 200000, 1250000, d, size=11.5)
add(s)

# =========================================================== 4 what it answers
s = content("What the Assessment Answers", "Four answers in one report - in about 45 minutes of the client's time.")
ans = [("How ready are we?", "Readiness score and maturity (Partial / Informed / Repeatable / Adaptive) overall and for each of five domains.", GT),
       ("How risky is our AI?", "Inherent AI risk tier (Low to Very high) from how AI is used, the decisions it influences, the data it touches and the oversight in place - based on the NSW AI Assessment Framework.", GT2),
       ("Is our readiness enough for our risk?", "Overall AI risk exposure and a recommended pathway: proceed, proceed with conditions, remediate before scaling, or pause.", TEAL),
       ("What should we do first?", "Findings per question, top-5 priority actions, a phased roadmap, regulatory considerations (PDPL, CBB, EU AI Act, ISO/IEC 42001) and next steps.", CORAL)]
for i, (h, d, col) in enumerate(ans):
    y = TOP - 50000 + i * 1100000
    box(s, L, y, 900000, 1000000, fill=col, text=str(i + 1), size=28, bold=True, color=WHITE)
    box(s, L + 900000, y, 3600000, 1000000, fill=LAV, text=h, size=16, bold=True, color=GT, align=PP_ALIGN.LEFT)
    box(s, L + 4500000, y, CW - 4500000, 1000000, fill=WHITE, line=LAV2, text=d, size=12.5, align=PP_ALIGN.LEFT)
add(s)

# =========================================================== 5 divider
add(divider("Approach &\nMethodology", 2))

# =========================================================== 6 approach overview
s = content("Our Approach - Five Steps", "Light-touch for the client, rigorous behind the scenes. The scoring engine stays with GT - the client receives a clear, independent report.")
w = chevrons(s, TOP - 50000, ["1  Kick-off", "2  Online questionnaire", "3  GT analysis", "4  Report", "5  Debrief & roadmap"], h=650000, size=12.5)
det = [(["30-minute call with the sponsor", "Confirm scope (whole organisation, unit or AI system)", "Agree respondents (risk, IT, business)", "Confidentiality terms"], "Day 0"),
       ([f"Secure Microsoft Forms questionnaire", f"{NPROF + NCTX + NREAD} questions in three parts: profile ({NPROF}), AI use profile ({NCTX}), readiness ({NREAD})", "Plain-language answer options; about 45 minutes",
         "Optional evidence upload"], "Week 1-2"),
        (["Responses processed in GT's proprietary assessment engine", "Inherent risk tier, domain maturity, exposure and pathway calculated", "Findings, priorities and roadmap generated", "Reviewed by a GT AI GRC professional"], "Week 2-3"),
        (["Confidential PDF report (7-9 pages)", "Emailed to the sponsor", "Headline results in the cover email"], "Week 3"),
        (["60-90 minute debrief with leadership", "Validate key answers (optional evidence review)", "Agree priorities and owners", "Option: AI governance build, RCM review or ISO/IEC 42001 readiness"], "Week 3-4")]
for i, (items, when) in enumerate(det):
    x = L + i * (w - 60000) + 30000; ww = w - 120000
    box(s, x, TOP + 700000, ww, 3150000, fill=LAV)
    txt(s, x + 30000, TOP + 780000, ww - 60000, 3050000, items, size=10.5, bullet=True, spacing=7)
    box(s, x, TOP + 3900000, ww, 450000, fill=WHITE, line=GT, text=when, size=11, bold=True, color=GT)
add(s)

# =========================================================== 7 five domains
s = content("Five Assessment Domains", f"{NREAD} readiness questions across five equally weighted domains, aligned to GT AI Ready7, ISO/IEC 42001, NIST AI RMF and the NSW AI Assessment Framework.")
dw = (CW - 4 * 100000) // 5
for i, d in enumerate(DOMAINS):
    x = L + i * (dw + 100000)
    box(s, x, TOP - 50000, dw, 700000, fill=DCOL[i], text=[[(d, True)]], size=12.5, color=WHITE)
    box(s, x, TOP + 650000, dw, 300000, fill=LAV2, text=f"{len(NQ[d])} question{'s' if len(NQ[d]) > 1 else ''}", size=10.5, bold=True, color=GT)
    box(s, x, TOP + 950000, dw, 3150000, fill=LAV)
    txt(s, x + 30000, TOP + 1020000, dw - 60000, 3050000, [q["topic"] for q in NQ[d]], size=10.5, bullet=True, spacing=6)
box(s, L, TOP + 4200000, CW, 480000, fill=DARK, text=[[("New in v2.0: ", True), ("five additional governance and risk questions - governance structure, AI risk management framework, risk appetite & KRIs, third-party AI risk and model risk management.", False)]],
    size=11, color=WHITE)
add(s)

# =========================================================== 8 inherent risk
s = content("Inherent AI Risk Profile (NSW AIAF-based)", "Four questions establish how much risk the organisation's AI use carries - before considering its controls.")
ctx = [(c["topic"], c["q"]) for c in CONTEXT]
for i, (h, q) in enumerate(ctx):
    y = TOP - 50000 + i * 820000
    box(s, L, y, 2700000, 740000, fill=[GT, GT2, TEAL, DARK][i], text=h, size=13, bold=True, color=WHITE)
    box(s, L + 2700000, y, 4200000, 740000, fill=LAV, text=q, size=10.5, align=PP_ALIGN.LEFT)
x2 = L + 7100000; w2 = R - x2
tiers = [("Low", RGBColor(0xC6, 0xE0, 0xB4)), ("Medium", RGBColor(0xFF, 0xE6, 0x99)), ("High", RGBColor(0xF4, 0xB1, 0x83)), ("Very high", RGBColor(0xE0, 0x66, 0x66))]
box(s, x2, TOP - 50000, w2, 380000, fill=GT, text="Inherent risk tier", size=12, bold=True, color=WHITE)
for i, (t, col) in enumerate(tiers):
    box(s, x2, TOP + 380000 + i * 430000, w2, 400000, fill=col, text=t, size=12, bold=True)
txt(s, x2, TOP + 2150000, w2, 1300000, "Overrides apply - for example, AI that makes or materially influences decisions about individuals' rights or access to services cannot be rated Low, "
    "whatever the other answers.", size=10.5)
box(s, L, TOP + 3400000, CW, 650000, fill=LAV2, text=[[("Exposure = inherent risk tier x readiness maturity. ", True), ("A highly mature organisation with high-risk AI can be at moderate exposure; "
    "an immature organisation with the same AI is at critical exposure. The pathway follows from the exposure.", False)]], size=11.5, align=PP_ALIGN.LEFT)
txt(s, L, TOP + 4150000, CW, 400000, "Basis: NSW AI Assessment Framework (inherent risk, oversight pathways); EU AI Act risk categories; ISO/IEC 23894 and ISO/IEC 42005 (impact assessment).", size=9, color=GREY)
add(s)

# =========================================================== 9 scoring
s = content("How Results Are Expressed", "Clients see clear, comparable results; the detailed scoring rules remain GT's proprietary methodology.")
box(s, L, TOP - 50000, CW, 380000, fill=GT, text="Five-level answer scale for every readiness question", size=12, bold=True, color=WHITE)
lv = ["Not in place", "Initial / ad hoc", "Defined", "Implemented", "Optimised"]
lw = (CW - 4 * 80000) // 5
for i, t in enumerate(lv):
    box(s, L + i * (lw + 80000), TOP + 400000, lw, 520000, fill=[RGBColor(0xF8, 0xCB, 0xAD), RGBColor(0xFC, 0xE4, 0xD6), RGBColor(0xFF, 0xE6, 0x99), RGBColor(0xC6, 0xE0, 0xB4), RGBColor(0x9B, 0xC2, 0xE6)][i],
        text=f"Level {i}  -  {t}", size=11.5, bold=True)
box(s, L, TOP + 1100000, CW, 380000, fill=GT2, text="Maturity bands (GT AI Ready7)", size=12, bold=True, color=WHITE)
bands = [("Partial", "< 50%"), ("Informed", "50 - 70%"), ("Repeatable", "70 - 90%"), ("Adaptive", "90%+")]
bw = (CW - 3 * 80000) // 4
for i, ((b, rng), (_, _, desc)) in enumerate(zip(bands, MATURITY)):
    box(s, L + i * (bw + 80000), TOP + 1550000, bw, 1250000, fill=[RGBColor(0xF8, 0xCB, 0xAD), RGBColor(0xFF, 0xE6, 0x99), RGBColor(0xC6, 0xE0, 0xB4), RGBColor(0x9B, 0xC2, 0xE6)][i],
        text=[[(f"{b}  ({rng})", True)], [(desc, False)]], size=10.5)
box(s, L, TOP + 2950000, CW, 380000, fill=TEAL, text="Pathways", size=12, bold=True, color=WHITE)
pw = (CW - 3 * 80000) // 4
for i, (exp, path, desc) in enumerate(PATHWAY):
    box(s, L + i * (pw + 80000), TOP + 3400000, pw, 900000, fill=LAV, text=[[(f"{exp} exposure", True)], [(path, True)]], size=10.5, color=GT)
txt(s, L, TOP + 4380000, CW, 350000, "Domains are equally weighted. Mapping of answers to findings, risks, controls (GT AI RCM v2.1) and recommendations is applied in the GT engine and reviewed by a GT professional.",
    size=9, color=GREY)
add(s)

# =========================================================== 10 collection & confidentiality
s = content("How Answers Are Collected - and Why the Engine Stays with GT", "The client never handles spreadsheets or formulas. Responses flow securely from a form to a reviewed PDF report.")
flow = [("Client", "Completes the secure Microsoft Forms questionnaire (one respondent or a coordinated team response)", GT),
        ("GT secure tenant", "Responses exported from Forms in GT's Microsoft 365 environment", GT2),
        ("GT assessment engine", "Proprietary scoring, findings and roadmap logic - not shared", DARK),
        ("GT reviewer", "AI GRC professional reviews results and context", TEAL),
        ("Client", "Receives the PDF report by email; debrief session", CORAL)]
fw = (CW - 4 * 180000) // 5
for i, (h, d, col) in enumerate(flow):
    x = L + i * (fw + 180000)
    box(s, x, TOP - 50000, fw, 520000, fill=col, text=h, size=13, bold=True, color=WHITE)
    box(s, x, TOP + 470000, fw, 1350000, fill=LAV, text=d, size=10.5)
    if i < 4:
        box(s, x + fw + 20000, TOP + 900000, 140000, 300000, fill=GT2, shape=MSO_SHAPE.RIGHT_ARROW)
hw = (CW - 150000) // 2
box(s, L, TOP + 2050000, hw, 400000, fill=GT, text="Why we do not send the Excel engine", size=12, bold=True, color=WHITE)
txt(s, L, TOP + 2500000, hw, 1900000, ["Protects GT methodology, weightings and content", "Avoids answers being 'tuned' to the score", "Consistent, quality-reviewed results across clients",
                                       "No macros, formulas or printing issues on client devices"], size=11.5, bullet=True, spacing=7)
x2 = L + hw + 150000
box(s, x2, TOP + 2050000, hw, 400000, fill=TEAL, text="How we protect the client's answers", size=12, bold=True, color=WHITE)
txt(s, x2, TOP + 2500000, hw, 1900000, ["Stored only in GT's Microsoft 365 tenant; access limited to the engagement team", "No personal data beyond contact details; Bahrain PDPL-compliant handling",
                                        "Never entered into public AI tools", "Retained per engagement terms; deleted on request", "Report marked confidential; shared only with the named sponsor"], size=11.5, bullet=True, spacing=7)
add(s)

# =========================================================== 11 report sample
s = content("The Report - What the Client Receives", "Illustrative pages from a sample report (GT Bahrain). The full report runs to 7-9 pages.")
ih = 4300000
pic(s, "rep0.png", L, TOP - 100000, h=ih)
pic(s, "rep2.png", L + 3150000, TOP - 100000, h=ih)
pic(s, "rep4.png", L + 6300000, TOP - 100000, h=ih)
x2 = L + 9450000
sections = ["Executive summary & pathway", "Readiness by domain (radar)", "AI use & inherent risk profile", "Responsible-AI principle coverage", "Detailed findings & recommendations",
            "Top-5 priority actions", "Suggested roadmap", "Regulatory & standards considerations", "How GT can help; next steps & sign-off"]
txt(s, x2, TOP - 100000, R - x2, ih, [[("Report sections", True)]] + sections, size=10, spacing=5, color=BLACK)
add(s)

# =========================================================== 12 our position
s = content("We Assessed Ourselves First - GT Bahrain's Position", "GT Bahrain's AI Ready7 self-assessment (Sept 2026) compared with the GT AI Readiness & Risk Assessment v2.0 sample run for GT Bahrain (Oct 2026).")
cd = CategoryChartData(); cd.categories = DOMAINS
cd.add_series("AI Ready7 (Sept 2026)", [R7[d] for d in DOMAINS]); cd.add_series("Readiness Assessment v2.0 (Oct 2026)", [SAMPLE[d] for d in DOMAINS])
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Emu(L), Emu(TOP - 50000), Emu(int(CW * 0.58)), Emu(4300000), cd)
ch = gf.chart; ch.has_title = False; ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.BOTTOM; ch.legend.include_in_layout = False; ch.legend.font.size = Pt(10)
pl = ch.plots[0]; pl.gap_width = 70; pl.overlap = -10; pl.has_data_labels = True
pl.data_labels.number_format = '0"%"'; pl.data_labels.number_format_is_linked = False; pl.data_labels.font.size = Pt(10); pl.data_labels.font.bold = True
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
for ser, col in zip(pl.series, [LAV2, GT]):
    ser.format.fill.solid(); ser.format.fill.fore_color.rgb = col
ch.value_axis.maximum_scale = 100; ch.value_axis.minimum_scale = 0; ch.value_axis.major_unit = 25; ch.value_axis.tick_labels.font.size = Pt(9)
ch.value_axis.has_major_gridlines = True; ch.value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
ch.category_axis.tick_labels.font.size = Pt(10)
x2 = L + int(CW * 0.60); w2 = R - x2
rows = [["Domain", "Ready7", "v2.0", "Change"]] + [[d, f"{R7[d]}%", f"{SAMPLE[d]}%", f"{SAMPLE[d] - R7[d]:+d}"] for d in DOMAINS] + [["Overall", f"{READY7_OVERALL:.0f}%", f"{SAMPLE_OVERALL}%", f"{SAMPLE_OVERALL - round(READY7_OVERALL):+d}"]]
table(s, x2, TOP - 50000, w2, rows, [3.3, 1, 0.9, 1], size=10, rowh=360000)
txt(s, x2, TOP + 2550000, w2, 1800000, [[("Both place GT Bahrain at 'Informed'. ", True), ("The v2.0 assessment goes deeper on governance and risk: the five new questions (risk framework, appetite & KRIs, third-party "
     "and model risk) lower Governance, Risk & Compliance from 60% to 53% - exactly the gaps a lighter check would miss.", False)],
     [("Result: ", True), ("inherent risk tier Medium, exposure Moderate - 'Proceed with conditions'.", False)]], size=10.5, spacing=8)
txt(s, L, TOP + 4300000, CW, 300000, "Ready7 pillars consolidated into the five domains (averages). v2.0 figures are from the sample report included with this proposal.", size=8.5, color=GREY)
add(s)

# =========================================================== 13 divider
add(divider("Engagement Plan &\nDeliverables", 3))

# =========================================================== 14 options & timeline
s = content("Engagement Options and Timeline")
opts = [("ESSENTIAL", TEAL, "2-3 weeks", ["Kick-off call", "Online questionnaire", "GT analysis and quality review", "Confidential PDF report", "Short results call (30 min)"]),
        ("STANDARD", GT, "3-4 weeks", ["Everything in Essential", "Multi-respondent input (risk, IT, business)", "Evidence review of key answers", "Leadership debrief workshop (90 min)", "Agreed 90-day action plan"]),
        ("PLUS", DARK, "6-8 weeks", ["Everything in Standard", "AI inventory and risk-tiering of AI systems (GT AI Risk Assessment Framework)", "Control review against GT AI RCM v2.1 or ISO/IEC 42001 gap assessment",
                                    "AI security & agent audit checklist (selected areas)", "Board-ready summary"])]
ow = (CW - 2 * 150000) // 3
for i, (h, col, dur, items) in enumerate(opts):
    x = L + i * (ow + 150000)
    box(s, x, TOP - 250000, ow, 500000, fill=col, text=[[(h + "   ", True), (dur, False)]], size=14, color=WHITE)
    box(s, x, TOP + 250000, ow, 2600000, fill=LAV)
    txt(s, x + 70000, TOP + 350000, ow - 140000, 2500000, items, size=12, bullet=True, spacing=8)
box(s, L, TOP + 3000000, CW, 380000, fill=GT2, text="Indicative timeline (Standard)", size=12, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
tl = [("Day 0", "Kick-off"), ("Week 1-2", "Questionnaire open"), ("Week 2-3", "GT analysis & review"), ("Week 3", "Report issued"), ("Week 3-4", "Debrief & action plan")]
tw = (CW - 4 * 80000) // 5
for i, (w_, t) in enumerate(tl):
    box(s, L + i * (tw + 80000), TOP + 3450000, tw, 700000, fill=LAV, text=[[(w_, True)], [(t, False)]], size=11.5, color=GT)
txt(s, L, TOP + 4250000, CW, 400000, "Fees are proposed per option once scope (entities, AI systems, respondents) is confirmed. Re-assessment after 6-12 months shows progress on the same scale.", size=10, color=GREY)
add(s)

# =========================================================== 15 deliverables & client inputs
s = content("Deliverables, Client Inputs and Standards")
hw = (CW - 2 * 150000) // 3
cols = [("DELIVERABLES", GT, ["Confidential AI Readiness & Risk Assessment report (PDF)", "Readiness by domain and overall maturity", "Inherent AI risk tier, exposure and pathway",
                              "Findings, top-5 priorities and roadmap", "Regulatory considerations", "Debrief presentation (Standard / Plus)"]),
        ("FROM THE CLIENT", GT2, ["An executive sponsor and a coordinator", "Completed questionnaire (about 45 minutes)", "Optional evidence: AI policy, AI inventory, risk register, vendor contracts",
                                  "Availability for the debrief"]),
        ("STANDARDS & FRAMEWORKS", TEAL, ["GT AI Ready7 maturity model", "NSW AI Assessment Framework", "ISO/IEC 42001, ISO/IEC 23894, ISO/IEC 42005", "NIST AI RMF 1.0", "EU AI Act; Bahrain PDPL; CBB; iGA AI Policy",
                                         "ISACA Securing AI Agents (2026); OWASP LLM Top 10"])]
for i, (h, col, items) in enumerate(cols):
    x = L + i * (hw + 150000)
    box(s, x, TOP - 250000, hw, 420000, fill=col, text=h, size=13, bold=True, color=WHITE)
    box(s, x, TOP + 170000, hw, 3700000, fill=LAV)
    txt(s, x + 70000, TOP + 270000, hw - 140000, 3600000, items, size=12, bullet=True, spacing=9)
box(s, L, TOP + 4000000, CW, 600000, fill=DARK, text=[[("Natural next step: ", True), ("the GT AI Strategy & Roadmap service uses the Opportunity Discovery Toolkit to show WHERE AI fits - "
    "the readiness assessment shows HOW READY the organisation is to adopt it safely.", False)]], size=11.5, color=WHITE)
add(s)

ty = duplicate(dc.S_THANKS)
t = find(ty, "Grant Thornton Bahrain")
if t: set_text(t, "© 2026 Grant Thornton Bahrain. All rights reserved")
add(ty)
dc.finish(OUT)
