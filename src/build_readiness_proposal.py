# -*- coding: utf-8 -*-
"""GT AI Readiness & Risk Assessment - proposal deck (approach & methodology), built on the GT AI Service Lines template.

Usage: build_readiness_proposal.py TEMPLATE.pptx OUT.pptx IMGDIR
IMGDIR holds rendered pages of SAMPLE_Report_GT_Bahrain.pdf (rep0.png, rep2.png, rep4.png, rep0_top.png, rep0_domains.png);
GT template icons and the case-study visuals are in src/assets.
Market context slides use content from GT Bahrain 'Session 1: Introduction to Artificial Intelligence' (2026).
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
from tool_data import DOMAIN_MAP, PILLARS as DOMAINS, READINESS, CONTEXT, MATURITY, PATHWAY, PROFILE_Q, PILLAR_RISK

TEMPLATE, OUT, IMG = sys.argv[1], sys.argv[2], sys.argv[3]
dc.init(TEMPLATE)
add = dc.NEW.append
DCOL = [GT, CORAL, TEAL, GT2, DARK]
cur = {}
for p, v in READY7: cur.setdefault(DOMAIN_MAP[p], []).append(v)
R7 = {d: int(sum(cur[d]) / len(cur[d]) + 0.5) for d in DOMAINS}
NQ = {d: [q for q in READINESS if q["pillar"] == d] for d in DOMAINS}
NPROF = len(PROFILE_Q); NCTX = len(CONTEXT); NREAD = len(READINESS)
SRC_NOTE = "Source: McKinsey State of AI 2025 Survey; IDC Worldwide AI and Generative AI Spending Guide 2025 - as presented in GT Bahrain 'Introduction to Artificial Intelligence' (2026)."


def pic(s, name, x, y, w=None, h=None):
    p = s.shapes.add_picture(os.path.join(IMG, name), Emu(x), Emu(y), width=Emu(w) if w else None, height=Emu(h) if h else None)
    p.line.color.rgb = LAV2; p.line.width = Pt(0.75)
    return p


ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")


def icon(s, name, x, y, d, fill=None):
    """GT line icon (white) on an optional coloured disc."""
    if fill is not None:
        box(s, x, y, d, d, fill=fill, shape=MSO_SHAPE.OVAL)
    return s.shapes.add_picture(os.path.join(ASSETS, name + ".png"), Emu(x), Emu(y), width=Emu(d), height=Emu(d))


def asset(s, name, x, y, w=None, h=None):
    return s.shapes.add_picture(os.path.join(ASSETS, name + ".png"), Emu(x), Emu(y), width=Emu(w) if w else None, height=Emu(h) if h else None)


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


# =========================================================== 1 cover / 2 agenda
cov = duplicate(dc.S_COVER)
set_text(find(cov, "Artificial Intelligence"), "AI Readiness &\nRisk Assessment")
set_text(find(cov, "We go beyond"), "Proposal  |  October 2026")
add(cov)
add(agenda(["Why AI Readiness Matters Now", "Approach & Methodology", "Engagement Plan & Deliverables"]))

# =========================================================== 3 the shift has happened
s = content("The Shift Has Already Happened", "AI has moved from experiment to infrastructure. The competitive question in 2026 isn't whether your peers are using it - it's how far ahead they already are.")
stats = [("78-88%", "Enterprise AI adoption", "of organisations report using AI in at least one business function", "ic_people", GT),
         ("$407B", "Worldwide AI spending", "34.8% annual growth in AI and GenAI spending", "ic_chip", GT2),
         ("$15.7T", "Global economic impact", "potential contribution of AI to the global economy by 2030", "ic_ai", TEAL)]
tw = (CW - 2 * 150000) // 3; th_ = 1900000
for i, (big, h, d, ic, col) in enumerate(stats):
    x = L + i * (tw + 150000); y = TOP - 100000
    box(s, x, y, tw, th_, fill=col)
    icon(s, ic, x + 150000, y + 170000, 620000)
    txt(s, x + 850000, y + 120000, tw - 900000, 700000, [[(big, True)]], size=36, color=WHITE)
    txt(s, x + 850000, y + 800000, tw - 900000, 380000, [[(h, True)]], size=14, color=WHITE)
    txt(s, x + 150000, y + 1250000, tw - 300000, 600000, d, size=11.5, color=WHITE)
y2 = TOP + 1950000; hw = (CW - 150000) // 2
box(s, L, y2, hw, 380000, fill=DARK, text="What it means for leadership teams", size=12, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
txt(s, L, y2 + 450000, hw, 1700000, ["AI is already inside the organisation - in staff tools, purchased software and suppliers' services",
                                      "Value goes to organisations that scale AI safely, not those that run the most pilots",
                                      "Boards, regulators and clients increasingly ask for evidence that AI is governed"], size=12.5, bullet=True, spacing=12)
x2 = L + hw + 150000
box(s, x2, y2, hw, 380000, fill=CORAL, text="In Bahrain", size=12, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
txt(s, x2, y2 + 430000, hw, 1500000, ["Economic Vision 2030 positions AI as a lever for productivity and diversification",
                                       "iGA General Policy for the Use of AI (2025) sets expectations for compliance, adoption and awareness",
                                       "Bahrain PDPL applies whenever AI processes personal data; sector regulators (CBB, NHRA, TRA) expect controls"], size=12.5, bullet=True, spacing=12)
txt(s, L, TOP + 4450000, CW, 300000, SRC_NOTE, size=8.5, color=GREY)
add(s)

# =========================================================== 4 what AI unlocks
s = content("What AI Unlocks", "AI delivers measurable outcomes across the organisation - when it is adopted with the right foundations.")
unl = [("Automate", "-27%", "operational costs", "JPM Coin turns deposits into on-chain, programmable, always-on digital money.", "ex_automate", "ic_gears"),
       ("Analyze", "-31%", "downtime", "Shell scales AI-fuelled predictive maintenance to 10,000 assets.", "ex_analyze", "ic_search"),
       ("Generate", "+26-60%", "developer productivity", "AI coding assistants accelerate software delivery.", "ex_generate", "ic_bulb"),
       ("Predict", "22%", "efficiency gains", "UPS ORION optimises delivery routes with advanced analytics.", "ex_predict", "ic_data"),
       ("Assist", "<2 min", "resolution time  |  +$40M profit", "Klarna's AI assistant handled two-thirds of customer-service chats in its first month.", "ex_assist", "ic_person")]
uw = (CW - 4 * 110000) // 5
for i, (h, big, unit, ex, img, ic) in enumerate(unl):
    x = L + i * (uw + 110000); col = [GT, GT2, TEAL, DARK, CORAL][i]
    box(s, x, TOP - 100000, uw, 560000, fill=col)
    icon(s, ic, x + 70000, TOP - 50000, 460000)
    txt(s, x + 560000, TOP - 20000, uw - 600000, 420000, [[(h, True)]], size=17, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    box(s, x, TOP + 460000, uw, 1380000, fill=LAV)
    asset(s, img, x + 40000, TOP + 500000, w=uw - 80000, h=1300000)
    txt(s, x, TOP + 1880000, uw, 560000, [[(big, True)]], size=28 if len(big) < 7 else 22, color=col, align=PP_ALIGN.CENTER)
    txt(s, x, TOP + 2420000, uw, 330000, unit, size=11, bold=True, color=GREY, align=PP_ALIGN.CENTER)
    box(s, x, TOP + 2780000, uw, 950000, fill=WHITE, line=LAV2, text=ex, size=11.5, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
box(s, L, TOP + 3830000, CW, 500000, fill=DARK, text=[[("The same capabilities create new risks ", True), ("- to data, decisions, customers and compliance - unless governance keeps pace. That is what this assessment measures.", False)]],
    size=12, color=WHITE)
txt(s, L, TOP + 4400000, CW, 300000, SRC_NOTE, size=8.5, color=GREY)
add(s)

# =========================================================== 5 adoption maturity curve
s = content("The AI Adoption Maturity Curve", "Organisations move from awareness to governed, enterprise-wide AI - and the governance required rises at every stage. Example: JPMorgan Chase LLM Suite (2023-2025).")
stages = [("Awareness", "Late 2022-2023", "Pilots & hackathons", "Acceptable-use policy; AI literacy; ban on client data in public tools"),
          ("Use-case discovery", "Early-mid 2024", "Initial rollout to 60,000 employees, usage tracked", "AI inventory; use-case intake; owner for every tool"),
          ("Controlled adoption", "Late 2024", "Opt-in scaling to 140,000-200,000 employees in 8 months, proving ROI", "Risk assessment by tier; human oversight; vendor due diligence"),
          ("Enterprise enablement", "Late 2025", "Multi-model, firm-wide use by 250,000+ employees", "AI committee; monitoring & KRIs; secure architecture; model validation"),
          ("Governance maturity", "2025 onward", "Oversight & controls embedded", "Board oversight; assurance; ISO/IEC 42001-aligned management system")]
sw_ = (CW - 4 * 90000) // 5; base = TOP + 2350000
for i, (st, when, d, gov) in enumerate(stages):
    x = L + i * (sw_ + 90000); col = [LAV2, GT2, TEAL, GT, DARK][i]
    hgt = 700000 + i * 330000
    box(s, x, base - hgt, sw_, hgt, fill=col)
    txt(s, x + 50000, base - hgt + 60000, sw_ - 100000, 420000, [[(f"{i + 1}  {st}", True)]], size=12.5, color=GT if i == 0 else WHITE)
    txt(s, x + 50000, base - hgt + 430000, sw_ - 100000, hgt - 450000, [[(when + ": ", True), (d, False)]], size=9.5, color=BLACK if i == 0 else WHITE)
    box(s, x, base + 80000, sw_, 1050000, fill=LAV, text=gov, size=10, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
box(s, L - 10000, base + 1170000, CW, 60000, fill=GT)
txt(s, L, base + 1230000, CW, 280000, "Governance focus needed at each stage (what the assessment checks)", size=9.5, bold=True, color=GT)
box(s, L, TOP + 4080000, CW, 520000, fill=DARK, text=[[("Where are you on the curve? ", True), ("The assessment places the organisation on this journey and shows what must be in place - governance, risk, skills, strategy and data - to move up safely.", False)]],
    size=11, color=WHITE)
txt(s, L, TOP + 4630000, CW, 250000, "Source: GT Bahrain 'Introduction to Artificial Intelligence' (2026), based on public JPMorgan Chase disclosures.", size=8, color=GREY)
add(s)

# =========================================================== 6 moving fast
s = content("Why Organisations Are Moving Fast", "The commercial pressure for AI adoption is becoming difficult to ignore - but adoption is moving faster than governance.")
import math
pres = ["Faster decision cycles", "Productivity pressure", "Cost optimisation", "Customer expectations", "Data volume growth", "Competitive pressure"]
cx, cy = L + 2550000, TOP + 1350000
box(s, cx - 620000, cy - 360000, 1240000, 720000, fill=LAV2, text="AI adoption", size=13, bold=True, color=GT, shape=MSO_SHAPE.HEXAGON)
for i, p in enumerate(pres):
    a = math.radians(-90 + i * 60)
    x = int(cx + 1750000 * math.cos(a)) - 700000; y = int(cy + 1080000 * math.sin(a)) - 300000
    box(s, x, y, 1400000, 600000, fill=[GT, GT2, TEAL, DARK, CORAL, GT2][i], text=p, size=10.5, bold=True, color=WHITE, shape=MSO_SHAPE.HEXAGON)
x2 = L + 5500000; w2 = R - x2
box(s, x2, TOP - 100000, w2, 460000, fill=CORAL, text="But adoption is moving faster than governance.", size=14, bold=True, color=WHITE)
box(s, x2, TOP + 420000, w2, 1150000, fill=LAV, text=[[("Example: ", True), ("Morgan Stanley deployed Microsoft 365 Copilot and AI assistants to thousands of financial advisors - less time searching "
    "internal documents, more time serving clients and faster responses.", False)]], size=11, align=PP_ALIGN.LEFT)
box(s, x2, TOP + 1650000, w2, 1050000, fill=WHITE, line=GT, text=[[("“Bahrain's National AI Strategy promotes embedding AI in finance and operations while requiring governance to support economic "
    "diversification and regulatory compliance.”", False)], [("Information & eGovernment Authority (iGA), Bahrain AI Report, 2025", True)]], size=10.5, color=GT, align=PP_ALIGN.LEFT)
box(s, L, TOP + 2850000, CW, 360000, fill=DARK, text="When governance lags behind adoption", size=12, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
risks = [("ic_eye", "Shadow AI & data leakage", "Staff paste confidential or personal data into public AI tools."),
         ("ic_risk", "Unreliable or biased outputs", "Errors and bias reach customers and decisions unchecked."),
         ("ic_doccheck", "Regulatory exposure", "PDPL, sector-regulator and EU AI Act obligations missed."),
         ("ic_ai", "AI-enabled fraud", "Deepfake payment requests and AI-crafted phishing succeed.")]
rw = (CW - 3 * 110000) // 4
for i, (ic, h, d) in enumerate(risks):
    x = L + i * (rw + 110000)
    box(s, x, TOP + 3260000, rw, 1300000, fill=LAV)
    icon(s, ic, x + 70000, TOP + 3340000, 520000, fill=[GT, CORAL, TEAL, DARK][i])
    txt(s, x + 660000, TOP + 3300000, rw - 700000, 600000, [[(h, True)]], size=12, color=GT, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + 90000, TOP + 3920000, rw - 180000, 620000, d, size=10.5)
add(s)

# =========================================================== 7 why assess now
s = content("Why Assess AI Readiness and Risk Now", "AI is already in the organisation - in staff tools, in software and in suppliers' services. Boards and regulators now expect it to be governed.")
why = [("Regulation is arriving", "EU AI Act obligations phase in to 2026-27 (AI literacy since Feb 2025); Bahrain PDPL applies to AI processing personal data; iGA AI Policy (2025); CBB expectations for licensees; draft Bahrain AI law.", "ic_doccheck"),
       ("Shadow AI and embedded AI", "Staff use public GenAI tools; vendors switch on AI features in existing software. Without an inventory, risk is invisible.", "ic_eye"),
       ("Agentic AI raises the stakes", "AI that takes actions through tools and APIs introduces prompt injection, excessive agency and new fraud paths (ISACA, 2026).", "ic_gears"),
       ("AI-enabled threats", "Deepfake payment fraud, AI-crafted phishing and automated attacks target every organisation - whether or not it uses AI itself.", "ic_risk"),
       ("Boards are asking", "Who owns AI? What is our risk appetite? Which AI do we use, with what data, and who checks it?", "ic_people"),
       ("Investment needs a baseline", "A measured starting point focuses budgets on the gaps that matter and shows progress over time.", "ic_search")]
cw = (CW - 2 * 150000) // 3; chh = 2020000
for i, (h, d, ic) in enumerate(why):
    x = L + (i % 3) * (cw + 150000); y = TOP - 100000 + (i // 3) * (chh + 130000)
    col = [GT, GT2, TEAL, CORAL, DARK, GREEN][i]
    box(s, x, y, cw, chh, fill=LAV); box(s, x, y, cw, 620000, fill=col)
    icon(s, ic, x + 90000, y + 70000, 480000)
    txt(s, x + 650000, y + 60000, cw - 700000, 500000, [[(h, True)]], size=14, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + 90000, y + 700000, cw - 180000, chh - 750000, d, size=11.5)
add(s)

# =========================================================== 8 what it answers
s = content("What the Assessment Answers", "Four answers in one report - about 45 minutes of the client's time, delivered within 5 working days.")
ans = [("How ready are we?", "Readiness score and maturity (Partial / Informed / Repeatable / Adaptive) overall and for each of the five domains.", GT, "ic_search"),
       ("How risky is our AI?", "Inherent AI risk tier (Low to Very high) from how AI is used, the decisions it influences, the data it touches and the oversight in place.", GT2, "ic_risk"),
       ("Is our readiness enough for our risk?", "Overall AI risk exposure and a recommended pathway: proceed, proceed with conditions, remediate before scaling, or pause.", TEAL, "ic_eye"),
       ("What should we do first?", "Findings per question, top-5 priority actions, a phased roadmap, regulatory considerations and next steps.", CORAL, "ic_doccheck")]
lw_ = int(CW * 0.6)
for i, (h, d, col, ic) in enumerate(ans):
    y = TOP - 100000 + i * 1130000
    box(s, L, y, 1000000, 1030000, fill=col)
    icon(s, ic, L + 220000, y + 200000, 560000)
    box(s, L + 1000000, y, 2300000, 1030000, fill=LAV, text=h, size=14, bold=True, color=GT, align=PP_ALIGN.LEFT)
    box(s, L + 3300000, y, lw_ - 3300000, 1030000, fill=WHITE, line=LAV2, text=d, size=11.5, align=PP_ALIGN.LEFT)
x2 = L + lw_ + 150000; w2 = R - x2
asset_img = s.shapes.add_picture(os.path.join(IMG, "rep0_domains.png"), Emu(x2), Emu(TOP + 300000), width=Emu(w2))
asset_img.line.color.rgb = LAV2
txt(s, x2, TOP - 100000, w2, 380000, [[("Readiness by domain - as shown in the report", True)]], size=11, color=GT, align=PP_ALIGN.CENTER)
txt(s, x2, TOP + 300000 + int(w2 * 0.56) + 60000, w2, 300000, "Illustrative sample report", size=9, color=GREY, align=PP_ALIGN.CENTER)
add(s)

# =========================================================== 9 divider / 10 approach
add(divider("Approach &\nMethodology", 2))
s = content("Our Approach - Report Within 5 Working Days", "Light-touch for the client, rigorous behind the scenes. The scoring engine stays with GT - the client receives a clear, independent report.")
w = chevrons(s, TOP - 100000, ["Day 1  Kick-off", "Day 1-2  Questionnaire", "Day 3-4  GT analysis", "Day 5  Report", "Day 5  Debrief"], h=600000, size=12.5)
det = [("ic_people", ["30-minute call with the sponsor", "Scope: whole organisation, unit or AI system", "Respondents agreed; link sent the same day"], "30 min", "Scope confirmed"),
       ("ic_laptop", ["Secure Microsoft Forms questionnaire", f"{NPROF + NCTX + NREAD} questions: profile, AI use, readiness", "Plain-language options"], "~45 min", "Completed responses"),
       ("ic_gears", ["Responses processed in GT's proprietary engine", "Risk tier, maturity, exposure, findings, roadmap", "GT AI GRC professional review"], "None", "Quality-reviewed results"),
       ("ic_docs", ["Confidential PDF report (7-9 pages)", "Emailed to the sponsor", "Headline results in the email"], "None", "Report"),
       ("ic_bulb", ["60-90 minute leadership debrief", "Priorities and owners agreed", "Follow-on options discussed"], "60-90 min", "90-day action plan")]
for i, (ic, items, eff, out) in enumerate(det):
    x = L + i * (w - 60000) + 30000; ww = w - 120000
    box(s, x, TOP + 600000, ww, 2050000, fill=LAV)
    icon(s, ic, x + (ww - 520000) // 2, TOP + 660000, 520000, fill=[GT, GT2, TEAL, DARK, CORAL][i])
    txt(s, x + 30000, TOP + 1240000, ww - 60000, 1400000, items, size=10, bullet=True, spacing=5)
    box(s, x, TOP + 2720000, ww, 520000, fill=WHITE, line=LAV2, text=[[("Client time: ", True), (eff, False)]], size=10.5, color=GT)
    box(s, x, TOP + 3290000, ww, 520000, fill=LAV2, text=[[("Output: ", True), (out, False)]], size=10.5, color=GT)
box(s, L, TOP + 3920000, CW, 450000, fill=DARK, text="Timeline assumes the questionnaire is completed by Day 2; otherwise the report is issued within 3 working days of receiving complete responses.", size=11, color=WHITE)
add(s)

# =========================================================== 11 five domains
s = content("Five Assessment Domains", f"{NREAD} readiness questions across five equally weighted domains, aligned to ISO/IEC 42001, ISO/IEC 23894, NIST AI RMF and the NSW AI Assessment Framework.")
DIC = {"AI Governance": "ic_org", "AI Risk & Security": "ic_risk", "Training & Awareness": "ic_person", "Strategy & Value": "ic_bulb", "Data & Technology": "ic_data"}
dw = (CW - 4 * 100000) // 5
for i, d in enumerate(DOMAINS):
    x = L + i * (dw + 100000); col = DCOL[i]
    box(s, x, TOP - 100000, dw, 1000000, fill=col)
    icon(s, DIC[d], x + (dw - 480000) // 2, TOP - 50000, 480000)
    txt(s, x, TOP + 450000, dw, 420000, [[(d, True)]], size=12.5, color=WHITE, align=PP_ALIGN.CENTER)
    box(s, x, TOP + 900000, dw, 280000, fill=LAV2, text=f"{len(NQ[d])} question{'s' if len(NQ[d]) > 1 else ''}", size=10, bold=True, color=GT)
    box(s, x, TOP + 1180000, dw, 1650000, fill=LAV)
    txt(s, x + 30000, TOP + 1230000, dw - 60000, 1600000, [q["topic"] for q in NQ[d]], size=9.5, bullet=True, spacing=3)
    box(s, x, TOP + 2880000, dw, 1350000, fill=WHITE, line=col, text=[[("If weak: ", True), (PILLAR_RISK[d], False)]], size=9, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
box(s, L, TOP + 4300000, CW, 400000, fill=DARK, text="Each question maps to findings, recommendations and GT AI RCM v2.1 controls - so results convert directly into an action plan.", size=11, color=WHITE)
add(s)

# =========================================================== 12 inherent risk
s = content("Inherent AI Risk Profile (NSW AIAF-based)", "Four questions establish how much risk the organisation's AI use carries - before considering its controls.")
CIC = ["ic_ai", "ic_people", "ic_data", "ic_eye"]
for i, c in enumerate(CONTEXT):
    y = TOP - 100000 + i * 830000
    box(s, L, y, 2900000, 750000, fill=[GT, GT2, TEAL, DARK][i])
    icon(s, CIC[i], L + 80000, y + 125000, 500000)
    txt(s, L + 650000, y, 2200000, 750000, [[(c["topic"], True)]], size=13, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    box(s, L + 2900000, y, 4000000, 750000, fill=LAV, text=c["q"], size=10.5, align=PP_ALIGN.LEFT)
x2 = L + 7100000; w2 = R - x2
box(s, x2, TOP - 100000, w2, 380000, fill=GT, text="Inherent risk tier", size=12, bold=True, color=WHITE)
for i, (t, col, ex) in enumerate([("Low", RGBColor(0xC6, 0xE0, 0xB4), "Internal productivity tools, human checks"), ("Medium", RGBColor(0xFF, 0xE6, 0x99), "Customer-facing or confidential data"),
                                  ("High", RGBColor(0xF4, 0xB1, 0x83), "Decisions affecting people or sensitive data"), ("Very high", RGBColor(0xE0, 0x66, 0x66), "Significant decisions with little oversight")]):
    box(s, x2, TOP + 330000 + i * 520000, w2, 480000, fill=col, text=[[(t + "  ", True), (ex, False)]], size=10.5, align=PP_ALIGN.LEFT)
txt(s, x2, TOP + 2450000, w2, 800000, "Overrides apply - AI that makes or materially influences decisions about individuals' rights or access to services cannot be rated Low.", size=10, color=GREY)
box(s, L, TOP + 3350000, CW, 650000, fill=LAV2, text=[[("Exposure = inherent risk tier x readiness maturity. ", True), ("A mature organisation with high-risk AI can be at moderate exposure; "
    "an immature organisation with the same AI is at critical exposure. The pathway follows from the exposure.", False)]], size=11.5, align=PP_ALIGN.LEFT)
txt(s, L, TOP + 4100000, CW, 400000, "Basis: NSW AI Assessment Framework; EU AI Act risk categories; ISO/IEC 23894 and ISO/IEC 42005.", size=9, color=GREY)
add(s)

# =========================================================== 13 results
s = content("How Results Are Expressed", "Clear, comparable results for every client; the detailed scoring rules remain GT's proprietary methodology.")
lw_ = int(CW * 0.6)
box(s, L, TOP - 100000, lw_, 340000, fill=GT, text="1  Every answer on a five-level scale", size=11.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
lw5 = (lw_ - 4 * 60000) // 5
for i, t in enumerate(["Not in place", "Initial / ad hoc", "Defined", "Implemented", "Optimised"]):
    box(s, L + i * (lw5 + 60000), TOP + 290000, lw5, 520000, fill=[RGBColor(0xF8, 0xCB, 0xAD), RGBColor(0xFC, 0xE4, 0xD6), RGBColor(0xFF, 0xE6, 0x99), RGBColor(0xC6, 0xE0, 0xB4), RGBColor(0x9B, 0xC2, 0xE6)][i],
        text=[[(f"Level {i}", True)], [(t, False)]], size=10)
box(s, L, TOP + 900000, lw_, 340000, fill=GT2, text="2  Domain and overall scores placed in four maturity bands", size=11.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
bw = (lw_ - 3 * 60000) // 4
for i, ((b, rng), (_, _, desc)) in enumerate(zip([("Partial", "< 50%"), ("Informed", "50 - 70%"), ("Repeatable", "70 - 90%"), ("Adaptive", "90%+")], MATURITY)):
    box(s, L + i * (bw + 60000), TOP + 1290000, bw, 1150000, fill=[RGBColor(0xF8, 0xCB, 0xAD), RGBColor(0xFF, 0xE6, 0x99), RGBColor(0xC6, 0xE0, 0xB4), RGBColor(0x9B, 0xC2, 0xE6)][i],
        text=[[(f"{b}  {rng}", True)], [(desc, False)]], size=9.5)
box(s, L, TOP + 2530000, lw_, 340000, fill=TEAL, text="3  Exposure (risk tier x maturity) sets the recommended pathway", size=11.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
for i, (exp, path, desc) in enumerate(PATHWAY):
    box(s, L + i * (bw + 60000), TOP + 2920000, bw, 900000, fill=LAV, text=[[(f"{exp} exposure", True)], [({"Low": "Proceed - self-managed", "Moderate": "Proceed with conditions", "High": "Remediate before scaling", "Critical": "Pause high-impact AI and escalate"}.get(exp, path), False)]], size=10, color=GT)
txt(s, L, TOP + 3900000, lw_, 450000, "Domains are equally weighted. Findings, risks, controls (GT AI RCM v2.1) and recommendations are generated by the GT engine and reviewed by a GT professional.", size=9.5, color=GREY)
x2 = L + lw_ + 150000; w2 = R - x2
im = s.shapes.add_picture(os.path.join(IMG, "rep0_top.png"), Emu(x2), Emu(TOP - 100000), width=Emu(w2)); im.line.color.rgb = LAV2
txt(s, x2, TOP + 3850000, w2, 300000, "How it looks in the report (illustrative sample)", size=9, color=GREY, align=PP_ALIGN.CENTER)
add(s)

# =========================================================== 14 collection & confidentiality
s = content("How Answers Are Collected - and Why the Engine Stays with GT", "The client never handles spreadsheets or formulas. Responses flow securely from a form to a reviewed PDF report.")
flow = [("Client", "Completes the secure Microsoft Forms questionnaire (one respondent or a coordinated team response)", GT, "ic_laptop"),
        ("GT secure tenant", "Responses exported from Forms in GT's Microsoft 365 environment", GT2, "ic_data"),
        ("GT assessment engine", "Proprietary scoring, findings and roadmap logic - not shared", DARK, "ic_gears"),
        ("GT reviewer", "AI GRC professional reviews results and context", TEAL, "ic_search"),
        ("Client", "Receives the PDF report by email; debrief session", CORAL, "ic_doccheck")]
fw = (CW - 4 * 180000) // 5
for i, (h, d, col, ic) in enumerate(flow):
    x = L + i * (fw + 180000)
    box(s, x, TOP - 100000, fw, 1000000, fill=col)
    icon(s, ic, x + (fw - 500000) // 2, TOP - 40000, 500000)
    txt(s, x, TOP + 480000, fw, 380000, [[(h, True)]], size=12.5, color=WHITE, align=PP_ALIGN.CENTER)
    box(s, x, TOP + 900000, fw, 1000000, fill=LAV, text=d, size=10.5)
    if i < 4: box(s, x + fw + 20000, TOP + 300000, 140000, 300000, fill=GT2, shape=MSO_SHAPE.RIGHT_ARROW)
hw = (CW - 150000) // 2
box(s, L, TOP + 2050000, hw, 400000, fill=GT, text="Why we do not send the Excel engine", size=12, bold=True, color=WHITE)
txt(s, L, TOP + 2500000, hw, 1900000, ["Protects GT methodology, weightings and content", "Avoids answers being 'tuned' to the score", "Consistent, quality-reviewed results across clients",
                                       "No macros, formulas or printing issues on client devices"], size=11.5, bullet=True, spacing=7)
x2 = L + hw + 150000
box(s, x2, TOP + 2050000, hw, 400000, fill=TEAL, text="How we protect the client's answers", size=12, bold=True, color=WHITE)
txt(s, x2, TOP + 2500000, hw, 1900000, ["Stored only in GT's Microsoft 365 tenant; access limited to the engagement team", "No personal data beyond contact details; Bahrain PDPL-compliant handling",
                                        "Never entered into public AI tools", "Retained per engagement terms; deleted on request", "Report marked confidential; shared only with the named sponsor"], size=11.5, bullet=True, spacing=7)
add(s)

# =========================================================== 15 report sample
s = content("The Report - What the Client Receives", "Pages from an illustrative sample report (sample answers). The full report runs to 7-9 pages.")
ih = 4300000
pic(s, "rep0.png", L, TOP - 100000, h=ih); pic(s, "rep2.png", L + 3150000, TOP - 100000, h=ih); pic(s, "rep4.png", L + 6300000, TOP - 100000, h=ih)
for x, t in [(L, "Executive summary & domains"), (L + 3150000, "Findings & recommendations"), (L + 6300000, "Roadmap")]:
    txt(s, x, TOP + ih - 80000, 3000000, 300000, [[(t, True)]], size=10, color=GT)
x2 = L + 9450000
sections = ["Executive summary & pathway", "Readiness by domain (radar)", "AI use & inherent risk profile", "Responsible-AI principle coverage", "Detailed findings & recommendations",
            "Top-5 priority actions", "Suggested roadmap", "Regulatory & standards considerations", "How GT can help; next steps & sign-off"]
txt(s, x2, TOP - 100000, R - x2, ih, [[("Report sections", True)]] + sections, size=10, spacing=5, color=BLACK)
add(s)

# =========================================================== 16 GT's own position (actual Ready7 only)
s = content("Our Own Starting Point - GT Bahrain AI Ready7 (Sept 2026)", "We assessed ourselves first. GT Bahrain's AI Ready7 results by pillar, and consolidated into the five assessment domains.")
cd = CategoryChartData(); cd.categories = [p for p, _ in READY7]
cd.add_series("AI Ready7 score (%)", [v for _, v in READY7])
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Emu(L), Emu(TOP - 50000), Emu(int(CW * 0.52)), Emu(4200000), cd)
ch = gf.chart; ch.has_legend = False; ch.has_title = False
pl = ch.plots[0]; pl.gap_width = 60; pl.has_data_labels = True
pl.data_labels.number_format = '0"%"'; pl.data_labels.number_format_is_linked = False; pl.data_labels.font.size = Pt(10); pl.data_labels.font.bold = True
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
ser = pl.series[0]
for idx, (_, v) in enumerate(READY7):
    pt = ser.points[idx]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = GT if v >= 60 else CORAL
ch.category_axis.reverse_order = True
ch.value_axis.maximum_scale = 100; ch.value_axis.minimum_scale = 0; ch.value_axis.major_unit = 25; ch.value_axis.tick_labels.font.size = Pt(9)
ch.category_axis.tick_labels.font.size = Pt(10); ch.value_axis.has_major_gridlines = True; ch.value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
x2 = L + int(CW * 0.55); w2 = R - x2
rows = [["Assessment domain", "Compared Ready7 pillar(s)", "Score"]] + [[d, " + ".join(p for p, _ in READY7 if DOMAIN_MAP[p] == d), f"{R7[d]}%"] for d in DOMAINS] + \
       [["Overall", "Readiness status: Informed", f"{READY7_OVERALL:.1f}%"]]
fills = {(i, 2): RGBColor(0xF8, 0xCB, 0xAD) if R7[d] < 60 else RGBColor(0xFF, 0xE6, 0x99) for i, d in enumerate(DOMAINS, 1)}
table(s, x2, TOP - 50000, w2, rows, [2.4, 3.4, 0.8], size=9.5, rowh=420000, fills=fills)
txt(s, x2, TOP + 3050000, w2, 1200000, [[("Our priorities: ", True), ("Strategy & Value (50%) and AI Governance (60%) - including lifecycle AI risk assessment (Ready7 2.4) and data classification for AI (5.1), "
     "both 'Not Met'. The AI Readiness & Risk Assessment adds deeper governance and risk-management questions to measure exactly these areas.", False)]], size=10.5)
txt(s, L, TOP + 4250000, CW, 300000, "Source: GT Bahrain AI Ready7 Result Report (Sept 2026). Domain = average of its Ready7 pillars. Coral = below 60%.", size=8.5, color=GREY)
add(s)

# =========================================================== 17 divider / 18 options & timeline
add(divider("Engagement Plan &\nDeliverables", 3))
s = content("Engagement Options - Report Within 5 Working Days")
opts = [("ESSENTIAL", TEAL, "ic_search", "Best for: a fast, independent baseline", ["Kick-off call", "Online questionnaire", "GT analysis and quality review", "Confidential PDF report", "30-minute results call"]),
        ("STANDARD", GT, "ic_people", "Best for: leadership alignment and an agreed plan", ["Everything in Essential", "Multi-respondent input (risk, IT, business)", "Review of key evidence provided with the questionnaire",
                                                                                         "Leadership debrief workshop (90 min)", "Agreed 90-day action plan"]),
        ("PLUS", DARK, "ic_doccheck", "Best for: regulated organisations and boards", ["Everything in Standard", "AI inventory risk-tiering of up to 10 AI systems", "Board-ready summary",
                                                                                     "Follow-on (scoped separately): RCM v2.1 control review, ISO/IEC 42001 gap assessment, AI security & agent audit"])]
ow = (CW - 2 * 150000) // 3
for i, (h, col, ic, best, items) in enumerate(opts):
    x = L + i * (ow + 150000)
    box(s, x, TOP - 250000, ow, 640000, fill=col)
    icon(s, ic, x + 90000, TOP - 190000, 520000)
    txt(s, x + 700000, TOP - 250000, ow - 750000, 640000, [[(h, True)], [("report by Day 5", False)]], size=14, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, spacing=0)
    box(s, x, TOP + 390000, ow, 400000, fill=LAV2, text=best, size=10.5, bold=True, color=GT)
    box(s, x, TOP + 790000, ow, 2050000, fill=LAV)
    txt(s, x + 70000, TOP + 870000, ow - 140000, 1950000, items, size=11.5, bullet=True, spacing=7)
box(s, L, TOP + 3000000, CW, 380000, fill=GT2, text="Timeline (all options)", size=12, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
tl = [("Day 1", "Kick-off; questionnaire sent"), ("Day 1-2", "Client completes questionnaire"), ("Day 3-4", "GT analysis & quality review"), ("Day 5", "Report issued"), ("Day 5", "Debrief & action plan")]
tw = (CW - 4 * 80000) // 5
for i, (w_, t) in enumerate(tl):
    box(s, L + i * (tw + 80000), TOP + 3450000, tw, 700000, fill=LAV, text=[[(w_, True)], [(t, False)]], size=11.5, color=GT)
txt(s, L, TOP + 4250000, CW, 400000, "Fees are proposed per option once scope (entities, AI systems, respondents) is confirmed. Re-assessment after 6-12 months shows progress on the same scale.", size=10, color=GREY)
add(s)

# =========================================================== 19 deliverables
s = content("Deliverables, Client Inputs and Standards")
hw = (CW - 2 * 150000) // 3
cols = [("DELIVERABLES", GT, "ic_docs", ["Confidential AI Readiness & Risk Assessment report (PDF) by Day 5", "Readiness by domain and overall maturity", "Inherent AI risk tier, exposure and pathway",
                                          "Findings, top-5 priorities and roadmap", "Regulatory considerations", "Debrief presentation (Standard / Plus)"]),
        ("FROM THE CLIENT", GT2, "ic_people", ["An executive sponsor and a coordinator", "Completed questionnaire by Day 2 (about 45 minutes)", "Optional evidence: AI policy, AI inventory, risk register, vendor contracts",
                                                "Availability for the debrief"]),
        ("STANDARDS & FRAMEWORKS", TEAL, "ic_doccheck", ["NSW AI Assessment Framework", "ISO/IEC 42001, ISO/IEC 23894, ISO/IEC 42005", "NIST AI RMF 1.0", "EU AI Act; Bahrain PDPL; CBB; iGA AI Policy",
                                                         "ISACA Securing AI Agents (2026); OWASP LLM Top 10", "GT AI RCM v2.1 (81 controls)"])]
for i, (h, col, ic, items) in enumerate(cols):
    x = L + i * (hw + 150000)
    box(s, x, TOP - 250000, hw, 620000, fill=col)
    icon(s, ic, x + 90000, TOP - 200000, 520000)
    txt(s, x + 700000, TOP - 250000, hw - 750000, 620000, [[(h, True)]], size=13, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    box(s, x, TOP + 370000, hw, 2300000, fill=LAV)
    txt(s, x + 70000, TOP + 440000, hw - 140000, 2200000, items, size=11.5, bullet=True, spacing=6)
box(s, L, TOP + 2760000, CW, 330000, fill=GT2, text="What the report contains", size=11.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
secs = ["Executive summary & pathway", "Readiness by domain (radar)", "AI use & inherent risk", "Responsible-AI principles", "Findings & recommendations",
        "Top-5 priority actions", "Implementation roadmap", "Regulatory considerations", "How GT can help & sign-off"]
cw9 = (CW - 8 * 50000) // 9
for i, t in enumerate(secs):
    box(s, L + i * (cw9 + 50000), TOP + 3130000, cw9, 760000, fill=WHITE, line=GT2, text=[[(str(i + 1), True)], [(t, False)]], size=9, color=GT)
box(s, L, TOP + 4000000, CW, 600000, fill=DARK, text=[[("Natural next step: ", True), ("the GT AI Strategy & Roadmap service uses the Opportunity Discovery Toolkit to show WHERE AI fits - "
    "the readiness assessment shows HOW READY the organisation is to adopt it safely. Both can run together within the same 5 working days.", False)]], size=11.5, color=WHITE)
add(s)

ty = duplicate(dc.S_THANKS)
t = find(ty, "Grant Thornton Bahrain")
if t: set_text(t, "© 2026 Grant Thornton Bahrain. All rights reserved")
add(ty)
dc.finish(OUT)
