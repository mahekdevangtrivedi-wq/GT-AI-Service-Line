# -*- coding: utf-8 -*-
"""GT AI Readiness & Risk Assessment - proposal deck (approach & methodology), built on the GT AI Service Lines template
(the finished deck is used as the template source: cover, agenda, dividers, thank-you).

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
dc.init_user(TEMPLATE)
add = dc.NEW.append
DCOL = [GT, CORAL, TEAL, GT2, DARK]
cur = {}
for p, v in READY7: cur.setdefault(DOMAIN_MAP[p], []).append(v)
R7 = {d: int(sum(cur[d]) / len(cur[d]) + 0.5) for d in DOMAINS if d in cur}
NQ = {d: [q for q in READINESS if q["pillar"] == d] for d in DOMAINS}
NPROF = len(PROFILE_Q); NCTX = len(CONTEXT); NREAD = len(READINESS)
SRC_NOTE = "Source: McKinsey's State of AI 2025 Survey; IDC's Worldwide AI and Generative AI Spending Guide 2025."


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


# =========================================================== 1 cover / 2 agenda (taken unchanged from the existing deck)
add(duplicate(dc.S_COVER))
add(duplicate(dc.S_AGENDA))

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
unl = [("Automate", "-27%", "operational costs", "JPM Coin turns deposits into on-chain, programmable, always-on digital money.", "ic_gears", GT),
       ("Analyze", "-31%", "downtime", "Shell scales AI-fuelled predictive maintenance to 10,000 assets.", "ic_search", GT2),
       ("Generate", "+26-60%", "developer productivity", "AI coding assistants accelerate software delivery.", "ic_bulb", TEAL),
       ("Predict", "22%", "efficiency gains", "UPS ORION optimises delivery routes with advanced analytics.", "ic_data", DARK),
       ("Assist", "<2 min", "resolution time", "Klarna's AI assistant handled two-thirds of customer-service chats in its first month (+$40M profit).", "ic_person", CORAL)]
uw = (CW - 4 * 140000) // 5; ch_ = 3750000
for i, (h, big, unit, ex, ic, col) in enumerate(unl):
    x = L + i * (uw + 140000); y = TOP - 100000
    box(s, x, y, uw, ch_, fill=WHITE, line=LAV2, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    box(s, x, y, uw, 90000, fill=col)
    icon(s, ic, x + (uw - 640000) // 2, y + 250000, 640000, fill=col)
    txt(s, x, y + 960000, uw, 400000, [[(h, True)]], size=16, color=GT, align=PP_ALIGN.CENTER)
    txt(s, x, y + 1380000, uw, 620000, [[(big, True)]], size=30 if len(big) < 7 else 24, color=col, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x, y + 2000000, uw, 300000, unit, size=11, bold=True, color=GREY, align=PP_ALIGN.CENTER)
    box(s, x + uw // 2 - 250000, y + 2380000, 500000, 18000, fill=LAV2)
    txt(s, x + 100000, y + 2480000, uw - 200000, ch_ - 2520000, ex, size=11, color=BLACK, align=PP_ALIGN.CENTER)
box(s, L, TOP + 3800000, CW, 520000, fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2, text=[[("The same capabilities create new risks ", True), ("- to data, decisions, customers and compliance - unless governance keeps pace. That is what this assessment measures.", False)]], size=12, color=WHITE)
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
txt(s, L, TOP + 4630000, CW, 250000, "Source: JPMorgan Chase public disclosures on LLM Suite adoption (2023-2025).", size=8, color=GREY)
add(s)

# =========================================================== 6 moving fast
s = content("Why Organisations Are Moving Fast", "The commercial pressure for AI adoption is becoming difficult to ignore - but adoption is moving faster than governance.")
box(s, L, TOP - 100000, 5300000, 380000, fill=GT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3, text="Six pressures driving adoption", size=12.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
pres = [("Faster decision cycles", GT), ("Productivity pressure", GT2), ("Cost optimisation", TEAL), ("Customer expectations", CORAL), ("Data volume growth", DARK), ("Competitive pressure", GT2)]
pw_ = (5300000 - 120000) // 2
for i, (p_, col) in enumerate(pres):
    x = L + (i % 2) * (pw_ + 120000); y = TOP + 360000 + (i // 2) * 560000
    box(s, x, y, pw_, 480000, fill=LAV, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
    box(s, x + 130000, y + 150000, 180000, 180000, fill=col, shape=MSO_SHAPE.OVAL)
    txt(s, x + 380000, y, pw_ - 420000, 480000, [[(p_, True)]], size=12, color=GT, anchor=MSO_ANCHOR.MIDDLE)
x2 = L + 5600000; w2 = R - x2
box(s, x2, TOP - 100000, w2, 380000, fill=CORAL, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3, text="But adoption is moving faster than governance.", size=13, bold=True, color=WHITE)
box(s, x2, TOP + 360000, w2, 880000, fill=LAV, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08, text=[[("Example: ", True), ("Morgan Stanley deployed Microsoft 365 Copilot and AI assistants to thousands of financial advisors - less time searching "
    "internal documents, more time serving clients.", False)]], size=11, align=PP_ALIGN.LEFT)
box(s, x2, TOP + 1320000, w2, 880000, fill=WHITE, line=GT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08, text=[[("\u201cBahrain's National AI Strategy promotes embedding AI in finance and operations while requiring governance to support economic "
    "diversification and regulatory compliance.\u201d", False)], [("Information & eGovernment Authority (iGA), Bahrain AI Report, 2025", True)]], size=10.5, color=GT, align=PP_ALIGN.LEFT)
box(s, L, TOP + 2450000, CW, 380000, fill=DARK, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3, text="When governance lags behind adoption", size=12.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
risks = [("ic_eye", "Shadow AI & data leakage", "Staff paste confidential or personal data into public AI tools."),
         ("ic_risk", "Unreliable or biased outputs", "Errors and bias reach customers and decisions unchecked."),
         ("ic_doccheck", "Regulatory exposure", "PDPL, sector-regulator and EU AI Act obligations missed."),
         ("ic_ai", "AI-enabled fraud", "Deepfake payment requests and AI-crafted phishing succeed.")]
rw = (CW - 3 * 130000) // 4
for i, (ic, h, d) in enumerate(risks):
    x = L + i * (rw + 130000); y = TOP + 2900000; col = [GT, CORAL, TEAL, DARK][i]
    box(s, x, y, rw, 1600000, fill=WHITE, line=LAV2, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    icon(s, ic, x + 130000, y + 150000, 520000, fill=col)
    txt(s, x + 740000, y + 100000, rw - 800000, 620000, [[(h, True)]], size=12.5, color=GT, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + 130000, y + 850000, rw - 260000, 700000, d, size=11.5)
add(s)

# =========================================================== design helpers (clean card style used by the redesigned slides)
CHECK = "\u2713  "


def card(s, x, y, w, h, head, body, accent, ic=None, size=11.5, head_size=13.5, bullet=False, spacing=4, head_h=520000):
    """White card with a coloured top strip, optional icon disc, heading and body text."""
    box(s, x, y, w, h, fill=WHITE, line=LAV2)
    box(s, x, y, w, 70000, fill=accent)
    ty = y + 160000
    if ic:
        icon(s, ic, x + 120000, ty, 440000, fill=accent); tx = x + 680000; tw = w - 760000
    else:
        tx = x + 120000; tw = w - 240000
    txt(s, tx, ty - 30000, tw, head_h - 40000, [[(head, True)]], size=head_size, color=GT, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + 120000, ty + head_h, w - 240000, h - (ty - y) - head_h - 60000, body, size=size, bullet=bullet, spacing=spacing)


def band(s, y, text, fill=DARK, h=520000, size=11.5):
    return box(s, L, y, CW, h, fill=fill, text=text if isinstance(text, list) else [[(text, False)]], size=size, color=WHITE, align=PP_ALIGN.LEFT)


# =========================================================== 7 why assess now
s = content("Why Assess AI Readiness and Risk Now", "AI is already in the organisation - in staff tools, in software and in suppliers' services. Boards and regulators now expect it to be governed.")
why = [("Regulation is arriving", "EU AI Act obligations phase in to 2026-27 (AI literacy since Feb 2025); Bahrain PDPL applies to AI processing personal data; iGA AI Policy (2025); CBB expectations for licensees; draft Bahrain AI law.", "ic_doccheck", GT),
       ("Shadow AI and embedded AI", "Staff use public GenAI tools; vendors switch on AI features in existing software. Without an inventory, risk is invisible.", "ic_eye", GT2),
       ("Agentic AI raises the stakes", "AI that takes actions through tools and APIs introduces prompt injection, excessive agency and new fraud paths (ISACA, 2026).", "ic_gears", TEAL),
       ("Privacy and data quality", "AI feeds on personal data and depends on data quality. The PDPL, ISO/IEC 27701 and data-quality standards expect minimisation, impact assessments, accuracy and traceable data.", "ic_lock", CORAL),
       ("AI-enabled threats", "Deepfake payment fraud, AI-crafted phishing and automated attacks target every organisation - whether or not it uses AI itself.", "ic_risk", DARK),
       ("Boards are asking", "Who owns AI? What is our risk appetite? Which AI do we use, with what data, and who checks it? A measured baseline focuses investment and shows progress.", "ic_people", GT)]
cw = (CW - 2 * 150000) // 3; chh = 2250000
for i, (h, d, ic, col) in enumerate(why):
    card(s, L + (i % 3) * (cw + 150000), TOP - 100000 + (i // 3) * (chh + 130000), cw, chh, h, d, col, ic=ic, size=11.5)
add(s)

# =========================================================== 8 what it answers
s = content("What the Assessment Answers", "Four answers from one questionnaire (about 30 minutes of the client's time). The free online assessment returns a short report; the detailed report is provided on request.")
ans = [("How ready are we?", "Readiness score and maturity (Partial / Informed / Repeatable / Adaptive) overall and for each of the six domains - including AI privacy and data quality.", GT, "ic_search"),
       ("How risky is our AI?", "Inherent AI risk tier (Low to Very high) from how AI is used, the decisions it influences, the data it touches and the oversight in place.", GT2, "ic_risk"),
       ("Is our readiness enough for our risk?", "Overall AI risk exposure and a recommended pathway: proceed, proceed with conditions, remediate before scaling, or pause.", TEAL, "ic_eye"),
       ("What should we do first?", "Findings per question, top-5 priorities, a phased roadmap, an AI privacy and data-quality assessment (PDPL, ISO/IEC 27701) and regulatory view - in the detailed report.", CORAL, "ic_doccheck")]
lw_ = int(CW * 0.6)
for i, (h, d, col, ic) in enumerate(ans):
    y = TOP - 100000 + i * 1130000
    box(s, L, y, 1000000, 1030000, fill=col)
    icon(s, ic, L + 220000, y + 235000, 560000)
    box(s, L + 1000000, y, 2200000, 1030000, fill=LAV, text=h, size=14, bold=True, color=GT, align=PP_ALIGN.LEFT)
    box(s, L + 3200000, y, lw_ - 3200000, 1030000, fill=WHITE, line=LAV2, text=d, size=11.5, align=PP_ALIGN.LEFT)
x2 = L + lw_ + 150000; w2 = R - x2
txt(s, x2, TOP - 100000, w2, 330000, [[("As shown in the report", True)]], size=11, color=GT, align=PP_ALIGN.CENTER)
h1 = int(w2 * 525 / 1075); h2 = int(w2 * 380 / 1075)
i1 = s.shapes.add_picture(os.path.join(IMG, "rep0_top.png"), Emu(x2), Emu(TOP + 250000), width=Emu(w2)); i1.line.color.rgb = LAV2
i2 = s.shapes.add_picture(os.path.join(IMG, "rep0_domains.png"), Emu(x2), Emu(TOP + 250000 + h1 + 100000), width=Emu(w2)); i2.line.color.rgb = LAV2
txt(s, x2, TOP + 250000 + h1 + h2 + 160000, w2, 300000, "Illustrative sample report", size=9, color=GREY, align=PP_ALIGN.CENTER)
add(s)

# =========================================================== 9 divider / 10 approach
add(duplicate(dc.S_DIV1))
s = content("Our Approach - Core Assessment Cycle", "Light-touch for the client, rigorous behind the scenes. The scoring engine stays with GT - the client receives a clear, independent report.")
w = chevrons(s, TOP - 100000, ["Day 1  Kick-off", "Day 1-2  Questionnaire", "Day 3-4  GT analysis", "Day 5  Report", "Day 5  Debrief"], h=600000, size=12.5)
det = [("ic_people", ["30-minute call with the sponsor", "Scope: whole organisation, unit or AI system", "Respondents agreed; link sent the same day"], "30 min", "Scope confirmed"),
       ("ic_laptop", ["Secure Microsoft Forms questionnaire", f"{NPROF + NCTX + NREAD} questions: profile, AI use, readiness", "Plain-language options"], "~30 min", "Completed responses"),
       ("ic_gears", ["Responses processed in GT's engine", "Risk tier, maturity, exposure, privacy & data quality, roadmap", "GT AI GRC professional review"], "None", "Quality-reviewed results"),
       ("ic_docs", ["Confidential PDF report: short summary, or the detailed report on request", "Emailed to the sponsor", "Headline results in the email"], "None", "Report"),
       ("ic_bulb", ["60-90 minute leadership debrief", "Priorities and owners agreed", "Follow-on options discussed"], "60-90 min", "90-day action plan")]
for i, (ic, items, eff, out) in enumerate(det):
    x = L + i * (w - 60000) + 30000; ww = w - 120000
    box(s, x, TOP + 600000, ww, 2050000, fill=LAV)
    icon(s, ic, x + (ww - 520000) // 2, TOP + 660000, 520000, fill=[GT, GT2, TEAL, DARK, CORAL][i])
    txt(s, x + 30000, TOP + 1240000, ww - 60000, 1400000, items, size=10.5, bullet=True, spacing=5)
    box(s, x, TOP + 2720000, ww, 520000, fill=WHITE, line=LAV2, text=[[("Client time: ", True), (eff, False)]], size=10.5, color=GT)
    box(s, x, TOP + 3290000, ww, 520000, fill=LAV2, text=[[("Output: ", True), (out, False)]], size=10.5, color=GT)
box(s, L, TOP + 3920000, CW, 450000, fill=DARK, text="This core cycle is the Essential option (5 working days). Standard (2 weeks) adds respondents, evidence review and a leadership workshop; Plus (6-8 weeks) adds AI-system deep dives.", size=11, color=WHITE)
add(s)

# =========================================================== 11 six domains
s = content("Six Assessment Domains", f"{NREAD} readiness questions across six equally weighted domains, aligned to the NIST AI RMF, ISO/IEC 42001, ISO/IEC 23894, ISO/IEC 27701 and ISO/IEC 27001.")
DIC = {"AI Governance": "ic_org", "AI Risk & Security": "ic_risk", "Training & Awareness": "ic_person", "Strategy & Value": "ic_bulb", "Data & Technology": "ic_data", "AI Privacy & Data Quality": "ic_lock"}
DCOL = [GT, GT2, TEAL, DARK, GT2, CORAL]
dw = (CW - 2 * 130000) // 3; dh = 2100000
for i, d in enumerate(DOMAINS):
    x = L + (i % 3) * (dw + 130000); y = TOP - 100000 + (i // 3) * (dh + 120000)
    nq = len(NQ[d])
    card(s, x, y, dw, dh, d, [[(f"{nq} question{'s' if nq > 1 else ''}", True)]] + [q["topic"] for q in NQ[d]], DCOL[i], ic=DIC[d], size=10.5, head_size=13.5, spacing=2, head_h=450000)
band(s, TOP + 4420000, [[("Domains are equally weighted. ", True), ("Maturity uses a four-tier scale (Partial, Informed, Repeatable, Adaptive) modelled on the NIST implementation tiers; findings convert directly into an action plan.", False)]], h=400000, size=10.5)
add(s)

# =========================================================== 12 inherent risk
s = content("Inherent AI Risk Profile", "Four questions establish how much risk the organisation's AI use carries - before considering its controls.")
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
txt(s, x2, TOP + 2450000, w2, 800000, "Overrides apply - AI that makes or materially influences decisions about individuals' rights or access to services cannot be rated Low.", size=10.5, color=GREY)
box(s, L, TOP + 3350000, CW, 650000, fill=LAV2, text=[[("Exposure = inherent risk tier x readiness maturity. ", True), ("A mature organisation with high-risk AI can be at moderate exposure; "
    "an immature organisation with the same AI is at critical exposure. The pathway follows from the exposure.", False)]], size=11.5, align=PP_ALIGN.LEFT)
txt(s, L, TOP + 4100000, CW, 400000, "Basis: EU AI Act risk categories; ISO/IEC 23894 and ISO/IEC 42005.", size=9.5, color=GREY)
add(s)

# =========================================================== 13 AI privacy & data quality
s = content("AI Privacy & Data Quality", "A dedicated domain: how AI uses personal data and how the quality of the data behind AI is controlled - against the Bahrain PDPL and ISO/IEC 27701 with ISO/IEC 27001.")
PRIV = [("Privacy governance & PDPL", "Privacy lead, records of processing, lawful basis, notices and data-subject rights, including automated decisions.", "ic_org"),
        ("Data minimisation & retention", "Only the personal data AI needs; masking and anonymisation; retention for prompts, logs and training data.", "ic_lock"),
        ("Impact assessment & cross-border data", "Privacy impact assessment before deployment; processors, vendors and AI services hosted outside Bahrain.", "ic_doccheck"),
        ("Privacy & security management", "ISO/IEC 27701 (PIMS) with ISO/IEC 27001: scope covers AI, breach response, audit.", "ic_shield"),
        ("Data quality", "Accuracy, completeness, consistency, timeliness and validity - owned, monitored and fixed at source.", "ic_check"),
        ("Provenance, representativeness & bias", "Where AI data came from, the right to use it, lineage, and checks for bias and drift.", "ic_search")]
lw_ = int(CW * 0.62); rw_ = (lw_ - 130000) // 2; rh_ = 1260000
for i, (h, d, ic) in enumerate(PRIV):
    x = L + (i % 2) * (rw_ + 130000); y = TOP - 100000 + (i // 2) * (rh_ + 90000)
    col = [CORAL, CORAL, GT, GT, TEAL, TEAL][i]
    box(s, x, y, rw_, rh_, fill=WHITE, line=LAV2); box(s, x, y, 70000, rh_, fill=col)
    icon(s, ic, x + 170000, y + 150000, 420000, fill=col)
    txt(s, x + 680000, y + 90000, rw_ - 740000, 520000, [[(h, True)]], size=12, color=GT, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, x + 170000, y + 640000, rw_ - 240000, rh_ - 660000, d, size=10.5)
x2 = L + lw_ + 150000; w2 = R - x2
box(s, x2, TOP - 100000, w2, 380000, fill=GT, text="Standards and law it is measured against", size=11.5, bold=True, color=WHITE)
std = [("Bahrain PDPL (Law No. 30 of 2018)", "lawful basis, minimisation, data quality, rights, transfers"), ("ISO/IEC 27701 (PIMS)", "privacy information management; standalone since the 2025 edition"),
       ("ISO/IEC 27001", "information security management system it is operated alongside"), ("ISO/IEC 25012 and ISO/IEC 5259", "data-quality model and data quality for analytics and ML"),
       ("NIST AI RMF", "Map and Measure functions: data provenance, quality and representativeness")]
sh_ = (rh_ * 3 + 90000 * 2 - 480000) // 5
for i, (h, d) in enumerate(std):
    y = TOP + 330000 + i * (sh_ + 0)
    box(s, x2, y, w2, sh_ - 50000, fill=LAV)
    txt(s, x2 + 60000, y + 20000, w2 - 120000, sh_ - 90000, [[(h, True)], [(d, False)]], size=10, color=GT, anchor=MSO_ANCHOR.MIDDLE, spacing=1)
band(s, TOP + 4130000, [[("In the reports: ", True), ("the short report shows the domain score and the PDPL / ISO/IEC 27701 maturity; the detailed report adds a full section with dimension scores and a requirement-by-requirement alignment.", False)]], fill=DARK, h=560000, size=11)
add(s)

# =========================================================== 14 results
s = content("How Results Are Expressed", "Clear, comparable results for every client; the detailed scoring rules remain GT's proprietary methodology.")
lw_ = int(CW * 0.6)
box(s, L, TOP - 100000, lw_, 340000, fill=GT, text="1  Every answer on a five-level scale", size=11.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
lw5 = (lw_ - 4 * 60000) // 5
for i, t in enumerate(["Not in place", "Initial / ad hoc", "Defined", "Implemented", "Optimised"]):
    box(s, L + i * (lw5 + 60000), TOP + 290000, lw5, 520000, fill=[RGBColor(0xF8, 0xCB, 0xAD), RGBColor(0xFC, 0xE4, 0xD6), RGBColor(0xFF, 0xE6, 0x99), RGBColor(0xC6, 0xE0, 0xB4), RGBColor(0x9B, 0xC2, 0xE6)][i],
        text=[[(f"Level {i}", True)], [(t, False)]], size=10)
box(s, L, TOP + 900000, lw_, 340000, fill=GT2, text="2  Domain and overall scores in four maturity tiers (modelled on the NIST implementation tiers)", size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
bw = (lw_ - 3 * 60000) // 4
for i, ((b, rng), (_, _, desc)) in enumerate(zip([("Partial", "< 50%"), ("Informed", "50 - 70%"), ("Repeatable", "70 - 90%"), ("Adaptive", "90%+")], MATURITY)):
    box(s, L + i * (bw + 60000), TOP + 1290000, bw, 1150000, fill=[RGBColor(0xF8, 0xCB, 0xAD), RGBColor(0xFF, 0xE6, 0x99), RGBColor(0xC6, 0xE0, 0xB4), RGBColor(0x9B, 0xC2, 0xE6)][i],
        text=[[(f"{b}  {rng}", True)], [(desc, False)]], size=9.5)
box(s, L, TOP + 2530000, lw_, 340000, fill=TEAL, text="3  Exposure (risk tier x maturity) sets the recommended pathway", size=11.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
for i, (exp, path, desc) in enumerate(PATHWAY):
    box(s, L + i * (bw + 60000), TOP + 2920000, bw, 900000, fill=LAV, text=[[(f"{exp} exposure", True)], [({"Low": "Proceed - self-managed", "Moderate": "Proceed with conditions", "High": "Remediate before scaling", "Critical": "Pause high-impact AI and escalate"}.get(exp, path), False)]], size=10, color=GT)
txt(s, L, TOP + 3900000, lw_, 450000, "Domains are equally weighted. Findings, risks and recommendations are generated by the GT engine and reviewed by a GT professional.", size=10, color=GREY)
x2 = L + lw_ + 150000; w2 = R - x2
im = s.shapes.add_picture(os.path.join(IMG, "rep0_top.png"), Emu(x2), Emu(TOP - 100000), width=Emu(w2)); im.line.color.rgb = LAV2
txt(s, x2, TOP - 100000 + int(w2 * 525 / 1075) + 60000, w2, 300000, "How it looks in the report (illustrative sample)", size=9, color=GREY, align=PP_ALIGN.CENTER)
add(s)

# =========================================================== 15 two reports
s = content("Two Reports: Short and Detailed", "The same questionnaire and the same scoring. The free online assessment returns the short report; the detailed report is provided on request.")
hw = (CW - 200000) // 2; ty = TOP - 100000; chh_ = 3800000
SHORT_ITEMS = ["Readiness score, maturity level and inherent risk tier", "Overall AI risk exposure and recommended pathway", "Position on the AI adoption maturity curve",
               "Readiness by domain - six domains with radar chart", "Responsible-AI principle coverage", "Score and observation for each area", "Regulatory and standards view, including PDPL and ISO/IEC 27701"]
DET_ITEMS = ["Everything in the short report", "Findings, potential risks and recommendations for every question", "AI privacy & data-quality assessment: PDPL and ISO/IEC 27701 with ISO/IEC 27001",
             "NIST AI RMF view: Govern, Map, Measure, Manage", "Top-5 priority actions and a phased roadmap", "Next steps and sign-off page"]
for k, (head, tag, col, items, img, pages) in enumerate([("SHORT REPORT", "Free online assessment", TEAL, SHORT_ITEMS, "rep_s0.png", "4 pages"), ("DETAILED REPORT", "Provided on request", GT, DET_ITEMS, "rep0.png", "8 pages")]):
    x = L + k * (hw + 200000)
    box(s, x, ty, hw, chh_, fill=WHITE, line=LAV2); box(s, x, ty, hw, 640000, fill=col)
    txt(s, x + 150000, ty, hw - 300000, 640000, [[(head, True)], [(f"{tag}  |  {pages}", False)]], size=14, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, spacing=0)
    pw_ = 1950000; ph_ = int(pw_ * 1521 / 1075)
    p_ = s.shapes.add_picture(os.path.join(IMG, img), Emu(x + 140000), Emu(ty + 790000), width=Emu(pw_)); p_.line.color.rgb = LAV2
    txt(s, x + 140000 + pw_ + 140000, ty + 760000, hw - pw_ - 420000, chh_ - 800000, [CHECK + t for t in items], size=10.5, spacing=5)
band(s, ty + chh_ + 120000, [[("Detail without extra work for the client: ", True), ("both reports come from the same responses - the detailed report adds depth, privacy and data-quality analysis, priorities and a roadmap.", False)]], h=520000, size=11.5)
add(s)

# =========================================================== 16 detailed report pages
s = content("The Detailed Report - What the Client Receives", "One report per client (8 pages) - three of its pages shown, from an illustrative sample with sample answers.")
ih = 4300000; iw_ = int(ih * 1075 / 1521)
for k, (nm, cap) in enumerate([("rep0.png", "Page 1 - Executive summary & readiness by domain"), ("rep_priv.png", "Page 5 - AI privacy & data quality"), ("rep4.png", "Page 6 - Priorities & roadmap")]):
    x = L + k * 3150000
    if nm == "rep_priv.png":
        pw2 = iw_; p_ = s.shapes.add_picture(os.path.join(IMG, nm), Emu(x), Emu(TOP - 100000), width=Emu(pw2))
    else:
        p_ = s.shapes.add_picture(os.path.join(IMG, nm), Emu(x), Emu(TOP - 100000), height=Emu(ih))
    p_.line.color.rgb = LAV2
    txt(s, x, TOP + ih - 80000, 3000000, 300000, [[(cap, True)]], size=10, color=GT)
x2 = L + 9450000
sections = ["Executive summary, pathway & AI adoption maturity curve", "Readiness by domain (radar)", "AI use & inherent risk profile", "Responsible-AI principles & NIST AI RMF functions",
            "Detailed findings & recommendations", "AI privacy & data quality (PDPL, ISO/IEC 27701)", "Top-5 priority actions", "Suggested roadmap", "Regulatory & standards considerations", "How GT can help; next steps & sign-off"]
txt(s, x2, TOP - 100000, R - x2, ih, [[("Report sections", True)]] + sections, size=10, spacing=5, color=BLACK)
add(s)

# =========================================================== 17 GT's own position (AI Ready7)
s = content("Our Own Starting Point - GT Bahrain AI Ready7 (Sept 2026)", "We assessed ourselves first. GT Bahrain's AI Ready7 results by pillar, and consolidated into the assessment domains.")
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
rows = [["Assessment domain", "Compared Ready7 pillar(s)", "Score"]] + \
       [[d, (" + ".join(p for p, _ in READY7 if DOMAIN_MAP[p] == d) if d in R7 else "New domain - not covered by Ready7"), (f"{R7[d]}%" if d in R7 else "-")] for d in DOMAINS] + \
       [["Overall", "Readiness status: Informed", f"{READY7_OVERALL:.1f}%"]]
fills = {(i, 2): RGBColor(0xF8, 0xCB, 0xAD) if R7[d] < 60 else RGBColor(0xFF, 0xE6, 0x99) for i, d in enumerate(DOMAINS, 1) if d in R7}
table(s, x2, TOP - 50000, w2, rows, [2.4, 3.4, 0.8], size=9.5, rowh=400000, fills=fills)
txt(s, x2, TOP + 3250000, w2, 1000000, [[("Our priorities: ", True), ("Strategy & Value (50%) and AI Governance (60%) - including lifecycle AI risk assessment (Ready7 2.4) and data classification for AI (5.1), "
     "both 'Not Met'. The AI Readiness & Risk Assessment adds deeper governance, risk-management, privacy and data-quality questions to measure exactly these areas.", False)]], size=10.5)
txt(s, L, TOP + 4250000, CW, 300000, "Source: GT Bahrain AI Ready7 Result Report (Sept 2026). Domain = average of its Ready7 pillars. Coral = below 60%.", size=8.5, color=GREY)
add(s)

# =========================================================== 18 divider / 19 options
add(duplicate(dc.S_DIV2))
s = content("Engagement Options", "The free online assessment returns the short report. The detailed report is delivered in every option below, with different depth of review.")
opts = [("ESSENTIAL", "5 working days", TEAL, "ic_search", "Best for: a fast, independent baseline", ["Kick-off call", "Online questionnaire", "GT analysis and quality review", "Detailed PDF report", "30-minute results call"]),
        ("STANDARD", "2 weeks", GT, "ic_people", "Best for: leadership alignment and an agreed plan", ["Everything in Essential", "Multi-respondent input (risk, IT, business)", "Evidence review and validation interviews", "Leadership debrief workshop (90 min)", "Agreed 90-day action plan"]),
        ("PLUS", "6 weeks (extendable to 8)", DARK, "ic_doccheck", "Best for: regulated organisations and boards", ["Everything in Standard", "AI inventory and risk-tiering of AI systems", "Risk, impact and privacy assessments of high-risk AI systems",
                                                                               "Control review against the GT AI Risk Assessment Framework", "ISO/IEC 42001 gap assessment", "AI security & agent review (selected systems)", "Remediation roadmap and board presentation"])]
ow = (CW - 2 * 150000) // 3
for i, (h, dur, col, ic, best, items) in enumerate(opts):
    x = L + i * (ow + 150000)
    box(s, x, TOP - 150000, ow, 640000, fill=col)
    icon(s, ic, x + 90000, TOP - 90000, 520000)
    txt(s, x + 700000, TOP - 150000, ow - 750000, 640000, [[(h, True)], [(dur, False)]], size=14, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, spacing=0)
    box(s, x, TOP + 490000, ow, 400000, fill=LAV2, text=best, size=10.5, bold=True, color=GT)
    box(s, x, TOP + 890000, ow, 2150000, fill=LAV)
    txt(s, x + 70000, TOP + 960000, ow - 140000, 2050000, items, size=10.5 if i == 2 else 11.5, bullet=True, spacing=5 if i == 2 else 7)
box(s, L, TOP + 3150000, CW, 380000, fill=GT2, text="Duration by option - detailed timelines on the following slides", size=12, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
w3 = (CW - 2 * 80000) // 3
for i, (h, d, col) in enumerate([("ESSENTIAL  |  5 working days", "Report on Day 5; results call", TEAL), ("STANDARD  |  2 weeks", "Report end of week 2; leadership debrief & 90-day plan", GT),
                                 ("PLUS  |  6 weeks, extendable to 8", "Readiness report week 2; full report & board presentation week 6", DARK)]):
    box(s, L + i * (w3 + 80000), TOP + 3600000, w3, 700000, fill=WHITE, line=col, text=[[(h, True)], [(d, False)]], size=11, color=GT)
txt(s, L, TOP + 4380000, CW, 300000, "Fees are proposed per option once scope (entities, AI systems, respondents) is confirmed. Re-assessment after 6-12 months shows progress on the same scale.", size=10, color=GREY)
add(s)

# =========================================================== 20-22 timelines
LIGHT = RGBColor(0xF2, 0xF2, 0xF2)


def gantt(title, subtitle, cols, rows, ext_from=None, notes=None):
    """rows: (label, colour, start, duration, kind) - kind 'bar' | 'ms' (milestone at start) | 'hdr' (group header)."""
    s = content(title, subtitle)
    lab_w = 3300000; gx = L + lab_w; gw = CW - lab_w; n = len(cols); cw = gw // n
    y0 = TOP - 120000
    for k, c in enumerate(cols):
        ext = ext_from is not None and k >= ext_from
        box(s, gx + k * cw, y0, cw - 15000, 330000, fill=LAV2 if ext else GT, text=c, size=10 if n <= 8 else 9, bold=True, color=GT if ext else WHITE)
    rh = min(330000, (3950000 - 330000) // max(1, len(rows)))
    y = y0 + 380000
    for label, col, st, du, kind in rows:
        if kind == "hdr":
            box(s, L, y, CW, rh - 40000, fill=LAV, text=label, size=9.5, bold=True, color=GT, align=PP_ALIGN.LEFT)
            y += rh; continue
        txt(s, L, y, lab_w - 50000, rh, label, size=9.5, color=BLACK, anchor=MSO_ANCHOR.MIDDLE)
        for k in range(n):
            if ext_from is not None and k >= ext_from:
                box(s, gx + k * cw, y + 20000, cw - 15000, rh - 40000, fill=RGBColor(0xF7, 0xF4, 0xFB))
        if kind == "bar":
            box(s, gx + int(st * cw) + 15000, y + 60000, max(int(du * cw) - 45000, 60000), rh - 120000, fill=col)
        elif kind == "ms":
            d = rh - 90000
            for st_ in (st if isinstance(st, tuple) else (st,)):
                box(s, gx + int(st_ * cw) - d // 2, y + 45000, d, d, fill=col, shape=MSO_SHAPE.DIAMOND)
        y += rh
    if notes:
        nw = (CW - (len(notes) - 1) * 120000) // len(notes); ny = TOP + 4000000
        for i, (h, d) in enumerate(notes):
            box(s, L + i * (nw + 120000), ny, nw, 640000, fill=LAV, text=[[(h + ": ", True), (d, False)]], size=9.5, align=PP_ALIGN.LEFT)
    txt(s, R - 3800000, TOP + 4680000, 3800000, 250000, [[("◆ ", True), ("milestone / steering checkpoint", False)]], size=8.5, color=GREY, align=PP_ALIGN.RIGHT)
    return s


P1, P2, P3, P4 = TEAL, GT, GT2, CORAL
add(gantt("Essential - Timeline (5 Working Days)", "A fast, independent baseline: one respondent, GT-scored report on Day 5.",
          ["Day 1", "Day 2", "Day 3", "Day 4", "Day 5"],
          [("Kick-off call & scope confirmation", P1, 0, 0.5, "bar"), ("Questionnaire link sent", P1, 0.5, 0, "ms"),
           ("Client completes questionnaire (~30 min)", P2, 0.3, 1.7, "bar"), ("Responses received", P2, 2, 0, "ms"),
           ("GT analysis, scoring & findings", P3, 2, 1.6, "bar"), ("GT quality review", P3, 3.4, 0.6, "bar"),
           ("Report issued (PDF by email)", P4, 4.3, 0, "ms"), ("Results call (30 min)", P4, 4.4, 0.6, "bar")],
          notes=[("Client time", "about 2 hours in total"), ("Deliverables", "confidential PDF report, results call"), ("Assumption", "questionnaire completed by Day 2")]))
add(gantt("Standard - Timeline (2 Weeks)", "A validated result and an agreed plan: several respondents, evidence review and a leadership workshop.",
          ["D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10"],
          [("Week 1 - Collect", None, 0, 0, "hdr"),
           ("Kick-off & respondent briefing", P1, 0, 1, "bar"), ("Multi-respondent questionnaire", P2, 0.5, 3.5, "bar"),
           ("Evidence request & review (policies, inventory, risk register, privacy records)", P2, 2, 4, "bar"),
           ("Week 2 - Validate & report", None, 0, 0, "hdr"),
           ("Validation interviews (risk, IT, privacy, business)", P3, 5, 2, "bar"), ("GT analysis, scoring & quality review", P3, 5.5, 2.5, "bar"),
           ("Report issued", P4, 8, 0, "ms"), ("Leadership debrief workshop (90 min)", P4, 8.2, 1, "bar"), ("90-day action plan agreed", P4, 9.6, 0, "ms")],
          notes=[("Client time", "sponsor, 3-5 respondents, 2-3 interviews, 90-min workshop"), ("Deliverables", "PDF report, debrief deck, 90-day action plan"),
                 ("Checkpoints", "end of week 1 (evidence complete), Day 10 (plan agreed)")]))
add(gantt("Plus - Timeline (6 Weeks, Extendable to 8)", "Organisation-wide readiness plus a deep dive into the AI systems themselves - for regulated organisations and boards.",
          ["W1", "W2", "W3", "W4", "W5", "W6", "W7 (ext.)", "W8 (ext.)"],
          [("Mobilise & kick-off; AI system list requested", P1, 0, 0.6, "bar"), ("Readiness questionnaire, evidence & interviews", P2, 0.3, 1.7, "bar"),
           ("Readiness report & debrief", P2, 2, 0, "ms"),
           ("AI inventory & risk-tiering of AI systems", P3, 1, 2, "bar"), ("Risk, impact & privacy assessments - high-risk AI systems", P3, 2.5, 2.5, "bar"),
           ("Control review vs GT AI Risk Assessment Framework", P3, 2.5, 2.5, "bar"), ("ISO/IEC 42001 gap assessment", GT2, 3, 2, "bar"),
           ("AI security & agent review (selected systems)", GT2, 3.5, 2, "bar"), ("Remediation roadmap & board pack", P4, 4.8, 1.2, "bar"),
           ("Steering checkpoints", AMBER, (2, 4, 6), 0, "ms"), ("Final report & board presentation", P4, 6, 0, "ms"),
           ("Extension: additional entities / AI systems", LAV2, 6, 2, "bar")],
          ext_from=6,
          notes=[("Client time", "sponsor weekly 30 min; system owners 1-2 h each; board session"),
                 ("Deliverables", "readiness report, AI inventory & risk tiers, assessments, 42001 gap, security findings, roadmap, board pack"),
                 ("Extension", "weeks 7-8 when more entities or AI systems are in scope")]))

# =========================================================== 23 deliverables
s = content("Deliverables, Client Inputs and Standards")
hw = (CW - 2 * 150000) // 3
cols = [("DELIVERABLES", GT, "ic_docs", ["Detailed AI Readiness & Risk Assessment report (PDF): Day 5 / week 2 / week 6 by option", "Readiness by domain (six domains) and overall maturity", "Inherent AI risk tier, exposure and pathway",
                                          "AI privacy & data-quality assessment (PDPL, ISO/IEC 27701)", "Findings, top-5 priorities and roadmap", "Debrief presentation (Standard / Plus)"]),
        ("FROM THE CLIENT", GT2, "ic_people", ["An executive sponsor and a coordinator", "Completed questionnaire (about 30 minutes) and, for Standard / Plus, evidence and interviews",
                                                "Optional evidence: AI policy, AI inventory, risk register, vendor contracts, privacy records", "Availability for the debrief"]),
        ("STANDARDS & FRAMEWORKS", TEAL, "ic_doccheck", ["NIST AI RMF 1.0", "ISO/IEC 42001, ISO/IEC 23894, ISO/IEC 42005", "ISO/IEC 27701 (PIMS) with ISO/IEC 27001", "ISO/IEC 25012 and ISO/IEC 5259 (data quality)",
                                                         "Bahrain PDPL; EU AI Act; CBB; iGA AI Policy", "ISACA Securing AI Agents (2026); OWASP LLM Top 10"])]
for i, (h, col, ic, items) in enumerate(cols):
    x = L + i * (hw + 150000)
    box(s, x, TOP - 250000, hw, 620000, fill=col)
    icon(s, ic, x + 90000, TOP - 200000, 520000)
    txt(s, x + 700000, TOP - 250000, hw - 750000, 620000, [[(h, True)]], size=13, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    box(s, x, TOP + 370000, hw, 2300000, fill=LAV)
    txt(s, x + 70000, TOP + 440000, hw - 140000, 2200000, items, size=11.5, bullet=True, spacing=6)
box(s, L, TOP + 2760000, CW, 330000, fill=GT2, text="What the detailed report contains", size=11.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
secs = ["Executive summary, pathway & maturity curve", "Readiness by domain (radar)", "AI use & inherent risk", "Principles & NIST AI RMF functions", "Findings & recommendations",
        "AI privacy & data quality", "Top-5 actions & roadmap", "Regulatory & standards", "How GT can help & sign-off"]
cw9 = (CW - 8 * 50000) // 9
for i, t in enumerate(secs):
    box(s, L + i * (cw9 + 50000), TOP + 3130000, cw9, 760000, fill=WHITE, line=GT2, text=[[(str(i + 1), True)], [(t, False)]], size=9, color=GT)
box(s, L, TOP + 4000000, CW, 600000, fill=DARK, text=[[("Natural next step: ", True), ("the GT AI Strategy & Roadmap service uses the Opportunity Discovery Toolkit to show WHERE AI fits - "
    "the readiness assessment shows HOW READY the organisation is to adopt it safely. Both can run together.", False)]], size=11.5, color=WHITE)
add(s)

add(duplicate(dc.S_THANKS))
dc.finish(OUT)
