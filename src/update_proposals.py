# -*- coding: utf-8 -*-
"""In-place updates to the (manually edited) proposal decks - keeps the user's edits and adds / changes specific slides.

Usage: update_proposals.py READINESS.pptx STRATEGY.pptx
  Readiness: per-option durations (Essential 5 working days, Standard 2 weeks, Plus 6 weeks extendable to 8) and a
             Gantt timeline slide for each option.
  Strategy : 'How we formulate your AI strategy' slide after Phase 4; client self-assessment portal removed from the tools list.
"""
import sys, os, copy
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
sys.path.insert(0, os.path.dirname(__file__))
import deck_common as dc
from deck_common import txt, box, content, GT, GT2, LAV, LAV2, TEAL, CORAL, DARK, GREY, WHITE, BLACK, AMBER, GREEN, L, R, CW, TOP

READY, STRAT = sys.argv[1], sys.argv[2]
LIGHT = RGBColor(0xF2, 0xF2, 0xF2)


# ------------------------------------------------------------------ helpers
def use(prs, content_slide):
    dc.prs = prs; dc.S_CONTENT = content_slide


def move_last_to(prs, index):
    lst = prs.slides._sldIdLst; el = lst[-1]; lst.remove(el); lst.insert(index, el)


def slide_by_title(prs, title):
    for i, s in enumerate(prs.slides):
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip().startswith(title): return i, s
    raise KeyError(title)


def shape(s, name):
    for sh in s.shapes:
        if sh.name == name: return sh
    raise KeyError(name)


def set_para(sh, idx, text, size=None):
    pa = sh.text_frame.paragraphs[idx]
    pa.runs[0].text = text
    for r in pa.runs[1:]: r.text = ""
    if size:
        for r in pa.runs: r.font.size = Pt(size)


def set_paras(sh, texts, size=None):
    """Replace bullet paragraphs, reusing the first paragraph's formatting."""
    tf = sh.text_frame; first = tf.paragraphs[0]._p
    for pa in list(tf.paragraphs)[1:]: pa._p.getparent().remove(pa._p)
    for _ in texts[1:]: tf._txBody.append(copy.deepcopy(first))
    for i, t in enumerate(texts): set_para(sh, i, t, size)


def replace_text(s, old, new):
    for sh in s.shapes:
        if sh.has_text_frame:
            for i, pa in enumerate(sh.text_frame.paragraphs):
                if old in pa.text and pa.runs:
                    set_para(sh, i, pa.text.replace(old, new))


def gantt(title, subtitle, cols, rows, ext_from=None, notes=None, unit_label=""):
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
        for k in range(n):   # light grid
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


# ================================================================== READINESS PROPOSAL
prs = Presentation(READY)
oi, opt = slide_by_title(prs, "Engagement Options")
use(prs, opt)
# --- options slide
for sh in opt.shapes:
    if sh.has_text_frame and sh.text_frame.text.startswith("Engagement Options"):
        set_para(sh, 0, "Engagement Options")
for name, dur in [("TextBox 10", "5 working days"), ("TextBox 16", "2 weeks"), ("TextBox 22", "6 weeks (extendable to 8)")]:
    set_para(shape(opt, name), 1, dur)
B_ = "\u2022  "
set_paras(shape(opt, "TextBox 19"), [B_ + t for t in ["Everything in Essential", "Multi-respondent input (risk, IT, business)", "Evidence review and validation interviews",
                                     "Leadership debrief workshop (90 min)", "Agreed 90-day action plan"]])
set_paras(shape(opt, "TextBox 25"), [B_ + t for t in ["Everything in Standard", "AI inventory and risk-tiering of AI systems", "Risk & impact assessments of high-risk AI systems",
                                     "Control review against the GT AI Risk Assessment Framework", "ISO/IEC 42001 gap assessment", "AI security & agent review (selected systems)",
                                     "Remediation roadmap and board presentation"]], size=10.5)
for nm in ["Rectangle 27", "Rectangle 28", "Rectangle 29", "Rectangle 30", "Rectangle 31"]:
    el = shape(opt, nm)._element; el.getparent().remove(el)
set_para(shape(opt, "Rectangle 26"), 0, "Duration by option - detailed timelines on the following slides")
y = TOP + 3450000; w3 = (CW - 2 * 80000) // 3
for i, (h, d, col) in enumerate([("ESSENTIAL  |  5 working days", "Report on Day 5; results call", TEAL),
                                 ("STANDARD  |  2 weeks", "Report end of week 2; leadership debrief & 90-day plan", GT),
                                 ("PLUS  |  6 weeks, extendable to 8", "Readiness report week 2; full report & board presentation week 6", DARK)]):
    box(opt, L + i * (w3 + 80000), y, w3, 700000, fill=LAV, line=col, text=[[(h, True)], [(d, False)]], size=11, color=GT)
# --- Gantt slides (inserted after the options slide)
P1, P2, P3, P4 = TEAL, GT, GT2, CORAL
g1 = gantt("Essential - Timeline (5 Working Days)", "A fast, independent baseline: one respondent, GT-scored report on Day 5.",
           ["Day 1", "Day 2", "Day 3", "Day 4", "Day 5"],
           [("Kick-off call & scope confirmation", P1, 0, 0.5, "bar"), ("Questionnaire link sent", P1, 0.5, 0, "ms"),
            ("Client completes questionnaire (~45 min)", P2, 0.3, 1.7, "bar"), ("Responses received", P2, 2, 0, "ms"),
            ("GT analysis, scoring & findings", P3, 2, 1.6, "bar"), ("GT quality review", P3, 3.4, 0.6, "bar"),
            ("Report issued (PDF by email)", P4, 4.3, 0, "ms"), ("Results call (30 min)", P4, 4.4, 0.6, "bar")],
           notes=[("Client time", "about 2 hours in total"), ("Deliverables", "confidential PDF report, results call"), ("Assumption", "questionnaire completed by Day 2")])
move_last_to(prs, oi + 1)
g2 = gantt("Standard - Timeline (2 Weeks)", "A validated result and an agreed plan: several respondents, evidence review and a leadership workshop.",
           ["D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10"],
           [("Week 1 - Collect", None, 0, 0, "hdr"),
            ("Kick-off & respondent briefing", P1, 0, 1, "bar"), ("Multi-respondent questionnaire", P2, 0.5, 3.5, "bar"),
            ("Evidence request & review (policies, inventory, risk register)", P2, 2, 4, "bar"),
            ("Week 2 - Validate & report", None, 0, 0, "hdr"),
            ("Validation interviews (risk, IT, business)", P3, 5, 2, "bar"), ("GT analysis, scoring & quality review", P3, 5.5, 2.5, "bar"),
            ("Report issued", P4, 8, 0, "ms"), ("Leadership debrief workshop (90 min)", P4, 8.2, 1, "bar"), ("90-day action plan agreed", P4, 9.6, 0, "ms")],
           notes=[("Client time", "sponsor, 3-5 respondents, 2-3 interviews, 90-min workshop"), ("Deliverables", "PDF report, debrief deck, 90-day action plan"),
                  ("Checkpoints", "end of week 1 (evidence complete), Day 10 (plan agreed)")])
move_last_to(prs, oi + 2)
g3 = gantt("Plus - Timeline (6 Weeks, Extendable to 8)", "Organisation-wide readiness plus a deep dive into the AI systems themselves - for regulated organisations and boards.",
           ["W1", "W2", "W3", "W4", "W5", "W6", "W7 (ext.)", "W8 (ext.)"],
           [("Mobilise & kick-off; AI system list requested", P1, 0, 0.6, "bar"), ("Readiness questionnaire, evidence & interviews", P2, 0.3, 1.7, "bar"),
            ("Readiness report & debrief", P2, 2, 0, "ms"),
            ("AI inventory & risk-tiering of AI systems", P3, 1, 2, "bar"), ("Risk & impact assessments - high-risk AI systems", P3, 2.5, 2.5, "bar"),
            ("Control review vs GT AI Risk Assessment Framework", P3, 2.5, 2.5, "bar"), ("ISO/IEC 42001 gap assessment", GT2, 3, 2, "bar"),
            ("AI security & agent review (selected systems)", GT2, 3.5, 2, "bar"), ("Remediation roadmap & board pack", P4, 4.8, 1.2, "bar"),
            ("Steering checkpoints", AMBER, (2, 4, 6), 0, "ms"), ("Final report & board presentation", P4, 6, 0, "ms"),
            ("Extension: additional entities / AI systems", LAV2, 6, 2, "bar")],
           ext_from=6,
           notes=[("Client time", "sponsor weekly 30 min; system owners 1-2 h each; board session"),
                  ("Deliverables", "readiness report, AI inventory & risk tiers, assessments, 42001 gap, security findings, roadmap, board pack"),
                  ("Extension", "weeks 7-8 when more entities or AI systems are in scope")])
move_last_to(prs, oi + 3)
# --- other readiness slides that referred to a single 5-day timeline
for s in prs.slides:
    replace_text(s, "about 45 minutes of the client's time, delivered within 5 working days.",
                 "about 45 minutes of the client's time. Report in 5 working days (Essential), 2 weeks (Standard) or 6-8 weeks (Plus).")
    replace_text(s, "Our Approach - Report Within 5 Working Days", "Our Approach - Core Assessment Cycle")
    replace_text(s, "Timeline assumes the questionnaire is completed by Day 2; otherwise the report is issued within 3 working days of receiving complete responses.",
                 "This core cycle is the Essential option (5 working days). Standard (2 weeks) adds respondents, evidence review and a leadership workshop; Plus (6-8 weeks) adds AI-system deep dives.")
    replace_text(s, "Confidential AI Readiness & Risk Assessment report (PDF) by Day 5", "Confidential AI Readiness & Risk Assessment report (PDF): Day 5 / week 2 / week 6 by option")
    replace_text(s, "Completed questionnaire by Day 2 (about 45 minutes)", "Completed questionnaire (about 45 minutes) and, for Standard / Plus, evidence and interviews")
    replace_text(s, " Both can run together within the same 5 working days.", " Both can run together.")
prs.save(READY)
print("readiness updated:", len(prs.slides), "slides")

# ================================================================== STRATEGY PROPOSAL
prs = Presentation(STRAT)
# --- remove client self-assessment portal from the tools table
for s in prs.slides:
    for sh in s.shapes:
        if sh.has_table:
            tbl = sh.table
            for r in list(tbl.rows):
                if any("self-assessment portal" in c.text for c in r.cells):
                    h = r.height; r._tr.getparent().remove(r._tr); sh.height = sh.height - h
# --- 'How we formulate your AI strategy' after Phase 4
pi, p4 = slide_by_title(prs, "Phase 4: AI Strategy Formulation")
use(prs, p4)
s = content("How We Formulate Your AI Strategy", "Six steps on Days 3-5 turn the evidence (situational analysis, AI position and opportunity map) into a strategy leadership can approve.")
steps = [("1", "Synthesise the evidence", "Combine the situational analysis, AI position archetype and opportunity map into 4-6 strategic themes and options.", "Sponsor review call", "Themes & options"),
         ("2", "Set the ambition", "Leadership chooses the ambition level that fits the archetype - automate first, adopt & configure, or build & scale - and where AI will not be used.", "2-hour leadership workshop (Day 4)", "AI vision & ambition"),
         ("3", "Define objectives & KPIs", "3-5 measurable objectives with baselines and targets (hours released, cycle time, quality, customer, risk) from the library KPIs.", "Finance / business owners confirm baselines", "Objectives & KPI set"),
         ("4", "Shape the portfolio", "Group ranked opportunities into an automation programme and an AI programme; decide build / buy / partner; size effort and benefits.", "Function heads validate priorities", "Prioritised portfolio & business case"),
         ("5", "Set guardrails & operating model", "Risk appetite, approval by risk tier (GT AI Risk Assessment Framework), sponsor, AI owner / CoE, intake and stage-gates.", "Risk, IT and compliance input", "Risk appetite & operating model"),
         ("6", "Validate & sign off", "Test against capacity, data foundations and readiness gaps; align with budget; present on Day 5 for approval.", "Leadership presentation (Day 5)", "Strategy on a page & roadmap")]
sw = (CW - 5 * 80000) // 6
for i, (n, h, d, cin, out) in enumerate(steps):
    x = L + i * (sw + 80000); col = [GT, GT2, TEAL, DARK, CORAL, GREEN][i]
    box(s, x, TOP - 120000, sw, 560000, fill=col, text=[[(n + "  ", True), (h, True)]], size=11, color=WHITE)
    box(s, x, TOP + 440000, sw, 1650000, fill=LAV, text=d, size=9.5, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
    box(s, x, TOP + 2120000, sw, 520000, fill=WHITE, line=LAV2, text=[[("Client: ", True), (cin, False)]], size=8.5, color=GT, align=PP_ALIGN.LEFT)
    box(s, x, TOP + 2670000, sw, 520000, fill=LAV2, text=[[("Output: ", True), (out, False)]], size=9, color=GT, align=PP_ALIGN.LEFT)
box(s, L, TOP + 3330000, CW, 330000, fill=DARK, text="THE RESULT: YOUR AI STRATEGY ON A PAGE", size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
page = [("Vision & ambition", "where AI will and will not be used"), ("3-5 objectives & KPIs", "measurable targets and owners"), ("Priority themes", "automation programme + AI programme"),
        ("Enablers", "data, platforms, people & change"), ("Guardrails", "risk appetite, approval by tier, operating model")]
pw = (CW - 4 * 80000) // 5
for i, (h, d) in enumerate(page):
    box(s, L + i * (pw + 80000), TOP + 3720000, pw, 760000, fill=WHITE, line=GT, text=[[(h, True)], [(d, False)]], size=10, color=GT)
move_last_to(prs, pi + 1)
prs.save(STRAT)
print("strategy updated:", len(prs.slides), "slides")
