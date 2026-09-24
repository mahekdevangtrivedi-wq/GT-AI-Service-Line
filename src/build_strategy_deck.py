# -*- coding: utf-8 -*-
"""GT Bahrain AI Strategy & Roadmap deck, built on the GT AI Service Lines template."""
import sys, os, copy
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData, BubbleChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION
from pptx.oxml.ns import qn
sys.path.insert(0, os.path.dirname(__file__))
from strategy_data import *

TEMPLATE, OUT = sys.argv[1], sys.argv[2]
prs = Presentation(TEMPLATE)
orig = list(prs.slides)
S_COVER, S_AGENDA, S_DIVIDER, S_CONTENT, S_THANKS = orig[0], orig[1], orig[5], orig[8], orig[18]

GT = RGBColor(0x4F, 0x2D, 0x7F); GT2 = RGBColor(0x7B, 0x5B, 0xA6); LAV = RGBColor(0xED, 0xE7, 0xF6); LAV2 = RGBColor(0xD9, 0xCC, 0xEB)
TEAL = RGBColor(0x00, 0xA7, 0xB5); CORAL = RGBColor(0xE8, 0x70, 0x5F); DARK = RGBColor(0x2E, 0x1A, 0x47); GREY = RGBColor(0x59, 0x59, 0x59)
WHITE = RGBColor(255, 255, 255); BLACK = RGBColor(0x22, 0x22, 0x22); AMBER = RGBColor(0xF2, 0xB1, 0x3C); GREEN = RGBColor(0x5B, 0xA8, 0x5B)
FONT = "Arial"


# ------------------------------------------------------------------ slide duplication helpers
def clone_shapes(src, dst, names=None):
    rid_map = {}
    for rId, rel in src.part.rels.items():
        if rel.reltype.endswith("/slideLayout") or rel.reltype.endswith("/notesSlide"):
            continue
        rid_map[rId] = dst.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref) if rel.is_external else dst.part.rels.get_or_add(rel.reltype, rel.target_part)
    for el in src.shapes._spTree.iterchildren():
        tag = el.tag.split("}")[1]
        if tag in ("nvGrpSpPr", "grpSpPr", "extLst"):
            continue
        if names is not None:
            nm = el.find(".//" + qn("p:cNvPr"))
            if nm is None or nm.get("name") not in names:
                continue
        new = copy.deepcopy(el)
        for node in new.iter():
            for attr in (qn("r:embed"), qn("r:link"), qn("r:id")):
                v = node.get(attr)
                if v in rid_map: node.set(attr, rid_map[v])
        dst.shapes._spTree.append(new)
    bg = src._element.find(qn("p:cSld")).find(qn("p:bg"))
    if bg is not None and names is None:
        dst._element.find(qn("p:cSld")).insert(0, copy.deepcopy(bg))


def duplicate(src):
    s = prs.slides.add_slide(src.slide_layout)
    for shp in list(s.shapes): shp._element.getparent().remove(shp._element)
    clone_shapes(src, s)
    return s


def set_text(shape, text, size=None, bold=None, color=None):
    tf = shape.text_frame
    runs = [r for p in tf.paragraphs for r in p.runs]
    if not runs: return
    f0 = runs[0].font
    for p in list(tf.paragraphs)[1:]: p._p.getparent().remove(p._p)
    p = tf.paragraphs[0]
    for r in list(p.runs)[1:]: r._r.getparent().remove(r._r)
    lines = text.split("\n")
    p.runs[0].text = lines[0]
    for ln in lines[1:]:
        np_ = copy.deepcopy(p._p); tf._txBody.append(np_)
        tf.paragraphs[-1].runs[0].text = ln
    for p in tf.paragraphs:
        for r in p.runs:
            if size: r.font.size = Pt(size)
            if bold is not None: r.font.bold = bold
            if color: r.font.color.rgb = color


def find(slide, contains):
    for sh in slide.shapes:
        if sh.has_text_frame and contains in sh.text_frame.text: return sh
        if sh.shape_type == 6:
            for g in sh.shapes:
                if g.has_text_frame and contains in g.text_frame.text: return g
    return None


# ------------------------------------------------------------------ drawing helpers
def txt(slide, x, y, w, h, paras, size=11, color=BLACK, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, bullet=False, spacing=2):
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(45720); tf.margin_top = tf.margin_bottom = Emu(22860)
    if isinstance(paras, str): paras = [paras]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align; para.space_after = Pt(spacing)
        segs = p if isinstance(p, list) else [(p, bold)]
        for seg, b in segs:
            r = para.add_run(); r.text = (("•  " if bullet and seg is segs[0][0] else "") + seg)
            r.font.name = FONT; r.font.size = Pt(size); r.font.bold = b; r.font.color.rgb = color
    return tb


def box(slide, x, y, w, h, fill=LAV, line=None, text=None, size=11, color=BLACK, bold=False, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=MSO_SHAPE.RECTANGLE, radius=None):
    s = slide.shapes.add_shape(shape, Emu(x), Emu(y), Emu(w), Emu(h))
    if fill is None: s.fill.background()
    else:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None: s.line.fill.background()
    else:
        s.line.color.rgb = line; s.line.width = Pt(1)
    s.shadow.inherit = False
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE: s.adjustments[0] = radius
    if text is not None:
        tf = s.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = Emu(64008); tf.margin_top = tf.margin_bottom = Emu(36576)
        paras = text if isinstance(text, list) else [text]
        for i, p in enumerate(paras):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = align
            segs = p if isinstance(p, list) else [(p, bold)]
            for seg, b in segs:
                r = para.add_run(); r.text = seg; r.font.name = FONT; r.font.size = Pt(size); r.font.bold = b; r.font.color.rgb = color
    return s


def table(slide, x, y, w, rows, colw, size=9, header_fill=GT, rowh=None, first_bold=True, fills=None):
    nr, nc = len(rows), len(rows[0])
    shp = slide.shapes.add_table(nr, nc, Emu(x), Emu(y), Emu(w), Emu(rowh * nr if rowh else 300000 * nr))
    tbl = shp.table
    tot = sum(colw)
    for i, cw in enumerate(colw): tbl.columns[i].width = Emu(int(w * cw / tot))
    for ri, row in enumerate(rows):
        if rowh: tbl.rows[ri].height = Emu(rowh)
        for ci, val in enumerate(row):
            c = tbl.cell(ri, ci); c.text = ""
            c.margin_left = c.margin_right = Emu(54864); c.margin_top = c.margin_bottom = Emu(27432)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]; r = p.add_run(); r.text = str(val)
            r.font.name = FONT; r.font.size = Pt(size)
            if ri == 0:
                r.font.bold = True; r.font.color.rgb = WHITE; c.fill.solid(); c.fill.fore_color.rgb = header_fill
            else:
                r.font.color.rgb = BLACK; r.font.bold = first_bold and ci == 0
                c.fill.solid(); c.fill.fore_color.rgb = LAV if ri % 2 == 0 else WHITE
                if fills and (ri, ci) in fills:
                    c.fill.fore_color.rgb = fills[(ri, ci)]
    return tbl


SW, SH = prs.slide_width, prs.slide_height
L = 524396; R = 11787482; CW = R - L; TOP = 1480000


def content(title, subtitle=None):
    s = prs.slides.add_slide(S_CONTENT.slide_layout)
    for shp in list(s.shapes): shp._element.getparent().remove(shp._element)
    clone_shapes(S_CONTENT, s, names={"Copyright", "Straight Connector 4", "Slide Number Placeholder 1", "Picture 6"})
    cp = find(s, "Grant Thornton Bahrain")
    if cp: set_text(cp, "© 2026 Grant Thornton Bahrain. All rights reserved")
    txt(s, L, 400000, CW, 560000, [[(title, True)]], size=26, color=GT)
    if subtitle:
        txt(s, L - 20000, 930000, CW, 480000, subtitle, size=11.5, color=GREY)
    return s


def divider(title, number):
    s = duplicate(S_DIVIDER)
    set_text(find(s, "AI Service Lines"), title)
    t = find(s, "Grant Thornton Bahrain")
    if t: set_text(t, "© 2026 Grant Thornton Bahrain. All rights reserved")
    return s


NEW = []
# =========================================================== 1 cover
cov = duplicate(S_COVER)
set_text(find(cov, "Artificial Intelligence"), "AI Strategy &\nRoadmap 2026-2028")
set_text(find(cov, "We go beyond"), "Grant Thornton Bahrain  |  September 2026")
NEW.append(cov)

# =========================================================== 2 agenda
ag = duplicate(S_AGENDA)
set_text(find(ag, "Market Opportunities"), "Situational & SWOT Analysis")
set_text(find(ag, "AI Service Lines for"), "AI Strategy Formulation")
set_text(find(ag, "System Setup"), "Implementation Roadmap")
for sh in ag.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip() in ("Situational & SWOT Analysis", "AI Strategy Formulation", "Implementation Roadmap"):
        sh.left = Emu(6795700); sh.width = Emu(4485523); sh.text_frame.word_wrap = True
        for p_ in sh.text_frame.paragraphs: p_.alignment = PP_ALIGN.LEFT
        for r_ in [r for p_ in sh.text_frame.paragraphs for r in p_.runs]: r_.font.size = Pt(18)
    if sh.has_text_frame and "Grant Thornton Bahrain" in sh.text_frame.text and "\u00a9" in sh.text_frame.text:
        set_text(sh, "\u00a9 2026 Grant Thornton Bahrain. All rights reserved")
NEW.append(ag)

# =========================================================== 3 executive summary
s = content("Executive Summary", "Grant Thornton Bahrain will become the Kingdom's most trusted advisor for responsible, value-driven AI - and an AI-enabled firm that practises what it advises.")
cols = [("WHERE WE ARE", GT, ["AI Ready7 self-assessment (Sept 2026): 61.4% - 'Informed' maturity", "Strengths: regulatory compliance, AI inventory, data pipelines & lineage, literacy, threat defence",
                              "Critical gaps: lifecycle AI risk assessment (2.4) and data classification for AI (5.1) - both Not Met", "Strong market pull: Vision 2030, iGA AI Policy (2025), PDPL, CBB, EU AI Act reach"]),
        ("WHERE WE ARE GOING", GT2, ["Vision 2028: trusted AI advisor + AI-enabled firm", "Five strategic pillars: Trusted & Governed AI, AI-Fluent People, The Intelligent Firm, Secure AI Foundations, Market-Leading AI Advisory",
                                   "Target readiness >= 75% ('Repeatable') by end-2027 and >= 85% by end-2028", "Clear AI risk appetite: very low for client-data & compliance risk"]),
        ("HOW WE WILL GET THERE", TEAL, ["18-month roadmap in three horizons: Foundations (Q4-26 to Q1-27), Scale (2027), Differentiate (2028)", "5 workstreams, 21 activities; 100-day plan ready to start",
                                        "12 use cases prioritised; 4 launched in Horizon 1", "4 AI service lines packaged into entry offers built on GT tools (RCM v2.0, risk framework, readiness tool)"])]
cw = (CW - 2 * 150000) // 3
for i, (h, col, items) in enumerate(cols):
    x = L + i * (cw + 150000)
    box(s, x, TOP, cw, 420000, fill=col, text=h, size=13, bold=True, color=WHITE)
    box(s, x, TOP + 420000, cw, 3900000, fill=LAV, text=None)
    txt(s, x + 60000, TOP + 540000, cw - 120000, 3800000, items, size=13, bullet=True, spacing=12)
box(s, L, TOP + 4450000, CW, 520000, fill=DARK, text=[[("Decisions requested: ", True), ("approve vision, objectives and risk appetite; appoint AI sponsor and CoE lead; release Horizon 1 resources; endorse quick-win pilots and the service-line launch.", False)]], size=11, color=WHITE)
NEW.append(s)

# =========================================================== 4 approach
s = content("Our Approach", "A three-phase, standards-based methodology - the same approach GT Bahrain offers clients through the AI Strategy & Roadmap service line.")
phases = [("01", "Situational / SWOT Analysis", ["External scan: regulation, market, competitors, technology", "Internal baseline: AI Ready7 & GT AI Readiness Assessment", "Stakeholder interviews and AI inventory", "SWOT and TOWS strategic options"],
           "Outputs: baseline report, SWOT / TOWS, gap list"),
          ("02", "AI Strategy Formulation", ["Vision, ambition and strategic objectives", "Strategic pillars and guiding principles", "AI risk appetite & guardrails (RCM RM-6)", "Use-case prioritisation (value x feasibility x risk)", "Operating model and KPIs"],
           "Outputs: AI strategy, risk appetite, prioritised portfolio"),
          ("03", "Roadmap Creation", ["Phased horizons and workstreams", "Dependencies, resourcing and budget envelope", "100-day mobilisation plan", "Governance cadence and KPI tracking", "Traceability to readiness gaps"],
           "Outputs: 18-month roadmap, 100-day plan, KPI dashboard")]
pw = (CW - 2 * 200000) // 3
for i, (n, t, items, out) in enumerate(phases):
    x = L + i * (pw + 200000)
    box(s, x, TOP, pw, 700000, fill=GT if i != 1 else GT2, shape=MSO_SHAPE.CHEVRON if False else MSO_SHAPE.RECTANGLE,
        text=[[(n + "  ", True), (t, True)]], size=15, color=WHITE)
    box(s, x, TOP + 700000, pw, 2750000, fill=LAV)
    txt(s, x + 60000, TOP + 800000, pw - 120000, 2600000, items, size=13, bullet=True, spacing=10)
    box(s, x, TOP + 3500000, pw, 520000, fill=WHITE, line=GT, text=out, size=10, bold=True, color=GT)
box(s, L, TOP + 4200000, CW, 700000, fill=DARK,
    text=[[("GT tools used: ", True), ("GT AI Ready7  |  GT AI Readiness & Risk Assessment Tool  |  GT AI RCM v2.0 (72 controls)  |  GT AI Risk Assessment Framework & Toolkit.  ", False),
           ("Standards: ", True), ("ISO/IEC 42001, ISO/IEC 23894, NIST AI RMF, EU AI Act, Bahrain PDPL, iGA AI Policy.", False)]], size=10.5, color=WHITE, align=PP_ALIGN.LEFT)
NEW.append(s)

# =========================================================== 5 divider
NEW.append(divider("Situational &\nSWOT Analysis", 1))

# =========================================================== 6 external landscape
s = content("Bahrain AI Landscape - External Scan", "Policy momentum, a tightening regulatory perimeter and sector demand are creating a window for trusted, governance-led AI advice.")
quads = [("Policy & Regulation", ["Bahrain Economic Vision 2030 positions AI as a productivity and diversification lever", "iGA General Policy for the Use of AI v1.0 (2025): compliance, adoption, awareness, cooperation", "GCC Guiding Manual on the Ethics of AI Use adopted (2025)",
                                  "Draft standalone AI law (38 articles) approved by Shura Council (Apr 2024) - pending", "PDPL (Law 30/2018), CBB Rulebook, NCSC standards; EU AI Act applies extraterritorially"]),
         ("Economy & Market", ["ICT revenue opportunity estimated at US$27.15bn cumulatively (2022-2026)", "IT services market CAGR 7.03% (2024-2028) to US$201.8m", "Priority sectors: financial services, insurance, government, telecom, hospitality, healthcare",
                                "Clients shifting from AI experimentation to scaled, governed deployment"]),
         ("Society & Skills", ["UNESCO AI Readiness Assessment of Bahrain (Nov 2025) - 2nd GCC country assessed", "Recommendations: embed AI ethics, expand capacity building, transparency and public engagement",
                               "National skills initiatives (e.g., AI Academy); competition for AI talent across the GCC"]),
         ("Technology", ["Hyperscale cloud region in Bahrain and cloud-first policy", "GenAI and agentic AI entering enterprise software (Copilot-class tools)", "AI-enabled threats rising: deepfake fraud, AI-crafted phishing, automated exploitation",
                          "ISO/IEC 42001 creating an AI assurance and certification market"])]
qw = (CW - 150000) // 2; qh = 2400000
for i, (h, items) in enumerate(quads):
    x = L + (i % 2) * (qw + 150000); y = TOP + (i // 2) * (qh + 120000)
    box(s, x, y, 260000, qh, fill=[GT, GT2, TEAL, CORAL][i])
    txt(s, x + 330000, y + 30000, qw - 360000, 360000, [[(h, True)]], size=13, color=[GT, GT2, TEAL, CORAL][i])
    txt(s, x + 330000, y + 380000, qw - 360000, qh - 400000, items, size=11.5, bullet=True, spacing=5)
txt(s, L, 6250000, CW, 250000, "Sources: GT AI Service Lines deck v3.0 (market data); iGA; UNESCO; Shura Council; GT research, Sept 2026.", size=8, color=GREY)
NEW.append(s)

# =========================================================== 7 competition & positioning
s = content("Market & Competitive Positioning", "Competitors lead with technology; GT Bahrain can own the 'trusted, governance-led AI' space for regulated and mid-market clients.")
comp = [["Competitor type", "Examples (Bahrain / MENA)", "Typical focus", "Implication for GT"],
        ["Big Four peers", "EY (EY.ai platform, AI strategy & design)", "Large-scale transformation, own platforms", "Differentiate on agility, mid-market fit and GRC + cyber depth"],
        ["AI consultancies & trainers", "AI Enabled, AI Superior", "Strategy workshops, training, R&D", "Compete with certified methodology and assurance credibility"],
        ["Automation / software houses", "10xDS, Sanara Infotech, Abacus KSA, Purpix", "RPA, virtual agents, custom build", "Partner for build; GT leads governance, selection and assurance"],
        ["Hyperscalers & vendors", "Cloud and GenAI platform providers", "Tools and platforms", "Stay vendor-agnostic; advise on selection and secure adoption"]]
table(s, L, TOP, int(CW * 0.62), comp, [2.2, 3, 2.8, 3.6], size=10.5, rowh=720000)
x2 = L + int(CW * 0.62) + 200000; w2 = R - x2
box(s, x2, TOP, w2, 420000, fill=GT, text="GT Bahrain's differentiated position", size=12, bold=True, color=WHITE)
pos = [[("Trust-first: ", True), ("audit, GRC and cyber heritage; AI Ready7 & ISO/IEC 42001-aligned methods", False)],
       [("Regulation-ready: ", True), ("Bahrain PDPL, CBB, EU AI Act and GCC mapped into one RCM (72 controls)", False)],
       [("End-to-end: ", True), ("strategy -> governance -> automation -> training -> assurance", False)],
       [("Mid-market & regulated focus: ", True), ("banks, insurers, government, telecom, healthcare, hospitality", False)],
       [("Proof: ", True), ("GT Bahrain as 'client zero' with measured readiness uplift", False)]]
box(s, x2, TOP + 420000, w2, 3300000, fill=LAV)
txt(s, x2 + 60000, TOP + 500000, w2 - 120000, 3200000, pos, size=12, bullet=False, spacing=12)
NEW.append(s)

# =========================================================== 8 internal baseline
s = content("Internal Baseline - AI Ready7 Results (Sept 2026)", "Overall readiness 61.4% ('Informed'): solid compliance and data foundations; business alignment, architecture and AI security need the most work.")
cd = CategoryChartData(); cd.categories = [p for p, _ in READY7][::-1]
cd.add_series("Readiness %", [v for _, v in READY7][::-1])
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Emu(L), Emu(TOP), Emu(int(CW * 0.55)), Emu(4400000), cd)
ch = gf.chart; ch.has_legend = False; ch.has_title = False
pl = ch.plots[0]; pl.gap_width = 60; pl.has_data_labels = True
pl.data_labels.number_format = '0"%"'; pl.data_labels.number_format_is_linked = False; pl.data_labels.font.size = Pt(10); pl.data_labels.font.bold = True
pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
ser = pl.series[0]; ser.format.fill.solid(); ser.format.fill.fore_color.rgb = GT
for idx, (_, v) in enumerate(READY7[::-1]):
    pt = ser.points[idx]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = GT if v >= 60 else CORAL
ch.value_axis.maximum_scale = 100; ch.value_axis.minimum_scale = 0; ch.value_axis.major_unit = 25
ch.value_axis.tick_labels.font.size = Pt(9); ch.category_axis.tick_labels.font.size = Pt(10)
ch.value_axis.has_major_gridlines = True; ch.value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xDD, 0xDD, 0xDD)
txt(s, L, TOP + 4420000, int(CW * 0.55), 300000, "Maturity scale: <50% Partial | 50-70% Informed | 70-90% Repeatable | 90%+ Adaptive.  Coral = below 60%.", size=8.5, color=GREY)
x2 = L + int(CW * 0.57); w2 = R - x2
box(s, x2, TOP, w2, 380000, fill=GREEN, text="Strengths (All / Most Met)", size=11.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
txt(s, x2, TOP + 400000, w2, 1500000, ["2.2 Regulatory compliance & audit readiness; 2.5 AI inventory & ownership", "5.3 Secure data pipelines; 5.4 Data lineage & quality", "3.1-3.3 AI literacy, validation & security awareness",
                                       "4.5 Cost & platform governance; 6.1 Identity & privilege", "7.1, 7.2, 7.4, 7.5 AI threat intel, SOC, IR, exposure mgmt"], size=10, bullet=True, spacing=3)
box(s, x2, TOP + 1980000, w2, 380000, fill=CORAL, text="Gaps (Not Met / Some Met)", size=11.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
txt(s, x2, TOP + 2380000, w2, 2000000, [[("Not Met: ", True), ("2.4 Lifecycle AI risk assessment incl. third parties; 5.1 Data classification, access & handling for AI", False)],
                                       "1.1-1.4 Strategy, risk appetite, portfolio, value tracking, client assurance", "2.1 AI policy; 2.3 Guardrails / centre of excellence; 3.4 Role-based capability",
                                       "4.1-4.4 Reference architecture, execution authority, API control, monitoring", "6.2-6.5 Model protection, secure ModelOps, red teaming, runtime logging; 7.3 Deepfake defence"], size=10, bullet=True, spacing=3)
NEW.append(s)

# =========================================================== 9 SWOT
s = content("SWOT Analysis", "Internal factors from the AI Ready7 baseline; external factors from the Bahrain / GCC landscape scan.")
qw = (CW - 120000) // 2; qh = 2330000
spec = [("STRENGTHS", "S", GT), ("WEAKNESSES", "W", GT2), ("OPPORTUNITIES", "O", TEAL), ("THREATS", "T", CORAL)]
for i, (h, k, col) in enumerate(spec):
    x = L + (i % 2) * (qw + 120000); y = TOP - 60000 + (i // 2) * (qh + 90000)
    box(s, x, y, qw, 340000, fill=col, text=h, size=12, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    box(s, x, y + 340000, qw, qh - 340000, fill=WHITE, line=col)
    txt(s, x + 50000, y + 380000, qw - 100000, qh - 400000, SWOT[k], size=11, bullet=True, spacing=3)
NEW.append(s)

# =========================================================== 10 TOWS
s = content("From SWOT to Strategic Choices (TOWS)", "Four strategic thrusts translate the SWOT into the strategy; each is reflected in a strategic pillar and roadmap workstream.")
tw = (CW - 150000) // 2; th_ = 1900000
cols_ = [GT, GT2, TEAL, CORAL]
maps = ["Pillar 5 | WS5", "Pillars 1-3 | WS1, WS2, WS4", "Pillars 4-5 | WS3, WS5", "Pillars 1, 2, 4 | WS1-WS3"]
for i, (h, d) in enumerate(TOWS):
    x = L + (i % 2) * (tw + 150000); y = TOP + (i // 2) * (th_ + 150000)
    box(s, x, y, tw, th_, fill=LAV)
    box(s, x, y, 120000, th_, fill=cols_[i])
    txt(s, x + 220000, y + 80000, tw - 300000, 400000, [[(h, True)]], size=14, color=cols_[i])
    txt(s, x + 220000, y + 520000, tw - 300000, 1000000, d, size=13)
    txt(s, x + 220000, y + th_ - 380000, tw - 300000, 300000, [[("Delivered through: " + maps[i], True)]], size=9.5, color=GREY)
NEW.append(s)

# =========================================================== 11 divider
NEW.append(divider("AI Strategy\nFormulation", 2))

# =========================================================== 12 vision & objectives
s = content("Vision, Ambition & Strategic Objectives")
box(s, L, TOP - 350000, CW, 900000, fill=GT, text=[[("VISION 2028   ", True), ("By 2028, Grant Thornton Bahrain is the Kingdom's most trusted advisor for responsible, value-driven AI - and an AI-enabled firm that practises what it advises.", False)]],
    size=15, color=WHITE, align=PP_ALIGN.LEFT)
objs = [("1", "An AI-ready firm", ["AI Ready7 readiness from 61.4% to >= 75% (Repeatable) by end-2027 and >= 85% by end-2028", "100% of AI use cases registered, risk-tiered and assessed", "ISO/IEC 42001 certification-ready by end-2027"]),
        ("2", "Productivity & quality through AI", ["10+ production use cases with measured benefits by 2028", "10-15% time saved in targeted processes", "Zero client-data leakage incidents via AI"]),
        ("3", "A market-leading AI advisory practice", ["Four AI service lines launched with productised entry offers", "12 AI engagements in 2027, 25 in 2028 (proposed targets)", "Recognised thought leadership on responsible AI in Bahrain / GCC"])]
ow = (CW - 2 * 150000) // 3
for i, (n, h, items) in enumerate(objs):
    x = L + i * (ow + 150000); y = TOP + 700000
    box(s, x, y, ow, 600000, fill=[GT2, TEAL, CORAL][i], text=[[(f"Objective {n}  ", True), (h, True)]], size=13, color=WHITE)
    box(s, x, y + 600000, ow, 2500000, fill=LAV)
    txt(s, x + 60000, y + 700000, ow - 120000, 2400000, items, size=13, bullet=True, spacing=14)
txt(s, L, TOP + 3900000, CW, 400000, "Targets are proposed for leadership validation; baselines marked 'to be baselined' in the KPI dashboard will be measured in the first 100 days.", size=9, color=GREY)
NEW.append(s)

# =========================================================== 13 pillars & principles
s = content("Strategic Pillars & Guiding Principles", "Five pillars close the readiness gaps and build the advisory business - underpinned by six non-negotiable principles.")
pw = (CW - 4 * 100000) // 5
for i, (n, h, d, link) in enumerate(PILLARS):
    x = L + i * (pw + 100000)
    col = [GT, GT2, TEAL, DARK, CORAL][i]
    box(s, x, TOP, pw, 750000, fill=col, text=[[(h, True)]], size=13, color=WHITE)
    box(s, x, TOP + 750000, pw, 1700000, fill=LAV, text=d, size=12, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
    box(s, x, TOP + 2450000, pw, 380000, fill=WHITE, line=col, text=link, size=8.5, color=col, bold=True)
box(s, L, TOP + 2950000, CW, 360000, fill=DARK, text="GUIDING PRINCIPLES", size=11.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
half = (CW - 100000) // 2
txt(s, L, TOP + 3380000, half, 1500000, PRINCIPLES[:3], size=12, bullet=True, spacing=7)
txt(s, L + half + 100000, TOP + 3380000, half, 1500000, PRINCIPLES[3:], size=12, bullet=True, spacing=7)
NEW.append(s)

# =========================================================== 14 risk appetite
s = content("AI Risk Appetite & Guardrails", "A clear appetite enables confident experimentation while protecting what matters most: client trust, confidentiality and professional standards (RCM RM-6).")
rows = [["Risk category", "Appetite", "Statement / guardrail"]] + [[a, b, c] for a, b, c in APPETITE]
fills = {}
for i, (_, a, _) in enumerate(APPETITE, 1):
    fills[(i, 1)] = {"Very low": RGBColor(0xF4, 0xB1, 0x83), "Low": RGBColor(0xFF, 0xE6, 0x99), "Moderate": RGBColor(0xC6, 0xE0, 0xB4), "Moderate-High": RGBColor(0x9B, 0xC2, 0xE6)}[a]
table(s, L, TOP, CW, rows, [3, 1.3, 7.5], size=12, rowh=600000, fills=fills)
box(s, L, TOP + 3550000, CW, 800000, fill=LAV, text=[[("Operationalised through: ", True), ("AI acceptable-use policy (GL-5) | AI register & risk tiering (GL-8, RM-5) | GT AI Risk Assessment Framework (approval by tier) | KRIs with mandatory responses (RM-6) | human-in-charge sign-off for client deliverables (RS-1).", False)]],
    size=10.5, align=PP_ALIGN.LEFT)
NEW.append(s)

# =========================================================== 15 use-case prioritisation chart
s = content("Use-Case Portfolio Prioritisation", "12 candidate use cases scored on business value and feasibility (data readiness, technical complexity, change effort); bubble size reflects risk tier.")
bd = BubbleChartData()
size_map = {"Low": 6, "Medium": 10, "High": 16}
colmap = {"Low": GREEN, "Medium": AMBER, "High": CORAL}
ser_ = bd.add_series("Use cases")
for uc in USE_CASES: ser_.add_data_point(uc[3], uc[2], size_map[uc[4]])
gf = s.shapes.add_chart(XL_CHART_TYPE.BUBBLE, Emu(L), Emu(TOP - 100000), Emu(int(CW * 0.58)), Emu(4700000), bd)
ch = gf.chart
ch.has_legend = False; ch.has_title = False
sr = ch.plots[0].series[0]
for pi, uc in enumerate(USE_CASES):
    pt = sr.points[pi]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = colmap[uc[4]]
    dl = pt.data_label; dl.has_text_frame = True; dl.text_frame.text = str(uc[0]); dl.position = XL_LABEL_POSITION.CENTER
    dl.text_frame.paragraphs[0].runs[0].font.size = Pt(9); dl.text_frame.paragraphs[0].runs[0].font.bold = True
lx = L + 300000
for i, (lab, col) in enumerate([("Low risk", GREEN), ("Medium risk", AMBER), ("High risk", CORAL)]):
    box(s, lx + i * 1300000, TOP + 4650000, 160000, 160000, fill=col, shape=MSO_SHAPE.OVAL)
    txt(s, lx + i * 1300000 + 190000, TOP + 4600000, 1000000, 250000, lab, size=9)
ch.plots[0].bubble_scale = 35
va, ca = ch.value_axis, ch.category_axis
for ax, title in ((va, "Business value"), (ca, "Feasibility")):
    ax.minimum_scale = 1.5; ax.maximum_scale = 5.2; ax.major_unit = 0.5; ax.has_title = True
    ax.axis_title.text_frame.text = title; ax.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(10)
    ax.tick_labels.font.size = Pt(8)
va.has_major_gridlines = False
va.crosses_at = 3.5; ca.crosses_at = 3.5
va.tick_label_position = XL_TICK_LABEL_POSITION.LOW; ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
cx = L + int(CW * 0.58)
# quadrant labels overlay
for lab, fx, fy, col in [("STRATEGIC BETS", 0.12, 0.03, GT2), ("QUICK WINS", 0.78, 0.03, GREEN), ("DEPRIORITISE", 0.12, 0.80, GREY), ("FILL-INS", 0.80, 0.80, TEAL)]:
    txt(s, L + int(CW * 0.58 * fx), TOP + int(4400000 * fy), 1400000, 260000, [[(lab, True)]], size=9, color=col)
x2 = cx + 150000; w2 = R - x2
rows = [["#", "Use case", "V", "F", "Quadrant", "H"]] + [[uc[0], uc[1], f"{uc[2]:.2f}", f"{uc[3]:.2f}", quadrant(uc[2], uc[3]), uc[5]] for uc in sorted(USE_CASES, key=lambda u: -(u[2] + u[3]))]
table(s, x2, TOP - 100000, w2, rows, [0.35, 3.6, 0.6, 0.6, 1.3, 0.45], size=8.5, rowh=355000)
txt(s, x2, TOP + 4550000, w2, 300000, "V = value, F = feasibility (weighted 1-5 scores - see Strategy Toolkit), H = horizon. Quadrant threshold 3.5.", size=8, color=GREY)
NEW.append(s)

# =========================================================== 16 priority use cases
s = content("Priority Use Cases", "Horizon 1 quick wins build confidence and adoption; Horizon 2 pilots target the highest-value professional-services processes.")
pick = [u for u in USE_CASES if u[0] in (1, 3, 2, 4, 5, 8)]
rows = [["Use case", "Area", "Description", "Risk tier", "Horizon"]] + [[u[1], u[6], u[7], u[4], u[5]] for u in pick]
fills = {(i, 3): {"Low": RGBColor(0xC6, 0xE0, 0xB4), "Medium": RGBColor(0xFF, 0xE6, 0x99), "High": RGBColor(0xF4, 0xB1, 0x83)}[u[4]] for i, u in enumerate(pick, 1)}
table(s, L, TOP, CW, rows, [2.6, 1.8, 6, 1, 0.9], size=10, rowh=560000, fills=fills)
box(s, L, TOP + 4080000, CW, 700000, fill=LAV, text=[[("Every use case follows the GT AI Risk Assessment Framework: ", True), ("prohibited-practice screen -> tiering -> risk register -> approval by tier (High = AI Committee). Benefit KPIs and a business owner are set before build.", False)]],
    size=10.5, align=PP_ALIGN.LEFT)
NEW.append(s)

# =========================================================== 17 service line GTM
s = content("AI Advisory Service Lines - Go-to-Market", "Productised offerings built on GT tools, launched through entry offers and sector campaigns, and scaled via the GT network.")
ow = (CW - 3 * 100000) // 4
for i, (name, items, sectors) in enumerate(OFFERINGS):
    x = L + i * (ow + 100000); col = [GT, GT2, TEAL, CORAL][i]
    box(s, x, TOP, ow, 560000, fill=col, text=name, size=12, bold=True, color=WHITE)
    box(s, x, TOP + 560000, ow, 2500000, fill=LAV)
    txt(s, x + 50000, TOP + 640000, ow - 100000, 2400000, items, size=11.5, bullet=True, spacing=7)
    box(s, x, TOP + 3060000, ow, 460000, fill=WHITE, line=col, text=[[("Target: ", True), (sectors, False)]], size=9, color=col)
box(s, L, TOP + 3650000, CW, 850000, fill=DARK, text=[
    [("Entry offers: ", True), ("AI Readiness & Risk Assessment (2-3 weeks)  |  AI GRC / RCM gap review  |  AI strategy sprint  |  Board AI governance briefing", False)],
    [("Channels: ", True), ("existing audit / advisory client base, GTIL AI Ready7 network, regulator & industry events, thought leadership on PDPL / EU AI Act / ISO/IEC 42001", False)]],
    size=10.5, color=WHITE, align=PP_ALIGN.LEFT)
NEW.append(s)

# =========================================================== 18 operating model
s = content("Operating Model & Governance", "A light, federated model: central CoE for standards and assurance; champions embedded in every service line.")
bw = 3400000
box(s, L + (CW - bw) // 2, TOP - 100000, bw, 480000, fill=DARK, text="Managing Partner & Leadership Team", size=12, bold=True, color=WHITE)
box(s, L + (CW - bw) // 2 - 700000, TOP + 480000, bw + 1400000, 620000, fill=GT, text=[[("AI Steering Committee", True)], [("approves strategy, risk appetite and High / Critical use cases", False)]], size=10.5, color=WHITE)
cw_ = (CW - 3 * 120000) // 4
coe = [("AI CoE lead", "Strategy, portfolio, KPIs, standards"), ("AI GRC lead", "Policy, register, risk framework, ISO/IEC 42001"), ("Data & technology lead", "Platforms, architecture, ModelOps, monitoring"), ("Cyber / AI security lead", "Red teaming, guardrails, deepfake defence")]
box(s, L, TOP + 1200000, CW, 340000, fill=GT2, text="AI Centre of Excellence (CoE)", size=11.5, bold=True, color=WHITE)
for i, (h, d) in enumerate(coe):
    x = L + i * (cw_ + 120000)
    box(s, x, TOP + 1600000, cw_, 800000, fill=LAV, line=GT2, text=[[(h, True)], [(d, False)]], size=11)
box(s, L, TOP + 2550000, CW, 340000, fill=TEAL, text="AI Champions network - Audit | Tax | Advisory | Risk | Operations", size=11.5, bold=True, color=WHITE)
cad = [["Forum", "Cadence", "Key inputs / outputs"],
       ["AI Steering Committee", "Quarterly (monthly in H1)", "Roadmap status, KPI dashboard, risk profile, approvals"],
       ["AI CoE working group", "Bi-weekly", "Use-case intake, guardrails, incidents, vendor changes"],
       ["Champions forum", "Monthly", "Adoption, feedback, training needs, new ideas"],
       ["Link to GTIL", "Ongoing", "AI Ready7 re-assessment, shared tools, global AI policies"]]
table(s, L, TOP + 3000000, CW, cad, [2.5, 2, 6], size=10.5, rowh=340000)
NEW.append(s)

# =========================================================== 19 KPIs
s = content("KPIs & Value Realisation", "Progress is measured against the AI Ready7 baseline and a balanced set of adoption, value, risk and growth indicators.")
rows = [["KPI", "Baseline (Sept 2026)", "Target end-2027", "Target end-2028", "Owner"]] + [list(k) for k in KPIS]
table(s, L, TOP, CW, rows, [4, 2.6, 2.3, 2.3, 2], size=11.5, rowh=480000)
NEW.append(s)

# =========================================================== 20 divider
NEW.append(divider("Implementation\nRoadmap", 3))

# =========================================================== 21 roadmap on a page
s = content("Roadmap on a Page", "Three horizons move GT Bahrain from 'Informed' to 'Repeatable' and then to a differentiated, AI-enabled advisory firm.")
hz = [("HORIZON 1  |  FOUNDATIONS", "Q4 2026 - Q1 2027", GT, ["Strategy, risk appetite, AI policy & CoE", "AI register refresh; risk framework on every use case", "Data classification & DLP for GenAI", "AI literacy for all staff", "Quick wins: drafting, transcription, knowledge assistant", "Launch entry offers; first client pilots"],
       "Exit: readiness ~68%; no Not-Met items"),
      ("HORIZON 2  |  SCALE", "Q2 - Q4 2027", GT2, ["Reference architecture, monitoring & ModelOps", "AI red teaming; deepfake-resilient processes", "Pilots: audit analytics, tax assistant, document review", "Role-based pathways & champions", "Sector campaigns (FS, government, telecom)", "ISO/IEC 42001 AIMS build"],
       "Exit: readiness >= 75% (Repeatable)"),
      ("HORIZON 3  |  DIFFERENTIATE", "2028", TEAL, ["Scale proven use cases; agentic pilots", "ISO/IEC 42001 certification (optional)", "Client self-assessment portal", "Regional delivery via GT network", "Thought leadership on responsible AI"],
       "Exit: readiness >= 85%; recognised AI advisor")]
hw = (CW - 2 * 120000) // 3
for i, (h, when, col, items, ex) in enumerate(hz):
    x = L + i * (hw + 120000)
    box(s, x, TOP, hw, 520000, fill=col, shape=MSO_SHAPE.PENTAGON if i < 2 else MSO_SHAPE.RECTANGLE, text=[[(h, True)]], size=12, color=WHITE)
    txt(s, x, TOP + 540000, hw, 300000, [[(when, True)]], size=10.5, color=col)
    box(s, x, TOP + 860000, hw, 2900000, fill=LAV)
    txt(s, x + 60000, TOP + 940000, hw - 120000, 2800000, items, size=12.5, bullet=True, spacing=10)
    box(s, x, TOP + 3820000, hw, 480000, fill=WHITE, line=col, text=ex, size=10, bold=True, color=col)
NEW.append(s)

# =========================================================== 22 Gantt
s = content("Detailed Roadmap - Workstreams & Timeline", "18 months, five workstreams, 21 activities (Oct 2026 - Mar 2028).")
lab_w = 3500000; gx = L + lab_w; gw = R - gx; months = 18; mw = gw / months
qlabels = ["Q4 2026", "Q1 2027", "Q2 2027", "Q3 2027", "Q4 2027", "Q1 2028"]
y0 = TOP - 230000
for q in range(6):
    box(s, int(gx + q * 3 * mw), y0, int(3 * mw) - 8000, 260000, fill=GT if q % 2 == 0 else GT2, text=qlabels[q], size=9, bold=True, color=WHITE)
for hx, lab in [(0, "H1 Foundations"), (6, "H2 Scale"), (15, "H3 Differentiate")]:
    pass
rowh = 164000; y = y0 + 285000
wscol = {"WS1": GT, "WS2": TEAL, "WS3": DARK, "WS4": CORAL, "WS5": GT2}
current = None
for ws_, act, st, du in GANTT:
    if ws_ != current:
        current = ws_
        txt(s, L, y - 10000, lab_w, rowh, [[(ws_, True)]], size=8.5, color=wscol[ws_[:3]])
        y += rowh
    txt(s, L + 120000, y - 12000, lab_w - 120000, rowh, act, size=7.5)
    for m in range(0, months, 3):
        pass
    bar = box(s, int(gx + st * mw), y + 25000, int(du * mw) - 10000, rowh - 50000, fill=wscol[ws_[:3]])
    y += rowh
for m in (6, 15):
    ln = s.shapes.add_connector(1, Emu(int(gx + m * mw)), Emu(y0 + 270000), Emu(int(gx + m * mw)), Emu(y))
    ln.line.color.rgb = TEAL; ln.line.width = Pt(1.25); ln.line.dash_style = 4
txt(s, L, y + 20000, lab_w, 220000, [[("Dashed lines: start of H2 (Apr 2027) and H3 (Jan 2028)", True)]], size=8, color=TEAL)
NEW.append(s)

# =========================================================== 23 100-day plan
s = content("First 100 Days - Mobilisation Plan", "Concrete actions and owners to start immediately after approval.")
cw_ = (CW - 2 * 120000) // 3
for i, (phase, acts) in enumerate(HUNDRED_DAYS):
    x = L + i * (cw_ + 120000); col = [GT, GT2, TEAL][i]
    box(s, x, TOP, cw_, 480000, fill=col, text=phase, size=14, bold=True, color=WHITE)
    for j, (a, o) in enumerate(acts):
        yy = TOP + 560000 + j * 930000
        box(s, x, yy, cw_, 860000, fill=LAV, line=None, text=[[(a, True)], [("Owner: " + o, False)]], size=12, align=PP_ALIGN.LEFT)
box(s, L, TOP + 4350000, CW, 500000, fill=DARK, text="Day-100 checkpoint: AI Ready7 re-run on Governance & Ethics and Data & Quality items; target no 'Not Met' items and readiness ~65%.", size=11, bold=True, color=WHITE)
NEW.append(s)

# =========================================================== 24 resourcing
s = content("Resourcing, Investment & Dependencies", "Indicative requirements - to be sized in the Horizon 1 business case.")
res = [("People", GT, ["AI service line lead (partner / director sponsor)", "AI CoE lead and 2 AI GRC specialists", "1-2 data / AI engineers (or partner capacity)", "Cyber AI-security specialist (shared with Cyber team)", "AI champion in each service line (10-20% time)"]),
       ("Platforms & tools", GT2, ["Enterprise GenAI suite with no-training and data-residency terms", "Knowledge / RAG platform over GT methodologies", "GRC tooling for AI register & assessments (OneTrust already used for AI Ready7)", "DLP / CASB for GenAI; AI logging & monitoring", "Workstations per GT AI system-setup specification"]),
       ("Partners & ecosystem", TEAL, ["GTIL AI and cyber networks (AI Ready7, shared tools)", "Technology partners for build and integration", "Training partners / academia for AI literacy content", "Certification bodies for ISO/IEC 42001 readiness"])]
cw_ = (CW - 2 * 120000) // 3
for i, (h, col, items) in enumerate(res):
    x = L + i * (cw_ + 120000)
    box(s, x, TOP, cw_, 460000, fill=col, text=h, size=13, bold=True, color=WHITE)
    box(s, x, TOP + 460000, cw_, 2700000, fill=LAV)
    txt(s, x + 60000, TOP + 540000, cw_ - 120000, 2600000, items, size=12, bullet=True, spacing=9)
box(s, L, TOP + 3300000, CW, 1000000, fill=WHITE, line=GT, text=[[("Key dependencies: ", True), ("leadership sponsorship and time of champions; IT capacity for platform and DLP changes; client consent terms for AI use on engagements; GTIL global AI policies and approved-tool lists; budget release for Horizon 1.", False)]],
    size=10.5, align=PP_ALIGN.LEFT)
NEW.append(s)

# =========================================================== 25 strategy risks
s = content("Risks to the Strategy & Mitigations")
rows = [["Risk", "Mitigation"]] + [list(r) for r in STRAT_RISKS]
table(s, L, TOP - 250000, CW, rows, [3.5, 8], size=13, rowh=600000)
NEW.append(s)

# =========================================================== 26 next steps
s = content("Decisions Required & Next Steps")
dec = ["Approve the AI vision, three strategic objectives and the AI risk appetite", "Appoint the executive AI sponsor and the AI CoE lead", "Release Horizon 1 resources and budget envelope",
       "Endorse the four Horizon 1 use cases for pilot", "Endorse the AI service-line launch plan and entry offers"]
hw = (CW - 150000) // 2
box(s, L, TOP - 200000, hw, 460000, fill=GT, text="Decisions required from leadership", size=13, bold=True, color=WHITE)
for i, d in enumerate(dec):
    yy = TOP + 330000 + i * 720000
    box(s, L, yy, 520000, 620000, fill=GT2, text=str(i + 1), size=18, bold=True, color=WHITE)
    box(s, L + 560000, yy, hw - 560000, 620000, fill=LAV, text=d, size=12.5, align=PP_ALIGN.LEFT)
nx = ["Week 1-2: leadership review of this strategy and risk appetite", "Week 2: issue interim AI acceptable-use policy", "Week 3: CoE mobilised; 100-day plan kicked off",
      "Week 4: quick-win pilot scoping and baseline KPIs", "Day 100: progress checkpoint to the Steering Committee"]
x2 = L + hw + 150000
box(s, x2, TOP - 200000, hw, 460000, fill=TEAL, text="Next steps", size=13, bold=True, color=WHITE)
box(s, x2, TOP + 330000, hw, 3500000, fill=LAV)
txt(s, x2 + 80000, TOP + 420000, hw - 160000, 3400000, nx, size=13, bullet=True, spacing=18)
NEW.append(s)

# =========================================================== 27 appendix traceability
s = content("Appendix - AI Ready7 Gap-to-Roadmap Traceability", "Every item not fully met in the Sept 2026 AI Ready7 assessment is addressed by a roadmap action.")
rows = [["Item", "Title", "Result", "Roadmap response", "Horizon", "WS"]] + [list(r) for r in READY7_ITEMS]
fills = {(i, 2): {"Not Met": RGBColor(0xF4, 0xB1, 0x83), "Some Met": RGBColor(0xFF, 0xE6, 0x99), "Most Met": RGBColor(0xC6, 0xE0, 0xB4)}[r[2]] for i, r in enumerate(READY7_ITEMS, 1)}
table(s, L, TOP - 300000, CW, rows, [0.9, 3.2, 1, 5.2, 0.8, 0.6], size=7.5, rowh=205000, fills=fills)
NEW.append(s)

# =========================================================== 28 thank you
ty = duplicate(S_THANKS)
t = find(ty, "Grant Thornton Bahrain")
if t: set_text(t, "© 2026 Grant Thornton Bahrain. All rights reserved")
NEW.append(ty)

# ------------------------------------------------------------------ keep only new slides, in order
sldIdLst = prs.slides._sldIdLst
keep = {id(sl._element) for sl in NEW}
idmap = {}
for sldId in list(sldIdLst):
    rId = sldId.get(qn("r:id"))
    part = prs.part.related_part(rId)
    idmap[id(part._element)] = sldId
for sldId in list(sldIdLst):
    sldIdLst.remove(sldId)
for sl in NEW:
    sldIdLst.append(idmap[id(sl._element)])
for rel in list(prs.part.rels.values()):
    if rel.reltype.endswith("/slide") and id(rel.target_part._element) not in keep:
        prs.part.drop_rel(rel.rId)
def fix_year(shapes):
    for sh in shapes:
        if sh.shape_type == 6: fix_year(sh.shapes); continue
        if sh.has_text_frame:
            for p_ in sh.text_frame.paragraphs:
                for r_ in p_.runs:
                    if "2025" in r_.text and "Grant Thornton" in sh.text_frame.text: r_.text = r_.text.replace("2025", "2026")
for sl in NEW: fix_year(sl.shapes)
for master in prs.slide_masters:
    for el in [master._element] + [lay._element for lay in master.slide_layouts]:
        for para in el.iter(qn("a:p")):
            ts = list(para.iter(qn("a:t")))
            if "Grant Thornton Bahrain" in "".join(t.text or "" for t in ts):
                for t in ts:
                    if t.text:
                        for yr in ("2024", "2025"): t.text = t.text.replace(yr, "2026")
prs.save(OUT)
print("saved", OUT, len(NEW), "slides")
