# -*- coding: utf-8 -*-
"""Microsoft Forms build pack for the GT AI Readiness & Risk Assessment.

Microsoft Forms has no public create-form API, so this script prepares everything needed to stand the form up in
minutes and to turn each response into an engine-ready input:

  online_form/
    GT_AI_Readiness_Form_QuickImport.docx     numbered questions + lettered options, in form order (Forms > Quick import)
    GT_AI_Readiness_Form_Build_Guide.docx     page-by-page layout, copy-paste text, theme and settings, how to process responses
    GT_AI_Readiness_Form_Response_Converter.xlsx   paste a Forms export -> engine-ready blocks (no scoring logic inside)
    form_images/                              lavender background + cover banner
    form_spec.json                            machine-readable form definition

Usage: python build_forms_pack.py QUESTIONNAIRE.docx OUT_DIR
The question text and options are read from the questionnaire Word file, so the form always matches it.
"""
import sys, os, re, json
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from PIL import Image, ImageDraw, ImageFilter, ImageFont

SRC, OUT = sys.argv[1], sys.argv[2]
IMG = os.path.join(OUT, "form_images")
os.makedirs(IMG, exist_ok=True)

# ------------------------------------------------------------------ palette (lavender theme)
LAV_DARK = "3B2A6B"     # headings / text on light
LAV = "7C5CC4"          # theme colour to enter in Forms
LAV_MID = "B7A4E8"
LAV_SOFT = "DDD3F5"
LAV_PALE = "F1ECFC"

# ------------------------------------------------------------------ 1. parse the questionnaire
QRE = re.compile(r"^(\d+)\.\s+\[(P\d+|A\d|Q\d+)\]\s+(.*)$")
doc = Document(SRC)
lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
questions, cur, area = {}, None, None
AREAS = ["AI Governance", "AI Risk & Security", "Training & Awareness", "Strategy & Value", "Data & Technology"]
for ln in lines:
    m = QRE.match(ln)
    if m:
        cur = dict(tag=m.group(2), title=m.group(3), hint="", options=[], area=area, text=False)
        questions[cur["tag"]] = cur
        continue
    if ln in AREAS: area = ln; cur = None; continue
    if cur is None: continue
    if ln.startswith("Answer:"): cur["text"] = True; continue
    mo = re.match(r"^(?:[A-E]|\d+)\.\s+(.*)$", ln)
    if mo and not ln.startswith(("Part ", "Examples", "Consider", "Include", "Answer for", "AIAF")):
        cur["options"].append(mo.group(1)); continue
    if ln.startswith("Part ") or ln in ("Thank you",): cur = None; continue
    cur["hint"] = (cur["hint"] + " " + ln).strip() if not cur["options"] else cur["hint"]
for t, q in questions.items():
    q["area"] = q["area"] or ("Part 1" if t.startswith("P") else "Part 2")
assert len(questions) == 34, len(questions)
for t, q in questions.items():
    assert q["text"] or len(q["options"]) >= 2, t
    assert not (q["text"] and q["options"]), t
# optional field + free text fields
OPTIONAL = {"P5"}

# ------------------------------------------------------------------ 2. form definition
def Q(tag, required=True, text=None, long=False):
    q = questions[tag]
    d = dict(tag=tag, title=q['title'], subtitle=q["hint"], required=required and tag not in OPTIONAL,
             type="text" if q["text"] else "choice", options=q["options"])
    if long: d["type"] = "longtext"
    return d

FORM_TITLE = "AI Readiness & Risk Assessment"
INTRO = ("Grant Thornton Bahrain  |  AI Governance, Risk & Compliance\n\n"
         "This assessment takes about 20-25 minutes. For each question choose ONE answer - the statement that best describes "
         "your organisation today. If you are unsure between two answers, choose the lower one: Grant Thornton may ask for "
         "supporting evidence. Your answers are confidential and used only to prepare your report.\n\n"
         "Your report will be generated and sent to your email within 5 working days of submission.")
WHEN = ("WHEN TO USE THIS TOOL\n"
        "- Before or early in your AI journey, to baseline readiness\n"
        "- Before scaling AI or deploying AI that affects customers or decisions\n"
        "- When AI is discovered in use without assessment (shadow AI)\n"
        "- When risk or context changes: new data, model, vendor, purpose or less human oversight\n"
        "- Annually, to track maturity progress\n\n"
        "Public GenAI tools used for low-risk drafting may be governed by your Acceptable Use Policy instead (AIAF approach).")
GET = ("WHAT YOU GET\n"
       "- Overall AI readiness score and maturity level\n"
       "- Inherent AI risk tier (Low / Medium / High / Critical)\n"
       "- Overall AI risk exposure and recommended oversight pathway\n"
       "- Domain radar chart and responsible-AI principle coverage\n"
       "- Findings, potential risks and recommendations per area\n"
       "- Top-5 priorities and a phased roadmap\n"
       "- Applicable regulations (Bahrain PDPL, CBB, EU AI Act, GCC)\n"
       "- Printable report you can save as PDF and email")
CONFIRM = ("Thank you. Your AI Readiness & Risk Assessment has been submitted. Grant Thornton Bahrain will generate your report "
           "and send it to the email address you provided within 5 working days.")
CHOICE_RULE = ("Choose the ONE statement that best describes the situation today. If unsure between two answers, choose the "
               "lower one.")

PAGES = [
    dict(no=1, section=None, title=FORM_TITLE, desc=INTRO + "\n\n" + WHEN + "\n\n" + GET,
         items=[Q(t) for t in ["P2", "P3", "P4", "P5", "P7", "P9", "P10", "P11", "P12", "P13", "P14"]]),
    dict(no=2, section="Part 2 - How AI is used in your organisation",
         desc="These questions describe your AI use and its potential impact. There are no right or wrong answers.",
         items=[Q(t) for t in ["A1", "A2", "A3", "A4"]]),
]
n = 3
for a in AREAS:
    PAGES.append(dict(no=n, section=f"Part 3 - {a}", desc=CHOICE_RULE,
                      items=[Q(t) for t, q in questions.items() if q["area"] == a])); n += 1
FOLLOW = [
    ("F1", "Would your organisation like to opt for AI training from us - for example AI literacy and security awareness for all staff, "
           "or role-based sessions for leaders, developers and oversight teams? Tell us what you have in mind."),
    ("F2", "Would you like us to review your AI or automation systems (for example AI agents, chatbots, copilots or automated workflows) "
           "for security, controls and safe operation? Tell us which systems and what concerns you."),
    ("F3", "Would you like a formal AI risk assessment - risk-tiering of your AI use cases, a risk register and recommended controls? "
           "Tell us which use cases or business areas matter most."),
    ("F4", "Would you like our help to develop an AI strategy and roadmap - identifying where AI can add value, prioritising use cases "
           "and planning delivery? Tell us your goals and timeline."),
]
PAGES.append(dict(no=n, section="How we can help",
                  desc="Optional - a few open questions so we can tailor our follow-up. A short answer such as 'Yes, for all staff' or 'Not at this time' is fine.",
                  items=[dict(tag=t, title=x, subtitle="", required=False, type="longtext", options=[]) for t, x in FOLLOW])); n += 1
PAGES.append(dict(no=n, section="Where should we send your report?",
                  desc="Last step. We use these details only to prepare and send your report and to respond to the requests you made on the previous page.",
                  items=[Q("P1"), Q("P6"), Q("P8")]))
# the final-page questions get shorter, clearer titles in the form (tag kept for mapping)
PAGES[-1]["items"][2]["subtitle"] = "Please use your company email address."
for p in PAGES[-1]["items"]: p["type"] = "text"

flat = [i for p in PAGES for i in p["items"]]
assert len(flat) == 34 + 4, len(flat)
spec = dict(title=FORM_TITLE, theme=dict(colour="#" + LAV, background="form_images/GT_Form_Background_2000x1200.png",
            banner="form_images/GT_Form_Cover_Banner_1600x500.png"), confirmation=CONFIRM, pages=PAGES)
json.dump(spec, open(os.path.join(OUT, "form_spec.json"), "w"), indent=1, ensure_ascii=False)
DESC = INTRO + "\n\n" + WHEN + "\n\n" + GET
print("description length (chars):", len(DESC))

# ------------------------------------------------------------------ 3. images
def font(sz, bold=False):
    f = "/usr/share/fonts/opentype/inter/Inter-%s.otf" % ("Bold" if bold else "Regular")
    return ImageFont.truetype(f, sz)

def hexrgb(h): return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

def gradient(w, h, stops):
    img = Image.new("RGB", (w, h)); px = img.load()
    for y in range(h):
        for x in range(w):
            t = (x / w * 0.55 + y / h * 0.45)
            for i in range(len(stops) - 1):
                if t <= stops[i + 1][0] or i == len(stops) - 2:
                    a, b = stops[i], stops[i + 1]; u = min(1, max(0, (t - a[0]) / (b[0] - a[0])))
                    px[x, y] = tuple(int(a[1][k] + (b[1][k] - a[1][k]) * u) for k in range(3)); break
    return img

def blob(img, bbox, rgb, alpha, blur):
    """Soft coloured ellipse composited without dark fringes (blur the mask, not the RGBA)."""
    m = Image.new("L", img.size, 0); ImageDraw.Draw(m).ellipse(bbox, fill=alpha)
    m = m.filter(ImageFilter.GaussianBlur(blur))
    img.paste(Image.new("RGBA", img.size, rgb + (255,)), (0, 0), m)

def network(img, seed, centre, spread, n, line_a, dot_a, maxd, xmin=0):
    import random
    rnd = random.Random(seed); ov = Image.new("RGBA", img.size, (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    pts = []
    while len(pts) < n:
        x = int(rnd.gauss(centre[0], spread[0])); y = int(rnd.gauss(centre[1], spread[1]))
        if xmin <= x < img.size[0] - 30 and 30 <= y < img.size[1] - 30: pts.append((x, y))
    for i, p in enumerate(pts):
        for q in pts[i + 1:]:
            if ((p[0] - q[0]) ** 2 + (p[1] - q[1]) ** 2) ** .5 < maxd: d.line([p, q], fill=(255, 255, 255, line_a), width=2)
    for p in pts:
        r = rnd.choice([4, 5, 6, 8]); d.ellipse([p[0] - r, p[1] - r, p[0] + r, p[1] + r], fill=(255, 255, 255, dot_a))
    img.alpha_composite(ov)

def make_bg(w, h):
    base = gradient(w, h, [(0, hexrgb("F4F0FD")), (0.5, hexrgb("DCD2F6")), (1, hexrgb("C4B5EE"))]).convert("RGBA")
    blob(base, [w * .62, -h * .35, w * 1.25, h * .6], (255, 255, 255), 90, w // 40)
    blob(base, [-w * .18, h * .55, w * .30, h * 1.25], (150, 120, 220), 60, w // 40)
    blob(base, [w * .35, h * .80, w * .80, h * 1.4], (255, 255, 255), 80, w // 40)
    network(base, 7, (w * .80, h * .26), (w * .10, h * .16), 40, 80, 150, w * .12, xmin=int(w * .55))
    return base.convert("RGB")

make_bg(2000, 1200).save(os.path.join(IMG, "GT_Form_Background_2000x1200.png"), optimize=True)

W, H = 1600, 500
b = gradient(W, H, [(0, hexrgb("5B3FA8")), (0.55, hexrgb("7C5CC4")), (1, hexrgb("B7A4E8"))]).convert("RGBA")
blob(b, [W * .70, -H * .5, W * 1.2, H * .9], (255, 255, 255), 50, 28)
blob(b, [W * .55, H * .45, W * .95, H * 1.5], (255, 255, 255), 38, 28)
network(b, 11, (W * .84, H * .50), (W * .07, H * .22), 34, 100, 190, 170, xmin=1030)
dr = ImageDraw.Draw(b)
dr.rounded_rectangle([96, 96, 176, 106], radius=5, fill=(255, 255, 255, 230))
dr.text((96, 132), "AI Readiness", font=font(84, True), fill="white")
dr.text((96, 232), "& Risk Assessment", font=font(84, True), fill="white")
dr.text((98, 362), "Grant Thornton Bahrain  |  AI Governance, Risk & Compliance", font=font(30), fill=(240, 234, 255, 255))
dr.text((98, 410), "About 20-25 minutes  -  Your report is sent within 5 working days", font=font(24), fill=(226, 217, 250, 255))
b.convert("RGB").save(os.path.join(IMG, "GT_Form_Cover_Banner_1600x500.png"), optimize=True)

# ------------------------------------------------------------------ 4. Quick Import Word document
qi = Document()
# Forms' importer reads the layout: one question per block, options on their own lines, no extra marks.
# Very wide page so no option wraps mid-line (a wrap after a hyphen shows up as "audit- ready" in Forms).
s = qi.sections[0]; s.page_width = Cm(50); s.page_height = Cm(29.7); s.left_margin = s.right_margin = Cm(2); s.top_margin = s.bottom_margin = Cm(1.8)
st = qi.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(11)
t = qi.add_paragraph(); r = t.add_run(FORM_TITLE); r.bold = True; r.font.size = Pt(20)
for blk in (INTRO, WHEN, GET):
    qi.add_paragraph()
    for ln in blk.split("\n"):
        if ln.strip(): qi.add_paragraph(ln)
qi.add_paragraph(); qi.add_paragraph(); qi.add_paragraph()
LET = "ABCDEFGHIJKL"
k = 0
for pg in PAGES:
    for it in pg["items"]:
        k += 1
        p = qi.add_paragraph(); p.paragraph_format.keep_with_next = bool(it["options"])
        title = f"{k}. {it['title']}" + (f" {it['subtitle']}" if it["subtitle"] and it["options"] else "")
        p.add_run(title)
        for i, o in enumerate(it["options"]):
            op = qi.add_paragraph(f"{LET[i]}. {o}"); op.paragraph_format.left_indent = Cm(0.8); op.paragraph_format.space_after = Pt(0)
        qi.add_paragraph(); qi.add_paragraph()
qi.save(os.path.join(OUT, "GT_AI_Readiness_Form_QuickImport.docx"))
print("quick import questions:", k)

# ------------------------------------------------------------------ 5. response converter workbook
THIN = Side(style="thin", color="CFC6EA"); BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HFILL = PatternFill("solid", fgColor=LAV_DARK); SOFT = PatternFill("solid", fgColor=LAV_PALE); YEL = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
st_ = wb.active; st_.title = "Start here"
ex = wb.create_sheet("Forms export")
ld = wb.create_sheet("Load response")
ex.sheet_properties.tabColor = LAV; ld.sheet_properties.tabColor = LAV_DARK; st_.sheet_properties.tabColor = LAV_MID

# Forms export sheet: expected header row (pre-filled; the real export can simply be pasted over A1)
sys_cols = ["ID", "Start time", "Completion time", "Email", "Name"]
heads = sys_cols + [i["title"] for i in flat]
for j, h in enumerate(heads, 1):
    c = ex.cell(1, j, h); c.font = Font(bold=True, color="FFFFFF", size=9); c.fill = HFILL; c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = BOX
    ex.column_dimensions[CL(j)].width = 18 if j <= 5 else 30
ex.row_dimensions[1].height = 64; ex.freeze_panes = "F2"
ex["A3"] = None

# Load response sheet
ld.sheet_view.showGridLines = False
ld.column_dimensions["A"].width = 2
for col, w in zip("BCDEFGH", (30, 8, 50, 46, 13, 4, 4)): ld.column_dimensions[col].width = w
ld.column_dimensions["G"].hidden = True; ld.column_dimensions["H"].hidden = True
ld["B1"] = "LOAD A RESPONSE"; ld["B1"].font = Font(bold=True, size=16, color=LAV_DARK)
ld["B2"] = "Choose which response from the 'Forms export' sheet to load, then copy each block into the engine (Paste > Values)."
ld["B2"].font = Font(italic=True, color="6B6B6B")
ld["B4"] = "Response number to load"; ld["B4"].font = Font(bold=True)
ld["C4"] = 1; ld["C4"].fill = YEL; ld["C4"].border = BOX; ld["C4"].alignment = Alignment(horizontal="center")
ld["D4"] = '="of "&MAX(0,COUNTA(\'Forms export\'!$A:$A)-1)&" responses pasted  (1 = first row under the header)"'
ld["B5"] = "Sheet row used"; ld["C5"] = "=C4+1"; ld["C5"].alignment = Alignment(horizontal="center")
ld["B6"] = "Check"; ld["B6"].font = Font(bold=True)

TITLE = {i["tag"]: i["title"] for i in flat}
def pick(tag):   # raw value for a question from the loaded export row: header starts with the question wording (or carries its [tag])
    pre = TITLE[tag][:60].replace('"', '""')
    col = (f"IFERROR(MATCH(\"{pre}*\",'Forms export'!$1:$1,0),IFERROR(MATCH(\"*] {pre}*\",'Forms export'!$1:$1,0),"
           f"MATCH(\"*[{tag}]*\",'Forms export'!$1:$1,0)))")
    return (f"=IFERROR(IF(INDEX('Forms export'!$A:$ZZ,$C$5,{col})=\"\",\"\",INDEX('Forms export'!$A:$ZZ,$C$5,{col})),\"\")")

def strip_prefix(ref):  # drop a leading option marker ("A. " / "12. ") if Quick import kept it
    f = f'FIND(". ",{ref})'
    return (f'IFERROR(IF(AND({f}<=3,OR(ISNUMBER(--LEFT({ref},{f}-1)),AND(LEN(LEFT({ref},{f}-1))=1,'
            f'CODE(UPPER(LEFT({ref},1)))>=65,CODE(UPPER(LEFT({ref},1)))<=69))),MID({ref},{f}+2,2000),{ref}),{ref})')

# hidden Options sheet: the client-facing option text per tag (used to expand a bare letter and to validate the text)
op = wb.create_sheet("Options"); op.sheet_state = "hidden"; OPTCOL = {}
for j, (tg, q) in enumerate([(t, q) for t, q in questions.items() if q["options"]], 1):
    op.cell(1, j, tg)
    for i, o in enumerate(q["options"], 2): op.cell(i, j, o)
    OPTCOL[tg] = (CL(j), len(q["options"]) + 1)

row = 8
BLOCKS = [
    ("BLOCK 1  ->  engine sheet '1. Profile', cells C4:C13",
     [("P1", "Organisation name"), ("P2", "Sector"), ("P3", "Organisation size"), ("P4", "Scope of assessment"), ("P5", "Scope description (optional)"),
      ("P6", "Completed by"), ("P7", "Role / title"), ("P8", "Email address"), ("DATE", "Assessment date (dd/mm/yyyy)"), ("GT", "Grant Thornton contact (optional)")], 4),
    ("BLOCK 2  ->  engine sheet '1. Profile', cells C16:C21",
     [("P9", "CBB licensed?"), ("P10", "Personal data of individuals in Bahrain?"), ("P11", "Offered / used in the EU?"),
      ("P12", "Operates in KSA, UAE or other GCC?"), ("P13", "Government entity / public body?"), ("P14", "AI provider or deployer?")], 16),
    ("BLOCK 3  ->  engine sheet '2. Assessment', cells E10:E13",
     [(f"A{i}", questions[f"A{i}"]["title"][:52] + ("..." if len(questions[f'A{i}']['title']) > 52 else "")) for i in range(1, 5)], 10),
    ("BLOCK 4  ->  engine sheet '2. Assessment', cells E15:E30",
     [(f"Q{i}", questions[f"Q{i}"]["title"][:52] + ("..." if len(questions[f'Q{i}']['title']) > 52 else "")) for i in range(5, 21)], 15),
]
first_val_rows = []
check_cells = []
for title, items, _cell in BLOCKS:
    ld.cell(row, 2, title).font = Font(bold=True, color="FFFFFF");
    for c in range(2, 7): ld.cell(row, c).fill = HFILL
    row += 1
    for c, h in zip(range(2, 7), ["Field", "Tag", "Value to paste  (select the yellow range, copy, then Paste > Values)", "As received from Forms", "Found"]):
        x = ld.cell(row, c, h); x.font = Font(bold=True, size=9, color=LAV_DARK); x.fill = PatternFill("solid", fgColor=LAV_SOFT); x.border = BOX
    row += 1; start = row
    for tag, label in items:
        ld.cell(row, 2, label).alignment = Alignment(wrap_text=True, vertical="top")
        ld.cell(row, 3, tag if tag not in ("DATE", "GT") else "-").alignment = Alignment(horizontal="center", vertical="top")
        if tag == "DATE":
            ld.cell(row, 5, "=IFERROR(INT(INDEX('Forms export'!$A:$ZZ,$C$5,MATCH(\"Completion time\",'Forms export'!$1:$1,0))),\"\")")
            ld.cell(row, 4, f"=E{row}"); ld.cell(row, 4).number_format = "dd/mm/yyyy"; ld.cell(row, 5).number_format = "dd/mm/yyyy"
            ld.cell(row, 6, f'=IF(E{row}="","MISSING","ok")')
        elif tag == "GT":
            ld.cell(row, 4, None); ld.cell(row, 5, "(leave blank - GT fills in)"); ld.cell(row, 6, "-")
        elif tag in OPTCOL:
            col_, last = OPTCOL[tag]; rng = f"Options!${col_}$2:${col_}${last}"
            ld.cell(row, 5, pick(tag)); ld.cell(row, 7, "=" + strip_prefix(f"E{row}"))
            # match ignoring spaces / case (Quick import can insert stray spaces, e.g. "audit- ready"), then output the canonical option text
            ld.cell(row, 8, f'=IF(G{row}="","",IFERROR(MATCH(1,INDEX(--(SUBSTITUTE({rng}," ","")=SUBSTITUTE(G{row}," ","")),0),0),""))')
            ld.cell(row, 4, f'=IF(E{row}="","",IF(AND(LEN(G{row})=1,CODE(UPPER(G{row}))>=65,CODE(UPPER(G{row}))<=69),'
                            f'IFERROR(INDEX({rng},CODE(UPPER(G{row}))-64),G{row}),IF(H{row}<>"",INDEX({rng},H{row}),G{row})))')
            ld.cell(row, 6, f'=IF(E{row}="","MISSING",IF(SUMPRODUCT(--({rng}=D{row}))=0,"CHECK TEXT","ok"))')
        else:
            ld.cell(row, 5, pick(tag)); ld.cell(row, 4, f"=E{row}")
            opt = tag in ("P5", "P7")
            ld.cell(row, 6, f'=IF(E{row}="","{"-" if tag == "P5" else "MISSING"}","ok")')
        for c in range(2, 7):
            ld.cell(row, c).border = BOX; ld.cell(row, c).alignment = Alignment(wrap_text=True, vertical="top", horizontal=ld.cell(row, c).alignment.horizontal)
        ld.cell(row, 4).fill = YEL; ld.cell(row, 5).font = Font(color="6B6B6B", size=9)
        ld.row_dimensions[row].height = 30 if tag not in ("P5", "GT", "DATE") else 18
        row += 1
    first_val_rows.append((start, row - 1)); check_cells.append(f"F{start}:F{row - 1}")
    row += 1
# overall check
ld["C6"] = '=IF(COUNTIF(F8:F200,"MISSING")+COUNTIF(F8:F200,"CHECK TEXT")=0,"All answers found and recognised","Attention: "&(COUNTIF(F8:F200,"MISSING")+COUNTIF(F8:F200,"CHECK TEXT"))&" field(s) - see the Found column")'
ld["C6"].font = Font(bold=True, color=LAV_DARK)
ld.merge_cells("C6:E6")

# follow-up interests + contact block
ld.cell(row, 2, "FOLLOW-UP (for GT sales / delivery - not for the engine)").font = Font(bold=True, color="FFFFFF")
for c in range(2, 7): ld.cell(row, c).fill = HFILL
row += 1
for tag, label in [(t, l) for t, l in zip(["F1", "F2", "F3", "F4"], ["AI training", "Review of AI / automation systems", "AI risk assessment", "AI strategy & roadmap"])]:
    ld.cell(row, 2, label); ld.cell(row, 3, tag).alignment = Alignment(horizontal="center", vertical="top")
    ld.cell(row, 4, pick(tag))
    ld.merge_cells(start_row=row, start_column=4, end_row=row, end_column=5)
    ld.cell(row, 4).alignment = Alignment(wrap_text=True, vertical="top"); ld.row_dimensions[row].height = 48
    for c in range(2, 7): ld.cell(row, c).border = BOX
    row += 1
ld.freeze_panes = "A8"

# Start here
st_.sheet_view.showGridLines = False
st_.column_dimensions["A"].width = 2; st_.column_dimensions["B"].width = 6; st_.column_dimensions["C"].width = 110
st_["B1"] = "GT AI Readiness - Response Converter"; st_["B1"].font = Font(bold=True, size=18, color=LAV_DARK)
st_["B2"] = "GT internal. Turns a Microsoft Forms response into engine-ready values. It contains no scoring logic."; st_["B2"].font = Font(italic=True, color="6B6B6B")
steps = [
    "In Microsoft Forms open the form > Responses > 'Open results in Excel' (or download). Select the whole table INCLUDING the header row.",
    "Paste it into the 'Forms export' sheet at cell A1 (it overwrites the example header row). Column order does not matter - answers are found by matching each column header to the question wording.",
    "Go to 'Load response', set the response number (1 = the first response under the header). Check that it says 'All required answers found'.",
    "Open GT_AI_Readiness_Assessment_Tool.xlsx (the engine). Save a copy first: <Organisation>_Engine.xlsx.",
    "For each of the four yellow blocks: select the yellow range, Copy, go to the engine cell named in the block title, Paste Special > Values. The engine then scores the answers and the 3. Report sheet fills in.",
    "Fill the Grant Thornton contact (engine '1. Profile' C13) if needed, then export '3. Report' to PDF and email it. Do NOT send the engine or this converter to the client.",
    "Faster alternative for many responses: run src/generate_readiness_reports.py on the Forms export - it does steps 3-6 automatically and also writes the PDF and an Outlook draft.",
]
for i, s_ in enumerate(steps, 1):
    st_.cell(4 + i, 2, i).font = Font(bold=True, color="FFFFFF"); st_.cell(4 + i, 2).fill = PatternFill("solid", fgColor=LAV); st_.cell(4 + i, 2).alignment = Alignment(horizontal="center", vertical="top")
    st_.cell(4 + i, 3, s_).alignment = Alignment(wrap_text=True, vertical="top"); st_.row_dimensions[4 + i].height = 46
st_["B13"] = "Paste Values matters: it keeps the engine's drop-down validation and formatting intact."; st_["B13"].font = Font(italic=True, color="6B6B6B")
wb.calculation.fullCalcOnLoad = True
wb.save(os.path.join(OUT, "GT_AI_Readiness_Form_Response_Converter.xlsx"))
json.dump(dict(blocks=[(t, a, b) for (t, _i, _c), (a, b) in zip(BLOCKS, first_val_rows)]), open(os.path.join(OUT, "..", "..", "..", "scratch_blocks.json"), "w")) if False else None
print("converter value rows:", first_val_rows)

# ------------------------------------------------------------------ 6. build guide (Word)
g = Document()
s = g.sections[0]; s.page_width = Cm(21); s.page_height = Cm(29.7); s.left_margin = s.right_margin = Cm(2); s.top_margin = Cm(1.6); s.bottom_margin = Cm(1.6)
st = g.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(10.5); st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
DK = RGBColor.from_string(LAV_DARK)
for lvl, sz in ((1, 15), (2, 12)):
    h = g.styles[f"Heading {lvl}"]; h.font.name = "Calibri"; h.font.size = Pt(sz); h.font.color.rgb = DK
    rpr = h.element.get_or_add_rPr(); rf = rpr.find(qn("w:rFonts"))
    if rf is None: rf = OxmlElement("w:rFonts"); rpr.append(rf)
    for a_ in ("w:ascii", "w:hAnsi", "w:cs"): rf.set(qn(a_), "Calibri")

def shade(cell, hexcol):
    tcPr = cell._tc.get_or_add_tcPr(); sh = OxmlElement("w:shd"); sh.set(qn("w:val"), "clear"); sh.set(qn("w:fill"), hexcol); tcPr.append(sh)

def box(text, fill=LAV_PALE, mono=False, size=9.5):
    t = g.add_table(rows=1, cols=1); c = t.rows[0].cells[0]; shade(c, fill); c.text = ""
    for i, ln in enumerate(text.split("\n")):
        p = c.paragraphs[0] if i == 0 else c.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(ln); r.font.size = Pt(size)
        if mono: r.font.name = "Consolas"
    g.add_paragraph().paragraph_format.space_after = Pt(2)

def bullet(txt, bold_lead=None):
    p = g.add_paragraph(style="List Bullet")
    if bold_lead: p.add_run(bold_lead).bold = True
    p.add_run(txt)

def step(n_, lead, txt):
    p = g.add_paragraph(); p.paragraph_format.left_indent = Cm(0.9); p.paragraph_format.first_line_indent = Cm(-0.9); p.paragraph_format.space_after = Pt(3)
    r = p.add_run(f"{n_}.  "); r.bold = True; r.font.color.rgb = RGBColor.from_string(LAV)
    p.add_run(lead + " ").bold = True; p.add_run(txt)

hero = g.add_table(rows=1, cols=1); hc = hero.rows[0].cells[0]; shade(hc, "5B3FA8"); hc.text = ""
r = hc.paragraphs[0].add_run("AI Readiness & Risk Assessment - Microsoft Form build guide"); r.bold = True; r.font.size = Pt(19); r.font.color.rgb = RGBColor(255, 255, 255)
r = hc.add_paragraph().add_run("GT internal  |  Set up the client form in about 15 minutes, collect answers in Excel, turn each response into a report"); r.font.size = Pt(10.5); r.font.color.rgb = RGBColor(0xE8, 0xE0, 0xF8)
g.add_paragraph()

g.add_heading("1. What this pack is", 1)
g.add_paragraph("The Microsoft 365 connector in this workspace can read and write OneDrive/SharePoint files and Outlook, but it cannot create or edit Microsoft Forms "
                "(Microsoft Forms has no public create-form API). So the form itself is built by you in Forms, from this pack, in about 15 minutes. "
                "Everything that takes time is already prepared: the question text and options (identical to the questionnaire), the page layout, the lavender theme and "
                "cover images, the intro text, the four open-ended questions and a converter that turns a Forms response into engine-ready values.")
bullet(" 38 questions as numbered items with lettered options, in form order, for Forms > Quick import.", "GT_AI_Readiness_Form_QuickImport.docx -")
bullet(" lavender background (2000 x 1200) and cover banner (1600 x 500).", "form_images/ -")
bullet(" paste a Forms export, get the values to paste into the engine. No scoring logic, safe to keep open next to the engine.", "GT_AI_Readiness_Form_Response_Converter.xlsx -")
bullet(" the complete form definition (pages, questions, options, required flags) as data.", "form_spec.json -")
g.add_paragraph("The engine (GT_AI_Readiness_Assessment_Tool.xlsx) is never shared: clients only see the form and later receive the PDF report.")

g.add_heading("2. Build the form", 1)
step(1, "Create.", "Go to forms.office.com > New Form. Name it 'AI Readiness & Risk Assessment'.")
step(2, "Fast path - Quick import.", "On the Forms home page choose Quick import and upload GT_AI_Readiness_Form_QuickImport.docx. Microsoft's documentation says Quick import converts titles, subtitles, multiple-choice and open-text questions from a Word or PDF file; "
     "check the result (open-text questions are plain numbered lines in the file and should import as text questions), make sure each question is single-answer (not multi-select), and that the question wording is unchanged (the converter finds answers by the wording). If Quick import is not offered on your account, add the questions manually - the layout table in section 3 and form_spec.json give every title, subtitle and option.")
step(3, "Title and introduction.", "Paste the form title and description from section 4. Insert the cover banner image above the title (the image icon in the title area); if your Forms version does not offer that, add it to the first question with Insert media. "
     "The Quick import file already carries the intro under the title, so check that the description box holds it (WHEN TO USE THIS TOOL and WHAT YOU GET); if it is cut off, paste it from section 4.")
step(4, "Sections (pages).", "Add one section per row of the layout table in section 3 (Add new > Section) and drag the questions into them. Mark every assessment question Required; leave the scope description and the four open-ended questions optional.")
step(5, "Theme.", f"Style (paint-brush icon) > custom colour, enter #{LAV}, then upload form_images/GT_Form_Background_2000x1200.png as the background image. Forms crops background images to fit, so the artwork keeps its key detail away from the edges.")
step(6, "Settings.", "... > Settings: Who can fill out this form = Anyone can respond; untick 'Record name'; do not tick 'One response per person' for anonymous links; add the thank-you message from section 4 under 'Customise thank you message'; leave 'Shuffle options' off. Preview on desktop and phone, submit one test response, then delete it.")
step(7, "Responses in Excel.", "Responses tab > Open results in Excel. Forms keeps one row per submission. Keep that workbook private to GT - it contains the clients' answers and the form owner's OneDrive holds it; never share it with clients.")
step(8, "Share.", "Share > Anyone can respond > copy the link (or QR code) and send it to the client.")

g.add_heading("3. Page-by-page layout", 1)
tb = g.add_table(rows=1, cols=4); tb.style = "Table Grid"
for c, h_ in zip(tb.rows[0].cells, ["Page", "Section title", "Questions (tag)", "Notes"]):
    c.text = ""; rr = c.paragraphs[0].add_run(h_); rr.bold = True; rr.font.size = Pt(9.5); rr.font.color.rgb = RGBColor(255, 255, 255); shade(c, LAV_DARK)
notes = {1: "Form header with the intro (WHEN TO USE, WHAT YOU GET, 5-day report promise), then the 11 profile questions. Name, organisation and email are collected on the last page instead.",
         2: "AI use and inherent-risk questions (A1-A4).", 3: "Governance structure, policy, guardrails, incidents.",
         4: "Risk framework, appetite, inventory, third parties, securing AI, AI-enabled threats.", 5: "AI literacy and skills.",
         6: "Strategy, use-case portfolio.", 7: "Data governance, architecture, model risk.",
         8: "Four optional open-ended questions: training, systems review, risk assessment, strategy and roadmap.",
         9: "Name, organisation, company email only, then Submit."}
for pg in PAGES:
    r_ = tb.add_row().cells
    tags = ", ".join(i["tag"] for i in pg["items"])
    for c, v in zip(r_, [str(pg["no"]), pg["section"] or "(form header)", tags, notes[pg["no"]]]):
        c.text = ""; rr = c.paragraphs[0].add_run(v); rr.font.size = Pt(9)
    shade(r_[0], LAV_PALE)
for row_ in tb.rows:
    for c, w in zip(row_.cells, (Cm(1.2), Cm(5.0), Cm(5.0), Cm(5.8))): c.width = w
g.add_paragraph()
g.add_paragraph("Order matters only for the respondent. The converter and the report generator find each answer by the question wording in the column header, so do not reword the questions (the engine ID in the table, e.g. Q5, is for reference and is not shown in the form).")

g.add_heading("4. Text to paste into the form", 1)
g.add_paragraph("Form title").runs[0].bold = True
box(FORM_TITLE)
g.add_paragraph("Form description (page 1) - about %d characters" % len(DESC)).runs[0].bold = True
box(DESC, mono=False)
g.add_paragraph("Section titles and descriptions").runs[0].bold = True
box("\n\n".join(f"{p['section']}\n{p['desc']}" for p in PAGES[1:]), size=9)
g.add_paragraph("Open-ended questions (page 9) - Long answer, optional").runs[0].bold = True
box("\n\n".join(x for t, x in FOLLOW), size=9)
g.add_paragraph("Last page - three short-answer questions, all required").runs[0].bold = True
box("Organisation name\nYour name\nYour email address (the report will be sent here)   - subtitle: Please use your company email address.", size=9)
g.add_paragraph("Thank-you message (Settings)").runs[0].bold = True
box(CONFIRM)

g.add_heading("5. From response to report", 1)
g.add_paragraph("Option A - paste into Excel (one response at a time)").runs[0].bold = True
step(1, "Export.", "Forms > Responses > Open results in Excel. Select the whole table including the header row.")
step(2, "Paste.", "Open GT_AI_Readiness_Form_Response_Converter.xlsx > 'Forms export' sheet > paste at A1. Go to 'Load response' and set the response number (1 = first response).")
step(3, "Check.", "The cell under the response number should read 'All answers found and recognised'. If a row says MISSING or CHECK TEXT, the question was skipped or its wording was edited in Forms - fix the answer or restore the original question wording.")
step(4, "Fill the engine.", "Save a copy of the engine as <Organisation>_Engine.xlsx. Copy each of the four yellow blocks and Paste Special > Values at the cell named in the block title: '1. Profile' C4:C13, '1. Profile' C16:C21, '2. Assessment' E10:E13, '2. Assessment' E15:E30. The date is the Forms completion date.")
step(5, "Report.", "The '3. Report' sheet fills in. Save it as PDF and email it from your own mailbox; the report is due within 5 working days of the submission. Never send the engine or the converter.")
g.add_paragraph("Option B - one command for one or many responses").runs[0].bold = True
box("python src/generate_readiness_reports.py  FormsExport.xlsx  GT_AI_Readiness_Assessment_Tool.xlsx  OUT_DIR", mono=True)
g.add_paragraph("It reads the same export directly and writes, per organisation, the PDF report, an Outlook email draft with the PDF attached and a filled copy of the engine. "
                "Both options were tested on a Forms-format export with the extra system columns and the follow-up questions: they give the same scores.")
g.add_paragraph("The four open-ended answers are ignored by the engine. Read them on the 'Load response' sheet (FOLLOW-UP block) or in the export to follow up with training, systems review, risk assessment and strategy proposals.")

g.add_heading("6. Good to know", 1)
for lead, txt in [
    ("Question wording. ", "Titles and options are copied from the questionnaire. If you edit an option in Forms the converter flags it (CHECK TEXT) and the generator rejects it, because the engine matches the exact text."),
    ("Letters in options. ", "The form shows plain option text. If Quick import keeps a leading 'A.' or '1.' the converter and the generator strip it automatically."),
    ("Character limits. ", "Community answers cite roughly 4,000 characters per question field and 1,000-1,700 for the form description; the description here is %d characters and the longest question block about 900. Check the preview for truncation." % len(DESC)),
    ("Email check. ", "Forms short-answer questions may not validate email format; the report email address is read from [P8], so look at it before sending."),
    ("Where the answers live. ", "In the Excel workbook linked to the form in the form owner's OneDrive. Move the export into the engagement folder; do not share the workbook with the client."),
    ("Changing a question. ", "Edit build_questionnaire.py / tool_data.py, regenerate the questionnaire, then rerun build_forms_pack.py so the form, converter and engine stay aligned."),
]: bullet(txt, lead)
for sec in g.sections:
    fp = sec.footer.paragraphs[0]; fp.text = "Grant Thornton Bahrain - AI Readiness Microsoft Form build guide - GT internal"; fp.runs[0].font.size = Pt(8)
g.save(os.path.join(OUT, "GT_AI_Readiness_Form_Build_Guide.docx"))
print("guide saved")
