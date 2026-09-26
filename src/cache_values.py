# -*- coding: utf-8 -*-
"""Write calculated (cached) values into every formula cell of an openpyxl-generated workbook.

openpyxl writes formulas with an empty <v></v>. Excel recalculates them only when the workbook is
fully opened for editing; in Protected View (files downloaded from email / the web) and in some print
paths the cells therefore have no value. This script recalculates the workbook with LibreOffice and
patches the results into the original file's XML, leaving all formatting, charts and formulas intact.
Usage: cache_values.py file.xlsx [file2.xlsx ...]"""
import sys, os, re, shutil, subprocess, tempfile, zipfile, datetime
from xml.sax.saxutils import escape
import openpyxl

LO_PROFILE = "file:///home/user/work/lo_profile"


def recalc(src):
    tmp = tempfile.mkdtemp(dir="/home/user/work")
    shutil.copy(src, os.path.join(tmp, "in.xlsx"))
    subprocess.run(["soffice", f"-env:UserInstallation={LO_PROFILE}", "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                    "--outdir", os.path.join(tmp, "out"), os.path.join(tmp, "in.xlsx")], capture_output=True, timeout=600)
    return openpyxl.load_workbook(os.path.join(tmp, "out", "in.xlsx"), data_only=True)


def sheet_paths(z):
    wbx = z.read("xl/workbook.xml").decode("utf8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    rid2target = dict(re.findall(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rid2target.update({a: b for b, a in re.findall(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    out = {}
    for name, rid in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
        t = rid2target[rid].lstrip("/")
        out[name.replace("&amp;", "&")] = t if t.startswith("xl/") else "xl/" + t
    return out


CELL = re.compile(r'<c r="([A-Z]+[0-9]+)"([^>]*)><f>(.*?)</f><v\s*/?>(?:</v>)?</c>', re.S)


def patch(path):
    vals = recalc(path)
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
        paths = sheet_paths(z)
    filled = 0
    for sname, spath in paths.items():
        ws = vals[sname]
        xml = items[spath].decode("utf8")

        def rep(m):
            nonlocal filled
            ref, attrs, f = m.group(1), m.group(2), m.group(3)
            attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
            v = ws[ref].value
            if v is None:
                return f'<c r="{ref}"{attrs} t="str"><f>{f}</f><v></v></c>'
            filled += 1
            if isinstance(v, bool):
                return f'<c r="{ref}"{attrs} t="b"><f>{f}</f><v>{int(v)}</v></c>'
            if isinstance(v, (int, float)):
                return f'<c r="{ref}"{attrs}><f>{f}</f><v>{repr(float(v)) if isinstance(v, float) else v}</v></c>'
            if isinstance(v, (datetime.datetime, datetime.date)):
                base = datetime.datetime(1899, 12, 30)
                d = (v if isinstance(v, datetime.datetime) else datetime.datetime(v.year, v.month, v.day)) - base
                return f'<c r="{ref}"{attrs}><f>{f}</f><v>{d.days + d.seconds / 86400}</v></c>'
            s = str(v)
            if s.startswith("#") and s.endswith(("!", "?", "A")):
                return f'<c r="{ref}"{attrs} t="e"><f>{f}</f><v>{escape(s)}</v></c>'
            return f'<c r="{ref}"{attrs} t="str"><f>{f}</f><v>{escape(s)}</v></c>'

        items[spath] = CELL.sub(rep, xml).encode("utf8")
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in items.items():
            z.writestr(n, data)
    os.replace(tmp, path)
    return filled


if __name__ == "__main__":
    for p in sys.argv[1:]:
        print(p, "cached values written:", patch(p))
