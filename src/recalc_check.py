"""Recalculate an xlsx via LibreOffice headless and report formula errors. Usage: recalc_check.py file.xlsx [sheet!cell ...]"""
import sys, subprocess, os, shutil, openpyxl, tempfile
src = os.path.abspath(sys.argv[1])
tmp = tempfile.mkdtemp(dir="/home/user/work")
shutil.copy(src, os.path.join(tmp, "in.xlsx"))
subprocess.run(["soffice", "-env:UserInstallation=file:///home/user/work/lo_profile", "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", os.path.join(tmp, "out"), os.path.join(tmp, "in.xlsx")], capture_output=True, timeout=300)
wb = openpyxl.load_workbook(os.path.join(tmp, "out", "in.xlsx"), data_only=True)
errs = 0
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith(("#REF", "#VALUE", "#NAME", "#DIV", "#N/A", "Err:", "#NUM")):
                errs += 1
                if errs < 30: print("ERR", ws.title, c.coordinate, c.value)
print("errors:", errs)
for spec in sys.argv[2:]:
    sh, cell = spec.rsplit("!", 1)
    print(spec, "=>", repr(wb[sh][cell].value))
print("RECALC_PATH", os.path.join(tmp, "out", "in.xlsx"))
