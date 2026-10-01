"""C699 activities: copy the WSQ activities/ tree and strip the WSQ/TSC layer (README, .md, .xlsx)."""
import pathlib, re, shutil, zipfile

SRC, OUT = pathlib.Path("source-wsq/activities"), pathlib.Path("activities")
shutil.rmtree(OUT, ignore_errors=True)
shutil.copytree(SRC, OUT)

readme = OUT / "README.md"
t = readme.read_text()
t = t.replace("Advanced Transactional Accounting with QuickBooks Online · TGS-2023018989", "QuickBooks Accounting Masterclass · C699")
t = re.sub(r" · [KA]\d+(?:, [KA]\d+)*$", "", t, flags=re.M)
readme.write_text(t)

for md in OUT.glob("*/*.md"):
    t = md.read_text()
    t = re.sub(r"^\*\*TSC mapping:\*\* .*\n", "", t, flags=re.M)
    t = t.replace("All K3/A4 impacts are addressed", "All standards impacts are addressed")
    md.write_text(t)

SS = [(r"(<t[^>]*>)Criteria</t>", r"\1Course</t>"),
      (r"(<t[^>]*>)K\d(?:, A\d)+</t>", r"\1QuickBooks Accounting Masterclass (C699)</t>"),
      (r"All K3/A4 impacts are addressed", "All standards impacts are addressed")]
for x in OUT.glob("*/*.xlsx"):
    zin = zipfile.ZipFile(x)
    items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    zin.close()
    with zipfile.ZipFile(x, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            if info.filename == "xl/sharedStrings.xml":
                s = data.decode()
                for a, b in SS:
                    s = re.sub(a, b, s)
                data = s.encode()
            zout.writestr(info, data)
print("activities rewritten")
