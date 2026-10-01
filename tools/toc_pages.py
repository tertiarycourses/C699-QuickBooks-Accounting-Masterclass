"""Report TOC entries whose page number disagrees with the rendered PDF; --fix rewrites them."""
import re, subprocess, sys, docx
from docxutil import set_text
DOCX, PDF = sys.argv[1], sys.argv[2]
pages = subprocess.run(["pdftotext", "-layout", PDF, "-"], capture_output=True, text=True).stdout.split("\f")
d = docx.Document(DOCX)
toc = [p for p in d.paragraphs if re.match(r".+\t\d+$", p.text)]
bad = 0
cursor = 3
for p in toc:
    title, num = p.text.rsplit("\t", 1)
    key = re.sub(r"\s+", " ", title.strip())
    found = None
    for i, pg in enumerate(pages):
        if i < cursor: continue                 # skip cover/TOC pages and pages before the previous entry
        if key in re.sub(r"\s+", " ", pg):
            found = i + 1; cursor = i; break
    if found and str(found) != num:
        bad += 1; print(f"{title!r}: {num} -> {found}")
        if "--fix" in sys.argv:
            # rewrite only the run holding the page number
            runs = [r for r in p.runs if r.text]
            if runs and runs[-1].text.strip() == num:
                runs[-1].text = runs[-1].text.replace(num, str(found))
            else:
                set_text(p, f"{title}\t{found}")
    elif not found:
        print(f"{title!r}: NOT FOUND")
if "--fix" in sys.argv: d.save(DOCX)
print("mismatches:", bad)
