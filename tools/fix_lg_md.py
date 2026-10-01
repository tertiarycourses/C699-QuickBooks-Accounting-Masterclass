"""C699: the LG Markdown mirror at the repo root, scrubbed like the DOCX."""
import re
SRC = "source-wsq/LG-Advanced Transactional Accounting with QuickBooks Online.md"
OUT = "LG-QuickBooks Accounting Masterclass (C699).md"
t = open(SRC).read()
t = t.replace("Advanced Transactional Accounting with QuickBooks Online", "QuickBooks Accounting Masterclass")
t = t.replace("**WSQ Course Code:** TGS-2023018989  \n", "**Course Code:** C699  \n")
t = re.sub(r"^\*\*TSC:\*\* .*\n", "", t, flags=re.M)
t = t.replace("**Version:** 11.1 · 16 August 2026", "**Version:** 1.0 · 1 October 2026")
t = t.replace("*A4 Decision Lens*", "*Decision Lens*").replace("*Impact Assessment*", "*Impact Review*")
t = re.sub(r"^\*\*Criteria:\*\* .*\n", "", t, flags=re.M)
t = t.replace("All K3/A4 impacts are addressed", "All standards impacts are addressed")
t, n = re.subn(r"## Assessment Preparation\n.*?(?=## Source Register)", "", t, flags=re.S)
assert n == 1
t = t.replace("2 days/16 hours; LO1-LO3; published topic and assessment scope. "
              "https://www.tertiarycourses.com.sg/wsq-advanced-transactional-accounting-with-quickbooks-online.html",
              "2 days/15 hours; LO1-LO3; published topics. "
              "https://www.tertiarycourses.com.sg/quickbooks-accounting-masterclass.html")
open(OUT, "w").write(t)
print("saved", OUT)
