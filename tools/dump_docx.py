import sys
from docx import Document
d = Document(sys.argv[1])
body = d.element.body
from docx.table import Table
from docx.text.paragraph import Paragraph
ti = 0
for el in body.iterchildren():
    tag = el.tag.split('}')[1]
    if tag == 'p':
        p = Paragraph(el, d)
        if p.text.strip(): print(f'P[{p.style.name}] {p.text}')
    elif tag == 'tbl':
        t = Table(el, d); print(f'--- TABLE {ti}')
        for r in t.rows: print('   | ' + ' | '.join(c.text.replace('\n', ' / ') for c in r.cells))
        ti += 1
for s in d.sections:
    for part in (s.header, s.footer):
        for p in part.paragraphs:
            if p.text.strip(): print('HF:', p.text)
