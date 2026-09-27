"""Development-only distillation of the user's approved packing documents."""
from copy import deepcopy
from hashlib import sha256
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn

app = Path(__file__).resolve().parents[1]
reference = app.parent / '20260926' / '000 Output'
assets = app / 'src' / 'assets'
assets.mkdir(parents=True, exist_ok=True)
notes = ['# Packing document template contract', '',
         'The retained September 26 documents control the layout. Templates contain sanitized component examples.',
         'A4 portrait; margins top 11 mm, bottom 12 mm, sides 10 mm. Header 4 mm; footer 5 mm.',
         'One bordered table per order. Fixed columns 567, 6293, 3061, 850 twips.',
         'Header merged across 4 columns, Arial 9.5 pt; order ID bold; Number and Time tabs retained.',
         'Column headers Arial 8 pt bold. Product rows Arial 9 pt. Cell padding 55 twips vertical, 80 horizontal.',
         'Borders D9D9D9, 5 eighth-points. Orange F8CBAD / FFF0E3 / FFF7F1; blue BED9ED / EAF4FC / F5FAFE.',
         'Body title and explanatory paragraph preserved. Batch heading 10 pt bold with 8 pt source filename.',
         'Page footer label, date and PAGE field retained. Blank 1 pt spacer separates tables.',
         'Each row cannot split. All paragraphs through the penultimate row keep with next; final row releases.',
         'Slots: body title, intro, batch heading, order header, product rows, footer label and date.',
         'The prototype table is cloned for every order. All order data is removed from distributable templates.',
         'No drawings, content controls, tracked changes, comments or custom text-box slots are present.',
         'Preserve section properties, styles, theme, numbering, relationships, settings and page field.',
         'Intentional correction: every order retains complete Number and Time underlines.',
         'Check final Word pagination: A4, each table starts and ends on one page, expected order content.', '']
for group, asset, pages in [('SPX','orange.docx',9),('LAZADA_TIKTOK','blue.docx',1)]:
    source = reference / f'{group}_Packing_Checklist_2026-09-26_1.docx'
    d = Document(source)
    body = d._element.body
    originals = list(body)
    retained = [deepcopy(d.paragraphs[i]._p) for i in range(3)]
    retained += [deepcopy(d.tables[0]._tbl), deepcopy(d.paragraphs[3]._p)]
    for node in originals:
        if node.tag != qn('w:sectPr'):
            body.remove(node)
    for node in retained:
        body.insert(len(body)-1, node)
    replacements = ['Packing List', 'One table is one order. Write the number used on the waybill and the packing completion time in its header.', 'Batch']
    for p, text in zip(d.paragraphs[:3], replacements):
        if p.runs:
            p.runs[0].text = text
            for r in p.runs[1:]: r.text = ''
    t = d.tables[0]
    for r in t.cell(0,0).paragraphs[0].runs:
        if 'Order ' in r.text: r.text = 'Order ID: SAMPLE'
    for cell, value in zip(t.rows[2].cells, ['[ ]','Product','Variation','1']):
        for index, run in enumerate(cell.paragraphs[0].runs):
            run.text = value if index == len(cell.paragraphs[0].runs)-1 else ''
    for p in d.sections[0].footer.paragraphs:
        for r in p.runs:
            if 'Page' in r.text: r.text = 'Packing List | Page '
    d.core_properties.title = 'Packing List Template'
    d.core_properties.author = 'Packing List Manager'
    d.core_properties.last_modified_by = 'Packing List Manager'
    d.save(assets / asset)
    notes += [f'Reference: {source}', f'SHA256: {sha256(source.read_bytes()).hexdigest()}',
              f'Pages: {pages}; section count: 1; render: qa/reference/{group}/page-*.png',
              f'Template: src/assets/{asset}', '']
(app / 'qa' / 'artifact.md').write_text('\n'.join(notes), encoding='utf-8')

