"""End-to-end v1.11.2 checks using copies of current PDFs and saved definitions."""
import copy
import importlib.util
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

import pdfplumber
from pypdf import PdfReader, PdfWriter
from docx import Document

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP / 'src'))
from extractors import parse_pdf, Batch, Order, Item, clean
from names import load_catalog, apply_names, lookup_key, source_pairs
from documents import build_document, verify_word, COMMENT_LABEL, item_prices

release = Path(sys.argv[1]).resolve()
source_root = Path(sys.argv[2]).resolve()
work = Path(tempfile.mkdtemp(prefix='v1112-smoke-', dir=APP / 'qa'))
folder = work / '20261005'
catalog_path = work / 'ProductNames.json'
shutil.copy2(APP.parent / 'Packing List Manager/ProductNames.json', catalog_path)
original_catalog = load_catalog(catalog_path)
selections = [('SPX/SPX PACKING LIST 1.pdf', [0, 10]),
              ('FLASH/FLASH PACKING LIST 1.pdf', [0]),
              ('LAZADA/LAZADA PACKING LIST 1.pdf', [0, 1, 2, 3])]
for relative, pages in selections:
    source = source_root / '20261005' / relative
    reader = PdfReader(source); writer = PdfWriter()
    for number in pages: writer.add_page(reader.pages[number])
    target = folder / relative; target.parent.mkdir(parents=True, exist_ok=True)
    with target.open('wb') as stream: writer.write(stream)
target = folder / 'TIKTOK/TIKTOK PACKING LIST.pdf'; target.parent.mkdir(parents=True)
shutil.copy2(APP / 'qa/names-regression-fixtures/20260926/LAZADA & TIKTOK/20260926_Tiktok_PL_1.pdf', target)


def call(action, **fields):
    request, response = work / 'request.json', work / 'response.json'
    request.write_text(json.dumps(dict(action=action, folder=str(folder), catalog_path=str(catalog_path), **fields)), encoding='utf-8')
    subprocess.run([str(release / 'App Files/PackingEngine.exe'), '--request', str(request), '--response', str(response)],
                   check=True, timeout=240, creationflags=subprocess.CREATE_NO_WINDOW)
    result = json.loads(response.read_text(encoding='utf-8'))
    assert result['ok'], result
    return result


scan = call('scan')
assert len(scan['entries']) == 4, scan
assert all(e['status'] in ('READY', 'DUPLICATE') for e in scan['entries']), scan
catalog = load_catalog(catalog_path)
sources = [parse_pdf(path) for path in folder.rglob('*.pdf') if '000 Output' not in path.parts]
legacy_path = APP.parent / 'Packing List Manager/App Files/Source/extractors.py'
spec = importlib.util.spec_from_file_location('previous_extractors', legacy_path)
legacy = importlib.util.module_from_spec(spec); sys.modules[spec.name] = legacy; spec.loader.exec_module(legacy)
for batch in sources:
    previous = legacy.parse_pdf(batch.path)
    assert {o.key: o.fingerprint() for o in previous.orders} == {o.key: o.fingerprint() for o in batch.orders}
printed = apply_names(sources, catalog)
expected = {order.order_id: order for batch in printed for order in batch.orders}
raw = {order.order_id: order for batch in sources for order in batch.orders}
assert any(o.buyer_comment for o in raw.values())
assert any(not o.buyer_comment for o in raw.values())
oxy = [(i, printed_item) for batch, named in zip(sources, printed) for order, named_order in zip(batch.orders, named.orders)
       for i, printed_item in zip(order.items, named_order.items) if "OXYHERBS',500ML" == ''.join(i.variation.upper().split())]
assert oxy and all(p.product == 'OXYHERBS SCENT' and p.variation == '500 ML' for _, p in oxy), oxy
new_rows = [r for r in catalog['rows'] if r['id'] not in {old['id'] for old in original_catalog['rows']}]
assert all(':' not in r['original_variation'] for r in new_rows), new_rows

result = call('generate', selected=[dict(path=e['path'], sha256=e['sha256']) for e in scan['entries']])
if result.get('confirmation_required'):
    result = call('generate', selected=[dict(path=e['path'], sha256=e['sha256']) for e in scan['entries']],
                  repeat_confirmation=result['repeat_token'], repeat_action='skip')
assert len(result['outputs']) == 3, result
seen = set(); comments = 0
for output in result['outputs']:
    with pdfplumber.open(output['path']) as pdf:
        for page in pdf.pages:
            assert abs(page.width - 595.3) < 1 and abs(page.height - 841.9) < 1
            for table in page.extract_tables():
                if not table or not table[0] or not (table[0][0] or '').startswith(('Order ID:', 'Order Number:')): continue
                identifier = table[0][0].split(':', 1)[1].split()[0]
                assert identifier not in seen; seen.add(identifier)
                order = expected[identifier]
                rows = [[clean(cell) for cell in row if clean(cell)] for row in table[2:]]
                priced = order.platform in ('SPX', 'FLASH')
                if priced and order.buyer_comment:
                    comment_row = rows.pop()
                    assert len(comment_row) == 2
                    # Some Word emoji glyphs have no extractable text. Verify the
                    # complete string below in the immutable printed snapshot.
                    assert comment_row[0] == COMMENT_LABEL
                    assert comment_row[1] in order.buyer_comment or order.buyer_comment in comment_row[1]
                    comments += 1
                wanted = [['[ ]', clean(i.product), clean(i.variation), str(i.pieces)] + (item_prices(i) if priced else []) for i in order.items]
                assert rows == wanted, (identifier, rows, wanted)
assert seen == set(expected), (seen, set(expected))
assert comments == sum(bool(o.buyer_comment) for o in expected.values() if o.platform in ('SPX', 'FLASH'))
history = json.loads((folder / '000 Output/.PackingData/Processing_History.json').read_text(encoding='utf-8'))
for order in raw.values():
    assert history['orders'][order.key]['order']['buyer_comment'] == order.buyer_comment
    assert history['orders'][order.key]['fingerprint'] == order.fingerprint()
for output in history['outputs'].values():
    for order in output['display_order_snapshots'].values():
        assert order['buyer_comment'] == expected[order['order_id']].buyer_comment
sales = call('sales_export')
book = work / 'Names.xlsx'; call('names_export', path=str(book))
call('names_import', path=str(book), rows=load_catalog(catalog_path)['rows'])

# Check a comment near a page boundary and a wrapped long comment in real Word.
stress = work / 'Comment pagination.docx'
orders = [Order('SPX', 'STRESS' + str(i), [Item('PRODUCT', '100 ML', 1, '45.00', '45.00')],
                'Please send an invoice and pack the bottles carefully. ' * (12 if i == 7 else 1) if i % 2 else '') for i in range(14)]
build_document([Batch('comments.pdf', 'SPX', 'hash', orders)], 'SPX', '20261005', stress)
layout = verify_word([stress], export_pdf=True)
assert layout[0]['orders'] == 14
report = dict(ok=True, version='1.11.2', work=str(work), orders=len(seen), comment_rows=comments,
              oxyherbs_correct=True, source_fingerprints_preserved=True, names_excel_roundtrip=True,
              packing_pdfs=[o['path'] for o in result['outputs']], sales=sales['path'],
              comment_pagination_pdf=str(stress.with_suffix('.pdf')), layout=layout)
(APP / 'qa/v1112-verification.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
print(json.dumps(report, indent=2))
