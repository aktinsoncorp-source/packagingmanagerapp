"""Verify full source names, per-file totals and their real A4 layout."""
from collections import Counter, defaultdict
import json
from pathlib import Path
import sys

APP = Path(__file__).resolve().parents[1]
ROOT = APP.parent
sys.path.insert(0, str(APP / 'src'))
from extractors import parse_pdf
from documents import build_document, verify_word
from docx import Document

work = APP / 'qa/original-names-layout'
work.mkdir(exist_ok=True)
groups = defaultdict(list)
for path in sorted((ROOT / '20260926').rglob('*.pdf')):
    batch = parse_pdf(path)
    groups[batch.group].append(batch)
paths, counts = [], {}
for group, batches in groups.items():
    path = work / f'{group}.docx'
    build_document(batches, group, '20260926', path)
    expected = Counter()
    for batch in batches:
        for order in batch.orders:
            for item in order.items:
                expected[(item.product, item.variation or '—')] += item.pieces
    table = Document(path).tables[-1]
    actual = {(r.cells[1].text, r.cells[2].text): int(r.cells[3].text) for r in table.rows[1:-1]}
    assert actual == dict(expected), group
    assert int(table.rows[-1].cells[3].text) == sum(expected.values())
    paths.append(path)
    counts[group] = {'rows': len(expected), 'pieces': sum(expected.values())}
result = verify_word(paths, export_pdf=True)
report = {'ok': True, 'counts': counts, 'layout': result}
(work / 'verification.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
print(json.dumps(report, indent=2))
