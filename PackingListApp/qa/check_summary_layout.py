"""Check real Word page breaks, multipage summaries, and export QA PDFs."""
import json
import sys
from pathlib import Path

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP / 'src'))
from documents import build_document, verify_word
from extractors import Batch, Item, Order

work = APP / 'qa/packaged-smoke'
paths = sorted((work / '20260926/000 Output').glob('*.docx'))
assert len(paths) == 3
long_path = work / 'Long_summary_test.docx'
batch = Batch('test.pdf', 'SPX', '', [Order('SPX', f'TEST{i:03}', [Item(f'Test product {i:03} food grade 1 kg', 'Original', 2)]) for i in range(75)])
build_document([batch], 'SPX', '20260926', long_path)
results = verify_word(paths + [long_path], export_pdf=True)
import pdfplumber
for path in paths + [long_path]:
    with pdfplumber.open(path.with_suffix('.pdf')) as pdf:
        texts = [p.extract_text() or '' for p in pdf.pages]
        summary_page = next(i for i, text in enumerate(texts) if 'Product Picking Summary' in text)
        assert 'Order ID:' not in texts[summary_page] and 'Order Number:' not in texts[summary_page]
        if path == long_path:
            assert len(texts) - summary_page >= 2
            assert all('Total pieces' in text for text in texts[summary_page:])
            assert '150' in texts[-1]
        print(path.name, len(texts), 'pages; summary begins page', summary_page+1)
(work / 'layout-verification.json').write_text(json.dumps(results, indent=2), encoding='utf-8')
