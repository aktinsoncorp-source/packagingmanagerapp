"""Run real Office export against a fresh isolated copy of the user's PDFs."""
from collections import Counter
from dataclasses import replace
from pathlib import Path
import argparse
import json
import shutil
import sys
import tempfile
import time

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP / 'src'))
import engine
import names
from extractors import file_hash, parse_pdf
from sku_registry import load_registry
import pdfplumber


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--reference', type=Path, default=APP.parent / 'Reference')
    parser.add_argument('--data-dir', type=Path, default=APP.parent / 'Packing List Manager')
    parser.add_argument('--destination', type=Path)
    args = parser.parse_args()
    root = args.destination
    if root:
        root.mkdir(parents=True, exist_ok=False)
    else:
        root = Path(tempfile.mkdtemp(prefix='reference-verification-', dir=APP / 'qa'))
    root = root.resolve()
    folder = root / '20261010'
    data = root / 'data'
    data.mkdir()
    for name in ('ProductNames.json', 'SkuCatalog.json'):
        shutil.copy2(args.data_dir / name, data / name)
    catalog_path = data / 'ProductNames.json'
    hashes = {}
    for source in sorted(args.reference.rglob('*.pdf')):
        relative = source.relative_to(args.reference)
        hashes[str(relative)] = file_hash(source)
        target = folder / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
    batches = [parse_pdf(p) for p in sorted(folder.rglob('*.pdf'))]
    original_orders = {order.key: order for batch in batches for order in batch.orders}
    assert len(original_orders) == sum(len(b.orders) for b in batches), 'Unexpected overlapping references'
    pieces = sum(b.pieces for b in batches)
    started = time.monotonic()
    scan = engine.handle({'action': 'scan', 'folder': str(folder), 'catalog_path': str(catalog_path)})
    print('Scan:', len(scan['entries']), 'PDFs;', len(original_orders), 'orders;', pieces,
          'pieces;', round(time.monotonic() - started, 2), 'seconds', flush=True)
    assert all(e['status'] == 'READY' for e in scan['entries']), scan['entries']
    catalog = names.load_catalog(catalog_path)
    registry = load_registry(catalog_path)
    printed = names.apply_names(batches, catalog, registry)
    changed = 0
    for batch, display_batch in zip(batches, printed):
        for order, display_order in zip(batch.orders, display_batch.orders):
            for item, displayed in zip(order.items, display_order.items):
                if item.seller_sku and registry.resolve(batch.platform, item.seller_sku):
                    identity, canonical = names._identity(batch.platform, item, registry)
                    definition = names._unique_sku_definition(
                        names._sku_matches(catalog['rows'], identity, batch.platform, canonical), identity['sku'])
                    if definition.get('alternative_product') and definition.get('alternative_variation'):
                        renamed = replace(item, product='CHANGED ECOMMERCE TITLE', variation='CHANGED ECOMMERCE VARIATION')
                        test = replace(batch, orders=[replace(order, items=[renamed])])
                        resolved = names.apply_names([test], catalog, registry)[0].orders[0].items[0]
                        assert (resolved.product, resolved.variation) == (displayed.product, displayed.variation)
                        changed += 1
    assert changed, 'No complete SKU definition exercised'
    selected = [{'path': e['path'], 'sha256': e['sha256']} for e in scan['entries']]
    request = {'action': 'generate', 'folder': str(folder), 'catalog_path': str(catalog_path), 'selected': selected}
    result = engine.handle(request)
    assert result['ok'] and len(result['outputs']) == 3, result
    layouts = []
    for output in result['outputs']:
        path = Path(output['path'])
        with pdfplumber.open(path) as pdf:
            texts = [page.extract_text() or '' for page in pdf.pages]
            for page in pdf.pages:
                assert abs(page.width - 595.3) < 2 and abs(page.height - 841.9) < 2, 'A4 changed'
                assert (page.crop((0, 0, page.width, page.height - 30)).extract_text() or '').strip(), 'Blank body page'
            summary = next(i for i, text in enumerate(texts) if 'Product Picking Summary' in text)
            assert summary > 0 and 'Order ID:' not in texts[summary] and 'Order Number:' not in texts[summary]
            layouts.append({'filename': path.name, 'pages': len(pdf.pages), 'summary_page': summary + 1})
            print('Verified PDF:', layouts[-1], flush=True)
    history_bytes = (folder / '000 Output' / engine.HISTORY).read_bytes()
    history = json.loads(history_bytes)
    audited = {}
    for output in history['outputs'].values():
        for key, raw in output['order_snapshots'].items():
            original = original_orders[key]
            assert output['order_fingerprints'][key] == original.fingerprint()
            assert raw['buyer_comment'] == original.buyer_comment
            for item, saved in zip(original.items, raw['items']):
                assert saved['seller_sku'] == item.seller_sku
                assert saved['raw_seller_sku'] == item.raw_seller_sku
                assert (saved['unit_price'], saved['total_price']) == (item.unit_price, item.total_price)
            audited[key] = raw
    assert set(audited) == set(original_orders)
    repeated = engine.handle(request)
    assert repeated['confirmation_required'] and repeated['repeat_count'] == len(original_orders)
    skipped = engine.handle(dict(request, repeat_confirmation=repeated['repeat_token'], repeat_action='skip'))
    assert not skipped['outputs'] and skipped['skipped_repeat_orders'] == len(original_orders)
    sales = engine.handle({'action': 'sales_export', 'folder': str(folder), 'catalog_path': str(catalog_path)})
    assert sales['quantity'] == pieces and sales['orders'] == len(original_orders), sales
    assert (folder / '000 Output' / engine.HISTORY).read_bytes() == history_bytes, 'Sales or No changed immutable history'
    assert hashes == {str(p.relative_to(args.reference)): file_hash(p) for p in args.reference.rglob('*.pdf')}
    report = {'ok': True, 'workspace': str(root), 'reference_hashes': hashes, 'pdfs': len(batches),
              'orders': len(original_orders), 'pieces': pieces, 'rename_checks': changed,
              'platform_orders': dict(Counter(o.platform for o in original_orders.values())),
              'layouts': layouts, 'sales': sales, 'repeat_no': skipped}
    (root / 'verification.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
    print('PASS:', root, flush=True)


if __name__ == '__main__':
    main()
