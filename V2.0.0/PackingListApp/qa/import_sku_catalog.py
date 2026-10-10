"""Prepare the private SKU master from local catalog exports. No marketplace writes.

Usage: python import_sku_catalog.py --ecommerce PATH --destination SkuCatalog.json
The optional --aliases JSON is the saved Lazada old-to-new SKU update plan.
"""
import argparse
from pathlib import Path
import hashlib
import json
import sys
import zipfile
import xml.etree.ElementTree as E

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from sku_registry import SkuRegistry, normalized_sku

N = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def read_rows(path, sheet_name):
    with zipfile.ZipFile(path) as z:
        strings = [''.join(s.itertext()) for s in E.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', N)] if 'xl/sharedStrings.xml' in z.namelist() else []
        relationships = {r.attrib['Id']: r.attrib['Target'] for r in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        sheets = E.fromstring(z.read('xl/workbook.xml')).findall('m:sheets/m:sheet', N)
        sheet = next(s for s in sheets if s.attrib['name'] == sheet_name)
        target = relationships[sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']]
        target = target.lstrip('/') if target.startswith('/') else 'xl/' + target
        rows = []
        for row in E.fromstring(z.read(target)).findall('m:sheetData/m:row', N):
            values = {}
            for c in row.findall('m:c', N):
                if c.find('m:f', N) is not None:
                    raise ValueError('Catalog identity fields must be plain text, not formulas.')
                v = c.find('m:v', N)
                value = v.text if v is not None else ''
                if c.attrib.get('t') == 's':
                    value = strings[int(value)]
                elif c.attrib.get('t') == 'inlineStr':
                    value = ''.join(c.find('m:is', N).itertext())
                values[''.join(x for x in c.attrib['r'] if x.isalpha())] = value or ''
            rows.append((int(row.attrib['r']), values))
        return rows


def prepare(root, alias_path=None):
    records, sources, identifiers = {}, [], {}
    for platform, filename, sheet, start, pc, vc, sc in [
        ('SHOPEE', 'Shopee Products.xlsx', 'Sheet1', 7, 'B', 'D', 'F'),
        ('LAZADA', 'Lazada Products.xlsx', 'template', 5, 'C', 'O', 'H'),
    ]:
        path = root / filename
        sources.append({'platform': platform, 'filename': filename, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()})
        seen = set()
        for number, values in read_rows(path, sheet):
            if number < start:
                continue
            sku = normalized_sku(values.get(sc))
            if not sku or sku in seen:
                raise ValueError(f'{filename} row {number}: blank or repeated SKU.')
            seen.add(sku)
            product, variation = values.get(pc, '').strip(), values.get(vc, '').strip() or '—'
            if not product:
                raise ValueError(f'{filename} row {number}: missing product.')
            entry = records.setdefault(sku, {'sku': sku, 'original_product': product.upper(), 'original_variation': variation.upper(), 'sku_aliases': [], 'ecommerce_sources': []})
            entry['sku_aliases'].append({'platform': platform, 'sku': sku})
            entry['ecommerce_sources'].append({'platform': platform, 'sku': sku, 'product': product, 'variation': variation})
            identifiers[(platform, sku)] = entry
    if alias_path:
        aliases = json.loads(alias_path.read_text(encoding='utf-8-sig'))
        for edit in aliases['edits']:
            entry = identifiers.get(('LAZADA', normalized_sku(edit['new_sku'])))
            if not entry:
                raise ValueError(f'Historical alias target {edit["new_sku"]} is absent from Lazada.')
            entry['sku_aliases'].append({'platform': 'LAZADA', 'sku': normalized_sku(edit['old_sku'])})
            entry['ecommerce_sources'].append({'platform': 'LAZADA', 'sku': normalized_sku(edit['old_sku']), 'product': edit['lazada_name'], 'variation': edit['lazada_variation'] or '—'})
    data = {'schema': 1, 'authority': 'SHOPEE', 'prepared_date': '2026-10-10', 'sources': sources, 'records': list(records.values())}
    SkuRegistry(data)
    return data


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--ecommerce', type=Path, required=True)
    parser.add_argument('--destination', type=Path, required=True)
    parser.add_argument('--aliases', type=Path)
    args = parser.parse_args()
    if args.destination.exists():
        raise SystemExit('Destination exists. Preserve it and choose a fresh catalog path.')
    result = prepare(args.ecommerce, args.aliases)
    args.destination.parent.mkdir(parents=True, exist_ok=True)
    args.destination.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'canonical_skus': len(result['records']), 'seller_aliases': sum(len(r['sku_aliases']) for r in result['records']), 'authority': 'SHOPEE'}))
