"""Read the supported Shopee, Lazada and TikTok packing-list layouts.

No OCR or guessed quantities: an unfamiliar layout is a review error.
"""
from __future__ import annotations

from collections import OrderedDict
from dataclasses import asdict, dataclass, field
from pathlib import Path
import hashlib
import re

import pdfplumber


class PackingError(Exception):
    pass


def clean(value: str | None) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def variation_text(value: str) -> str:
    # Shopee may wrap a size inside the digits, e.g. "Bubblegum,2\n50 mL".
    # Join only digits immediately before a known size unit, not arbitrary words.
    return re.sub(r'(,\d)\s+(\d{2,3}\s*(?:mL|ml|ML)\b)', r'\1\2', value)


@dataclass
class Item:
    product: str
    variation: str
    pieces: int


@dataclass
class Order:
    platform: str
    order_id: str
    items: list[Item] = field(default_factory=list)

    @property
    def key(self):
        return f"{'SHOPEE' if self.platform in ('SPX', 'FLASH') else self.platform}:{self.order_id}"

    @property
    def pieces(self):
        return sum(i.pieces for i in self.items)

    def fingerprint(self):
        # Whitespace normalization tolerates line wrapping, without changing words.
        import json
        rows = sorted((clean(i.product), clean(i.variation), i.pieces) for i in self.items)
        return hashlib.sha256(json.dumps(rows, ensure_ascii=False).encode('utf-8')).hexdigest()


@dataclass
class Batch:
    path: str
    platform: str
    sha256: str
    orders: list[Order]

    @property
    def group(self):
        return 'LAZADA_TIKTOK' if self.platform in ('LAZADA', 'TIKTOK') else self.platform

    @property
    def pieces(self):
        return sum(o.pieces for o in self.orders)

    def to_dict(self):
        return asdict(self)


def file_hash(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def region(page, left, top, right, bottom):
    # Character positions separate touching text such as "1L1" across columns.
    part = page.filter(lambda x: x.get('object_type') == 'char'
                       and left <= (x['x0'] + x['x1']) / 2 < right
                       and top <= (x['top'] + x['bottom']) / 2 < bottom)
    return clean(part.extract_text(x_tolerance=1.5, y_tolerance=2))


def positive_integer(value, context):
    value = clean(value)
    if not re.fullmatch(r'[0-9]+', value) or int(value) < 1:
        raise PackingError(f'{context}: the number of pieces could not be read safely ({value!r}).')
    return int(value)


def shopee_page(page, platform, page_num):
    if abs(page.width - 595) > 8 or abs(page.height - 842) > 8:
        raise PackingError(f'Page {page_num}: unsupported Shopee page size. Download the four-per-page packing list.')
    orders = []
    for row in range(2):
        for col in range(2):
            left, top = col * page.width / 2, row * page.height / 2
            p = page.crop((left, top, left + page.width / 2, top + page.height / 2))
            txt = p.extract_text() or ''
            ids = re.findall(r'Order\s+ID\s*[:：]\s*([A-Z0-9]+)', txt)
            if not ids:
                if clean(txt):
                    raise PackingError(f'Page {page_num}: a Shopee label has text but no readable Order ID.')
                continue
            if len(ids) != 2 or ids[0] != ids[1]:
                raise PackingError(f'Page {page_num}: incomplete or unfamiliar Shopee order label. Review {ids[0]}.')
            packages = re.findall(r'package\s+(\d+)', txt, flags=re.I)
            if packages != ['1', '1']:
                raise PackingError(f'Order {ids[0]}: multiple packages need review before packing.')
            words = p.extract_words(x_tolerance=1, y_tolerance=2)
            hashes = [w for w in words if w['text'] == '#']
            if len(hashes) != 1:
                raise PackingError(f'Order {ids[0]}: product columns were not recognized.')
            header_y = hashes[0]['top']
            header = [w for w in words if abs(w['top'] - header_y) < 1.5]
            def anchor(label):
                matches = [w for w in header if w['text'] == label]
                if len(matches) != 1:
                    raise PackingError(f'Order {ids[0]}: missing {label} column.')
                return matches[0]['x0'] - .25
            xproduct, xsku, xvar, xqty, xunit = [anchor(s) for s in ('Product', 'SKU', 'Variation', 'Qty', 'Unit')]
            footer = [w for w in words if w['text'] == 'Order' and w['top'] > header_y]
            if len(footer) != 1:
                raise PackingError(f'Order {ids[0]}: product table ending is unclear.')
            end_y = footer[0]['top'] - .5
            starts = [w for w in words if re.fullmatch(r'\d+', w['text'])
                      and abs(w['x0'] - hashes[0]['x0']) < 2
                      and header_y + 10 < w['top'] < end_y]
            starts.sort(key=lambda w: w['top'])
            if [int(w['text']) for w in starts] != list(range(1, len(starts) + 1)) or not starts:
                raise PackingError(f'Order {ids[0]}: missing or out-of-order product rows.')
            items = []
            for index, start in enumerate(starts):
                y1 = start['top'] - .5
                y2 = starts[index + 1]['top'] - .5 if index + 1 < len(starts) else end_y
                product = region(p, xproduct, y1, xsku, y2)
                variation = variation_text(region(p, xvar, y1, xqty, y2)) or '—'
                qty = positive_integer(region(p, xqty, y1, xunit, y2), f'Order {ids[0]}, item {index + 1}')
                if not product:
                    raise PackingError(f'Order {ids[0]}: a product name is blank.')
                items.append(Item(product, variation, qty))
            orders.append(Order(platform, ids[0], items))
    return orders


def lazada_page(page, page_num):
    text = page.extract_text() or ''
    ids = re.findall(r'Order\s+Number\s*[:：]\s*(\d+)', text)
    if len(ids) != 1:
        raise PackingError(f'Page {page_num}: expected one Lazada Order Number.')
    matches = []
    for table in page.extract_tables():
        if table and len(table[0]) == 4 and [clean(c).lower() for c in table[0]] == [
            'item name', 'item sku seller', 'item variation', 'quantity']:
            matches.append(table)
    if len(matches) != 1 or len(matches[0]) < 2:
        raise PackingError(f'Order {ids[0]}: the Lazada product table was not recognized.')
    items = []
    for index, row in enumerate(matches[0][1:], 1):
        if not any(clean(c) for c in row):
            continue
        product, _, variation, qty = row
        if not clean(product):
            raise PackingError(f'Order {ids[0]}: missing product name on row {index}.')
        items.append(Item(clean(product), clean(variation) or '—', positive_integer(qty, f'Order {ids[0]}, row {index}')))
    if not items:
        raise PackingError(f'Order {ids[0]} has no readable products.')
    return Order('LAZADA', ids[0], items)


def tiktok_page(page, page_num):
    text = page.extract_text() or ''
    ids = re.findall(r'Order\s+ID\s*[:：]\s*(\d+)', text)
    if len(ids) != 2 or ids[0] != ids[1]:
        raise PackingError(f'Page {page_num}: incomplete or unfamiliar TikTok order label.')
    totals = re.findall(r'Qty\s+Total\s*[:：]\s*(\d+)', text)
    words = page.extract_words(x_tolerance=1, y_tolerance=2)
    product_headers = [w for w in words if w['text'] == 'Product']
    candidates = []
    for h in product_headers:
        line = [w for w in words if abs(w['top'] - h['top']) < 1]
        if [w['text'] for w in line] == ['Product', 'Name', 'SKU', 'Seller', 'SKU', 'Qty']:
            candidates.append(line)
    if len(candidates) != 1 or len(totals) != 1:
        raise PackingError(f'Order {ids[0]}: TikTok columns or quantity total were not recognized.')
    header = candidates[0]
    xproduct, xsku, xseller, xqty = [header[i]['x0'] - .25 for i in (0, 2, 3, 5)]
    total_word = [w for w in words if w['text'] == 'Total:' and w['top'] > header[0]['top']]
    if len(total_word) != 1:
        raise PackingError(f'Order {ids[0]}: quantity total position is unclear.')
    end_y = total_word[0]['top'] - .5
    starts = [w for w in words if w['x0'] >= xqty and header[0]['bottom'] < w['top'] < end_y]
    if not starts:
        raise PackingError(f'Order {ids[0]}: no product quantities found.')
    items = []
    for index, start in enumerate(starts):
        y1 = start['top'] - .6
        y2 = starts[index + 1]['top'] - .6 if index + 1 < len(starts) else end_y
        product = region(page, xproduct, y1, xsku, y2)
        sku = region(page, xsku, y1, xseller, y2)
        seller = region(page, xseller, y1, xqty, y2)
        qty = positive_integer(start['text'], f'Order {ids[0]}, item {index + 1}')
        if not product:
            raise PackingError(f'Order {ids[0]}: a product name is blank.')
        items.append(Item(product, ', '.join(v for v in (sku, seller) if v) or '—', qty))
    order = Order('TIKTOK', ids[0], items)
    if order.pieces != int(totals[0]):
        raise PackingError(f'Order {ids[0]}: extracted quantities disagree with the printed total.')
    return order


def parse_pdf(path: str | Path, carrier: str | None = None) -> Batch:
    path = Path(path).resolve()
    before = file_hash(path)
    orders = []
    try:
        with pdfplumber.open(path) as pdf:
            if not pdf.pages:
                raise PackingError('The PDF is empty.')
            first = pdf.pages[0].extract_text() or ''
            if not clean(first):
                raise PackingError('This PDF is scanned or has no readable text. Download the original packing-list PDF.')
            if 'Lazada Seller Center' in first:
                platform = 'LAZADA'
            elif 'Seller SKU' in first and 'Qty Total' in first:
                platform = 'TIKTOK'
            elif 'Parent Product Name' in first and 'Order ID' in first:
                hints = set(re.findall(r'(?<![A-Z])(SPX|FLASH)(?![A-Z])', str(path).upper()))
                if carrier in ('SPX', 'FLASH'):
                    if len(hints) == 1 and carrier not in hints:
                        raise PackingError(f'This file or folder is labelled {next(iter(hints))}, but {carrier} was selected. Choose the correct courier or Detect automatically.')
                    platform = carrier
                elif len(hints) == 1:
                    platform = hints.pop()
                else:
                    raise PackingError('Use SPX or FLASH as a separate label in this Shopee PDF filename or folder path. Keep only one courier label in its path, then add or scan it again.')
            else:
                raise PackingError('PDF layout not recognized. Use an original Shopee, Lazada or TikTok packing list.')
            for n, page in enumerate(pdf.pages, 1):
                if platform in ('SPX', 'FLASH'):
                    orders.extend(shopee_page(page, platform, n))
                elif platform == 'LAZADA':
                    orders.append(lazada_page(page, n))
                else:
                    orders.append(tiktok_page(page, n))
    except PackingError:
        raise
    except Exception as exc:
        raise PackingError(f'Cannot read this PDF: {exc}') from exc
    if not orders:
        raise PackingError('No orders found.')
    combined = OrderedDict()
    for order in orders:
        if order.key in combined:
            raise PackingError(f'Order {order.order_id} appears more than once in this PDF. Review for duplicate labels or continued pages.')
        if platform == 'LAZADA':
            grouped = OrderedDict()
            for item in order.items:
                key = (item.product, item.variation)
                if key in grouped:
                    grouped[key].pieces += item.pieces
                else:
                    grouped[key] = item
            order.items = list(grouped.values())
        combined[order.key] = order
    if file_hash(path) != before:
        raise PackingError('The PDF changed while it was being read. Wait for the download to finish and scan again.')
    return Batch(str(path), platform, before, list(combined.values()))
