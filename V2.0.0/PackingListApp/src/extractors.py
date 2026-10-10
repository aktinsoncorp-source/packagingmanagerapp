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


class UnsupportedPdf(PackingError):
    """No supported readable packing-list layout; omit from the employee grid."""
    pass


def clean(value: str | None) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def variation_text(value: str) -> str:
    # Shopee may wrap a size inside the digits, e.g. "Bubblegum,2\n50 mL".
    # Join only digits immediately before a known size unit, not arbitrary words.
    return re.sub(r'(,\d)\s+(\d{2,3}\s*(?:mL|ml|ML)\b)', r'\1\2', value)


def normalize_sku(value: str | None) -> str:
    """Seller SKU columns can wrap mid-code. Keep the unmodified column too."""
    return re.sub(r'\s+', '', value or '').upper()


@dataclass
class Item:
    product: str
    variation: str
    pieces: int
    unit_price: str | None = None
    total_price: str | None = None
    seller_sku: str = ''
    raw_seller_sku: str = ''
    platform_sku: str = ''


@dataclass
class Order:
    platform: str
    order_id: str
    items: list[Item] = field(default_factory=list)
    buyer_comment: str = ''
    tracking_id: str = ''

    @property
    def key(self):
        key = f"{'SHOPEE' if self.platform in ('SPX', 'FLASH') else self.platform}:{self.order_id}"
        # Lazada can split one order across several independently packed parcels.
        return f'{key}:TRACKING:{self.tracking_id}' if self.platform == 'LAZADA' and self.tracking_id else key

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


def region(page, left, top, right, bottom, raw=False):
    # Character positions separate touching text such as "1L1" across columns.
    part = page.filter(lambda x: x.get('object_type') == 'char'
                       and left <= (x['x0'] + x['x1']) / 2 < right
                       and top <= (x['top'] + x['bottom']) / 2 < bottom)
    value = part.extract_text(x_tolerance=1.5, y_tolerance=2) or ''
    return value if raw else clean(value)


def source_price(value, context, label='unit price'):
    from decimal import Decimal
    value = clean(value)
    if not value or value in ('\u2014', '-', 'N/A'):
        return None
    # Read the printed value only; no assumed currency or guessed price.
    if not re.fullmatch(r'(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+)(?:\.[0-9]{1,2})?', value):
        raise PackingError(f'{context}: the {label} could not be read safely ({value!r}).')
    return format(Decimal(value.replace(',', '')), '.2f')


def positive_integer(value, context):
    value = clean(value)
    if not re.fullmatch(r'[0-9]+', value) or int(value) < 1:
        raise PackingError(f'{context}: the number of pieces could not be read safely ({value!r}).')
    return int(value)


@dataclass
class ShopeePart:
    order: Order
    first_item: int
    last_item: int


def shopee_page(page, platform, page_num):
    if abs(page.width - 595) > 8 or abs(page.height - 842) > 8:
        raise PackingError(f'Page {page_num}: unsupported Shopee page size. Download the four-per-page packing list.')
    orders = []
    for row in range(2):
        for col in range(2):
            left, top = col * page.width / 2, row * page.height / 2
            p = page.crop((left, top, left + page.width / 2, top + page.height / 2))
            txt = p.extract_text() or ''
            comment = re.search(r"Buyer['\u2019]s\s+Comments?\s*:\s*(.*)\Z", txt, re.I | re.S)
            label_text = txt[:comment.start()] if comment else txt
            ids = re.findall(r'Order\s+ID\s*[:：]\s*([A-Z0-9]+)', label_text)
            if not ids:
                if clean(txt):
                    raise PackingError(f'Page {page_num}: a Shopee label has text but no readable Order ID.')
                continue
            if len(ids) not in (1, 2) or len(set(ids)) != 1:
                raise PackingError(f'Page {page_num}: incomplete or unfamiliar Shopee order label. Review {ids[0]}.')
            packages = re.findall(r'package\s+(\d+)', label_text, flags=re.I)
            if packages != ['1'] * len(ids):
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
            xtotal = anchor('Total')
            footer = [w for index, w in enumerate(words) if w['text'] == 'Order' and w['top'] > header_y
                      and re.match(r'ID\s*[:：]\s*' + re.escape(ids[0]), ' '.join(v['text'] for v in words[index + 1:index + 4]))]
            if not footer:
                raise PackingError(f'Order {ids[0]}: product table ending is unclear.')
            end_y = footer[0]['top'] - .5
            starts = [w for w in words if re.fullmatch(r'\d+', w['text'])
                      and abs(w['x0'] - hashes[0]['x0']) < 2
                      and header_y + 10 < w['top'] < end_y]
            starts.sort(key=lambda w: w['top'])
            numbers = [int(w['text']) for w in starts]
            if not numbers or numbers[0] < 1 or numbers != list(range(numbers[0], numbers[0] + len(numbers))):
                raise PackingError(f'Order {ids[0]}: missing or out-of-order product rows.')
            if numbers[0] == 1 and len(ids) != 2:
                raise PackingError(f'Order {ids[0]}: the first product section is missing its order heading.')
            items = []
            for index, start in enumerate(starts):
                y1 = start['top'] - .5
                y2 = starts[index + 1]['top'] - .5 if index + 1 < len(starts) else end_y
                product = region(p, xproduct, y1, xsku, y2)
                seller_sku = region(p, xsku, y1, xvar, y2, raw=True)
                variation = variation_text(region(p, xvar, y1, xqty, y2)) or '—'
                qty = positive_integer(region(p, xqty, y1, xunit, y2), f'Order {ids[0]}, item {index + 1}')
                if not product:
                    raise PackingError(f'Order {ids[0]}: a product name is blank.')
                context = f'Order {ids[0]}, item {index + 1}'
                items.append(Item(product, variation, qty,
                                  source_price(region(p, xunit, y1, xtotal, y2), context),
                                  source_price(region(p, xtotal, y1, p.bbox[2], y2), context, 'total price'),
                                  normalize_sku(seller_sku), seller_sku))
            orders.append(ShopeePart(Order(platform, ids[0], items,
                                          clean(comment[1]) if comment else ''), numbers[0], numbers[-1]))
    return orders


def lazada_page(page, page_num):
    text = page.extract_text() or ''
    ids = re.findall(r'Order\s+Number\s*[:：]\s*(\d+)', text)
    if len(ids) != 1:
        raise PackingError(f'Page {page_num}: expected one Lazada Order Number.')
    tracking_labels = re.findall(r'Tracking\s+(?:ID|Number|No\.?)[ \t]*[:：][ \t]*([^\r\n]*)', text, re.I)
    tracking_id = ''
    if re.search(r'Tracking\s+(?:ID|Number|No)\b', text, re.I):
        if len(tracking_labels) != 1 or not re.fullmatch(r'[A-Z0-9][A-Z0-9-]*', tracking_labels[0].strip(), re.I):
            raise PackingError(f'Order {ids[0]}: the Lazada Tracking ID could not be read safely.')
        tracking_id = tracking_labels[0].strip().upper()
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
        product, seller_sku, variation, qty = row
        if not clean(product):
            raise PackingError(f'Order {ids[0]}: missing product name on row {index}.')
        items.append(Item(clean(product), clean(variation) or '—', positive_integer(qty, f'Order {ids[0]}, row {index}'),
                          seller_sku=normalize_sku(seller_sku), raw_seller_sku=seller_sku or ''))
    if not items:
        raise PackingError(f'Order {ids[0]} has no readable products.')
    return Order('LAZADA', ids[0], items, tracking_id=tracking_id)


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
        raw_seller = region(page, xseller, y1, xqty, y2, raw=True)
        seller = clean(raw_seller)
        qty = positive_integer(start['text'], f'Order {ids[0]}, item {index + 1}')
        if not product:
            raise PackingError(f'Order {ids[0]}: a product name is blank.')
        # Retain the historical variation string so existing fingerprints and
        # processing coverage remain identical after the V2 upgrade.
        items.append(Item(product, ', '.join(v for v in (sku, seller) if v) or '—', qty,
                          seller_sku=normalize_sku(seller), raw_seller_sku=raw_seller,
                          platform_sku=sku))
    order = Order('TIKTOK', ids[0], items)
    if order.pieces != int(totals[0]):
        raise PackingError(f'Order {ids[0]}: extracted quantities disagree with the printed total.')
    return order


def combine_shopee_parts(parts):
    combined, last_items = OrderedDict(), {}
    for part in parts:
        order = part.order
        if order.key not in combined:
            if part.first_item != 1:
                raise PackingError(f'Order {order.order_id}: continuation starts at item {part.first_item}, but its first section is missing.')
            combined[order.key] = Order(order.platform, order.order_id, list(order.items), order.buyer_comment)
        else:
            expected = last_items[order.key] + 1
            if part.first_item != expected:
                raise PackingError(f'Order {order.order_id}: repeated or incomplete product section; expected item {expected}, found {part.first_item}.')
            combined[order.key].items.extend(order.items)
            previous = combined[order.key].buyer_comment
            if previous and order.buyer_comment and previous != order.buyer_comment:
                raise PackingError(f'Order {order.order_id}: continued sections have different buyer comments. Review the source PDF.')
            combined[order.key].buyer_comment = previous or order.buyer_comment
        last_items[order.key] = part.last_item
    return list(combined.values())


def auto_courier(path):
    # The nearest classified folder is authoritative in Auto, even if the
    # filename still contains an older courier label.
    for parent in path.parents:
        if parent.name.upper() in ('SPX', 'FLASH'):
            return parent.name.upper()
    hints = set(re.findall(r'(?<![A-Z])(SPX|FLASH)(?![A-Z])', path.name.upper()))
    if len(hints) == 1:
        return next(iter(hints))
    hints = set(re.findall(r'(?<![A-Z])(SPX|FLASH)(?![A-Z])', str(path.parent).upper()))
    return next(iter(hints)) if len(hints) == 1 else None


def parse_pdf(path: str | Path, carrier: str | None = None, allow_unassigned=False) -> Batch:
    path = Path(path).resolve()
    orders = []
    platform = None
    try:
        before = file_hash(path)
        with pdfplumber.open(path) as pdf:
            if not pdf.pages:
                raise UnsupportedPdf('The PDF is empty.')
            first = pdf.pages[0].extract_text() or ''
            if not clean(first):
                raise UnsupportedPdf('This PDF is scanned or has no readable text. Download the original packing-list PDF.')
            if 'Lazada Seller Center' in first:
                platform = 'LAZADA'
            elif 'Seller SKU' in first and 'Qty Total' in first:
                platform = 'TIKTOK'
            elif 'Parent Product Name' in first and 'Order ID' in first:
                platform = 'SHOPEE'
                detected = auto_courier(path)
                if carrier in ('SPX', 'FLASH'):
                    platform = carrier
                elif detected:
                    platform = detected
                elif not allow_unassigned:
                    raise PackingError('Choose SPX or FLASH for this Shopee PDF, or include one courier label in its filename or folder path for Auto.')
            else:
                raise UnsupportedPdf('PDF layout not recognized. Use an original Shopee, Lazada or TikTok packing list.')
            for n, page in enumerate(pdf.pages, 1):
                if platform in ('SPX', 'FLASH', 'SHOPEE'):
                    orders.extend(shopee_page(page, platform, n))
                elif platform == 'LAZADA':
                    orders.append(lazada_page(page, n))
                else:
                    orders.append(tiktok_page(page, n))
    except PackingError as exc:
        exc.platform = platform or ''
        raise
    except Exception as exc:
        error = PackingError(f'Cannot read this packing list: {exc}') if platform else UnsupportedPdf(f'Cannot read this PDF: {exc}')
        error.platform = platform or ''
        raise error from exc
    if not orders:
        error = PackingError('No orders found.'); error.platform = platform
        raise error
    if platform in ('SPX', 'FLASH', 'SHOPEE'):
        try:
            orders = combine_shopee_parts(orders)
        except PackingError as exc:
            exc.platform = platform
            raise
    combined = OrderedDict()
    for order in orders:
        if order.key in combined:
            identity = f'Order {order.order_id}' + (f', Tracking ID {order.tracking_id}' if order.tracking_id else '')
            raise PackingError(f'{identity} appears more than once in this PDF. Review for duplicate labels or continued pages.')
        if platform == 'LAZADA':
            grouped = OrderedDict()
            for item in order.items:
                key = (item.product, item.variation)
                if key in grouped:
                    if grouped[key].seller_sku != item.seller_sku:
                        raise PackingError(f'Order {order.order_id}: identical product descriptions have different seller SKUs. Review the source PDF before packing.')
                    grouped[key].pieces += item.pieces
                else:
                    grouped[key] = item
            order.items = list(grouped.values())
        combined[order.key] = order
    if file_hash(path) != before:
        raise PackingError('The PDF changed while it was being read. Wait for the download to finish and scan again.')
    return Batch(str(path), platform, before, list(combined.values()))
