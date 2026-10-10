"""Keep separately tracked Lazada parcels under the same order number."""
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch
import hashlib
import json
import os
import sys
import unittest

APP = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP / 'src'))
from extractors import Batch, Item, Order, PackingError, lazada_page, parse_pdf


def source_page(order_id='123456', tracking='TRACK-A', quantities=(1,), label='Tracking ID'):
    text = f'Lazada Seller Center\nOrder Number：{order_id}\n'
    if tracking is not None:
        text += f'{label}：{tracking}\n'
    table = [['Item name', 'Item sku seller', 'Item variation', 'Quantity']]
    table.extend(['Source product', 'TEST-SKU-1L', 'Volume:1L', str(qty)] for qty in quantities)
    return SimpleNamespace(extract_text=lambda: text, extract_tables=lambda: [table])


def parse_pages(pages):
    pdf = MagicMock()
    pdf.__enter__.return_value = SimpleNamespace(pages=pages)
    with patch('extractors.pdfplumber.open', return_value=pdf), patch('extractors.file_hash', return_value='same-source-hash'):
        return parse_pdf(APP / 'qa' / 'synthetic-lazada-source.pdf')


class LazadaTrackingTests(unittest.TestCase):
    def test_tracking_is_extracted_with_fullwidth_colon(self):
        order = lazada_page(source_page(tracking='MP1504690825'), 1)
        self.assertEqual(order.tracking_id, 'MP1504690825')
        self.assertEqual(order.key, 'LAZADA:123456:TRACKING:MP1504690825')

    def test_tracking_number_label_and_case_are_supported(self):
        order = lazada_page(source_page(tracking='lex-ph-123', label='Tracking Number'), 1)
        self.assertEqual(order.tracking_id, 'LEX-PH-123')

    def test_same_order_different_tracking_keeps_both_parcels(self):
        batch = parse_pages([source_page(tracking='TRACK-A', quantities=(1, 1)),
                             source_page(tracking='TRACK-B', quantities=(3,))])
        self.assertEqual([o.order_id for o in batch.orders], ['123456', '123456'])
        self.assertEqual([o.tracking_id for o in batch.orders], ['TRACK-A', 'TRACK-B'])
        self.assertEqual([o.pieces for o in batch.orders], [2, 3])
        self.assertEqual([len(o.items) for o in batch.orders], [1, 1])
        self.assertEqual(batch.pieces, 5)

    def test_identical_items_in_different_parcels_are_not_dropped(self):
        batch = parse_pages([source_page(tracking='TRACK-A'), source_page(tracking='TRACK-B')])
        self.assertEqual(len(batch.orders), 2)
        self.assertEqual(batch.pieces, 2)
        self.assertEqual(batch.orders[0].fingerprint(), batch.orders[1].fingerprint())
        self.assertNotEqual(batch.orders[0].key, batch.orders[1].key)

    def test_same_order_and_tracking_still_rejects_duplicate_labels(self):
        with self.assertRaisesRegex(PackingError, 'Tracking ID TRACK-A appears more than once'):
            parse_pages([source_page(), source_page(quantities=(2,))])

    def test_tracking_must_not_discard_order_number_from_identity(self):
        batch = parse_pages([source_page(order_id='123'), source_page(order_id='456')])
        self.assertEqual(len(batch.orders), 2)
        self.assertNotEqual(batch.orders[0].key, batch.orders[1].key)

    def test_absent_tracking_preserves_legacy_source_support(self):
        order = lazada_page(source_page(tracking=None), 1)
        self.assertEqual(order.tracking_id, '')
        self.assertEqual(order.key, 'LAZADA:123456')
        with self.assertRaisesRegex(PackingError, 'appears more than once'):
            parse_pages([source_page(tracking=None), source_page(tracking=None)])

    def test_present_but_unsafe_tracking_is_not_silently_ignored(self):
        for value in ('', '---', 'TRACK ID', 'ID/OTHER', 'TRACK-A\nTracking ID：TRACK-B'):
            with self.subTest(value=value), self.assertRaisesRegex(PackingError, 'Tracking ID could not be read safely'):
                lazada_page(source_page(tracking=value), 1)
        page = source_page()
        page.extract_text = lambda: 'Lazada Seller Center\nOrder Number:123456\nTracking ID TRACK-A'
        with self.assertRaisesRegex(PackingError, 'Tracking ID could not be read safely'):
            lazada_page(page, 1)

    def test_tracking_does_not_change_source_item_fingerprint(self):
        items = [Item('Source product', '1L', 2)]
        expected = hashlib.sha256(json.dumps([('Source product', '1L', 2)], ensure_ascii=False).encode('utf-8')).hexdigest()
        for tracking in ('', 'TRACK-A', 'TRACK-B'):
            self.assertEqual(Order('LAZADA', '123', items, tracking_id=tracking).fingerprint(), expected)

    def test_new_field_preserves_comments_and_other_platform_keys(self):
        for platform, prefix in (('SPX', 'SHOPEE'), ('FLASH', 'SHOPEE'), ('TIKTOK', 'TIKTOK')):
            order = Order(platform, '123', [Item('Product', '1L', 1)], 'Buyer comment', 'TRACK-A')
            self.assertEqual(order.buyer_comment, 'Buyer comment')
            self.assertEqual(order.key, f'{prefix}:123')

    def test_source_snapshot_includes_tracking(self):
        batch = Batch('source.pdf', 'LAZADA', 'source-hash', [lazada_page(source_page(), 1)])
        self.assertEqual(batch.to_dict()['orders'][0]['tracking_id'], 'TRACK-A')

    def test_reported_pdf_retains_all_14_parcels_and_38_pieces(self):
        path = Path(os.environ.get('PACKING_LAZADA_SPLIT_PDF',
                                   r'C:\Users\Admin\Desktop\ONLINE SHOPS\20261007\LAZADA\LAZADA PACKING LIST 1.pdf'))
        if not path.exists():
            self.skipTest('Private Lazada source PDF is not available')
        batch = parse_pdf(path)
        self.assertEqual((len(batch.orders), len({o.order_id for o in batch.orders}), batch.pieces), (14, 13, 38))
        split = [o for o in batch.orders if o.order_id == '1128986725736007']
        self.assertEqual([(o.tracking_id, o.pieces) for o in split], [('820165464947', 4), ('310004587637', 16)])
        self.assertEqual([len(o.items) for o in split], [1, 1])
        self.assertTrue(all(o.items[0].seller_sku == 'SILICONE-OIL-1L' for o in split))


if __name__ == '__main__':
    unittest.main(verbosity=2)
