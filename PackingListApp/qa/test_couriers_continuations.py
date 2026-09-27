"""Shopee continuation, per-PDF couriers, and quiet PDF discovery."""
from pathlib import Path
import copy
import os
import shutil
import sys
import tempfile
import unittest
from docx import Document
from pypdf import PdfWriter
from reportlab.pdfgen import canvas

APP=Path(__file__).resolve().parents[1]
ROOT=Path(os.environ.get('PACKING_TEST_DATA',APP/'qa/names-regression-fixtures'))
sys.path.insert(0,str(APP/'src'))
import engine
from extractors import Item,Order,ShopeePart,PackingError,UnsupportedPdf,parse_pdf,combine_shopee_parts,auto_courier

class ContinuationTests(unittest.TestCase):
    def part(self,first,last,order='A',variation='Original'):
        return ShopeePart(Order('SPX',order,[Item('Product '+str(i),variation,i) for i in range(first,last+1)]),first,last)
    def test_continued_items_are_one_order_in_sequence(self):
        orders=combine_shopee_parts([self.part(1,9),self.part(10,11),self.part(1,2,'B')])
        self.assertEqual([o.order_id for o in orders],['A','B']);self.assertEqual(len(orders[0].items),11)
        self.assertEqual(orders[0].pieces,sum(range(1,12)))
    def test_repeated_complete_label_does_not_double_count(self):
        with self.assertRaisesRegex(PackingError,'expected item 3, found 1'):combine_shopee_parts([self.part(1,2),self.part(1,2)])
    def test_missing_or_overlapping_continuation_is_rejected(self):
        for parts in [[self.part(10,11)],[self.part(1,9),self.part(11,12)],[self.part(1,9),self.part(9,11)]]:
            with self.subTest(parts=parts),self.assertRaises(PackingError):combine_shopee_parts(parts)
    def test_more_than_two_sections_and_interleaved_orders_preserve_items(self):
        orders=combine_shopee_parts([self.part(1,2),self.part(1,1,'B'),self.part(3,4),self.part(5,6)])
        self.assertEqual(len(orders[0].items),6);self.assertEqual(len(orders[1].items),1)
    def test_reported_pdf_continuation(self):
        path=Path(os.environ.get('PACKING_CONTINUATION_PDF',r'C:\Users\Admin\Desktop\ONLINE SHOPS\20260928\SPX\PACKING LIST 1.pdf'))
        if not path.exists():self.skipTest('Private continuation PDF not restored')
        batch=parse_pdf(path)
        self.assertEqual((len(batch.orders),batch.pieces),(127,300))
        order=next(o for o in batch.orders if o.order_id=='2609263ADBUPHF')
        self.assertEqual((len(order.items),order.pieces),(11,14))
        self.assertEqual(order.items[-1].variation,'Premix 1L');self.assertEqual(order.items[-1].pieces,3)
        self.assertEqual(order.items[8].pieces,2)

class CourierTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(dir=APP/'qa');self.root=Path(self.temp.name)
        self.folder=self.root/'20260928';self.folder.mkdir()
    def tearDown(self):self.temp.cleanup()
    def copy(self,relative='packing.pdf',platform='SPX'):
        source=next((ROOT/'20260926').rglob({'SPX':'20260926_SPX_PL_4.pdf','FLASH':'20260926_Flash_PL_3.pdf','LAZADA':'20260926_Lazada_PL_1.pdf','TIKTOK':'20260926_Tiktok_PL_1.pdf'}[platform]))
        target=self.folder/relative;target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(source,target);return target
    def entry(self):return engine.scan(self.folder)['entries'][0]
    def choose(self,mode,entry=None):
        e=entry or self.entry();return engine.set_courier(self.folder,{'path':e['path'],'sha256':e['sha256']},mode)
    def test_auto_uses_nearest_classified_folder_before_filename(self):
        for relative,expected in [('SPX/FLASH_old.pdf','SPX'),('FLASH/SPX_old.pdf','FLASH'),('FLASH/SPX/Packing.pdf','SPX'),('Misc/abc_FLASH_1.pdf','FLASH'),('Misc/Packing.pdf',None)]:
            self.assertEqual(auto_courier(self.folder/relative),expected)
    def test_unclassified_shopee_visible_with_counts_then_manual_choice_persists(self):
        self.copy();e=self.entry()
        self.assertEqual((e['platform'],e['courier'],e['status'],e['orders'],e['pieces']),('SHOPEE','Auto','COURIER',7,8))
        with self.assertRaises(PackingError):engine.generate(self.folder,[{'path':e['path'],'sha256':e['sha256']}],verifier=lambda paths:[])
        self.choose('FLASH');e=self.entry()
        self.assertEqual((e['platform'],e['courier'],e['status']),('FLASH','FLASH','READY'))
        self.assertEqual(engine.load_couriers(self.folder)['packing.pdf']['carrier'],'FLASH')
        self.choose('Auto');self.assertEqual(self.entry()['status'],'COURIER')
    def test_auto_folder_and_explicit_override_and_reset(self):
        self.copy('SPX/PACKING.pdf');e=self.entry()
        self.assertEqual((e['platform'],e['courier']),('SPX','Auto'))
        self.choose('FLASH');self.assertEqual(self.entry()['platform'],'FLASH')
        self.choose('Auto');self.assertEqual((self.entry()['platform'],self.entry()['courier']),('SPX','Auto'))
    def test_lazada_and_tiktok_outside_platform_folders_need_no_courier(self):
        self.copy('other/lazada.pdf','LAZADA');self.copy('tiktok.pdf','TIKTOK')
        entries=engine.scan(self.folder)['entries'];self.assertEqual(len(entries),2)
        self.assertTrue(all(not e['is_shopee'] and e['status']=='READY' for e in entries))
        for e in entries:
            with self.assertRaises(PackingError):self.choose('SPX',e)
    def test_unknown_blank_and_corrupt_pdfs_hidden_but_recognized_errors_remain(self):
        self.copy()
        writer=PdfWriter();writer.add_blank_page(595,842)
        with (self.folder/'blank.pdf').open('wb') as f:writer.write(f)
        (self.folder/'corrupt.pdf').write_bytes(b'not a PDF')
        c=canvas.Canvas(str(self.folder/'unrelated.pdf'));c.drawString(40,800,'Unrelated invoice');c.save()
        c=canvas.Canvas(str(self.folder/'broken-packing.pdf'),pagesize=(595,842))
        c.drawString(10,800,'Packing List Order ID: TEST123 package 1');c.drawString(10,780,'Parent Product Name SKU Variation Qty Unit Total');c.save()
        entries=engine.scan(self.folder)['entries'];self.assertEqual({e['path'] for e in entries},{'packing.pdf','broken-packing.pdf'})
        broken=next(e for e in entries if e['path']=='broken-packing.pdf')
        self.assertEqual(broken['status'],'ERROR');self.assertTrue(broken['is_shopee'])
    def test_changed_pdf_cannot_reuse_old_courier_choice_or_stale_ui(self):
        path=self.copy();self.choose('SPX');e=self.entry()
        with path.open('ab') as f:f.write(b'\n% changed source\n')
        self.assertEqual((self.entry()['courier'],self.entry()['status']),('Auto','COURIER'))
        with self.assertRaisesRegex(PackingError,'changed after'):self.choose('FLASH',e)
    def test_generation_uses_selected_courier_and_rejects_stale_courier(self):
        self.copy();self.choose('FLASH');old=self.entry();self.choose('SPX')
        selected={'path':old['path'],'sha256':old['sha256'],'courier':old['courier'],'platform':old['platform']}
        def layout(paths):return [{'path':str(p),'pages':2} for p in paths]
        with self.assertRaisesRegex(PackingError,'courier changed'):engine.generate(self.folder,[selected],verifier=layout)
        now=self.entry();selected.update(courier=now['courier'],platform=now['platform'])
        made=engine.generate(self.folder,[selected],verifier=layout)
        self.assertTrue(made['outputs'][0]['filename'].startswith('SPX_'))
        self.assertEqual(self.entry()['status'],'DONE')
    def test_output_folder_pdfs_excluded(self):
        self.copy();self.copy('000 Output/old.pdf')
        self.assertEqual(len(engine.scan(self.folder)['entries']),1)
    def test_courier_path_escape_rejected(self):
        self.copy();e=self.entry();e['path']='../outside.pdf'
        with self.assertRaises(PackingError):self.choose('SPX',e)

if __name__=='__main__':unittest.main(verbosity=2)
