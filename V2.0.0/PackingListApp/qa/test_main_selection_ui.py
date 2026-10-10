"""Exercise actual mouse events for the main PDF Use heading."""
from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from app import PackingApp, selection_counts


class ShipmentCountTests(unittest.TestCase):
    def test_distinct_orders_and_lazada_shipments_count_separately(self):
        entries = [dict(path='one.pdf', platform='LAZADA', status='READY', order_totals=[
            {'key': 'LAZADA:001:TRACKING:A'}, {'key': 'LAZADA:001:TRACKING:B'}]),
            dict(path='copy.pdf', platform='LAZADA', status='READY', order_totals=[
                {'key': 'LAZADA:001:TRACKING:A'}])]
        self.assertEqual(selection_counts(entries, {'one.pdf', 'copy.pdf'}),
                         {'LAZADA': {'pdfs': 2, 'orders': 1, 'shipments': 2}})
        self.assertEqual(selection_counts(entries, {'copy.pdf'}), {'LAZADA': {'pdfs': 1, 'orders': 1}})


class UseHeadingTests(unittest.TestCase):
    def setUp(self):
        self.app = PackingApp(data_dir=Path(__file__).parent / "unused-ui-data", auto_scan=False)
        self.app.geometry("1100x688")
        self.app.entries = [dict(
            path=f"FLASH/{i}.pdf", platform="FLASH", status="DUPLICATE",
            is_shopee=True, courier="Auto", orders=1, message="Process again",
            order_totals=[{"key": f"FLASH:{i}"}]) for i in range(15)]
        for status in ("READY", "ERROR", "COURIER", "REVIEW"):
            self.app.entries.append(dict(path=status+".pdf", platform="SPX", status=status,
                is_shopee=True, courier="Auto", orders=1, message=status,
                requires_revision=status == "REVIEW", order_totals=[{"key": status}]))
        self.app.checked = {"FLASH/1.pdf", "FLASH/6.pdf"}
        self.app._populate_files()
        self.app.update()
        self.eligible = {f"FLASH/{i}.pdf" for i in range(15)} | {"READY.pdf"}
        self.clock = 1000

    def tearDown(self):
        if self.app.busy:
            self.app._set_busy(False, "Ready")
        self.app.close()

    def click_heading(self, x, elapsed=1000):
        tree = self.app.tree
        self.assertEqual(tree.identify_region(x, 20), "heading")
        tree.event_generate("<Motion>", x=x, y=20)
        tree.event_generate("<Button-1>", x=x, y=20, time=self.clock)
        tree.event_generate("<ButtonRelease-1>", x=x, y=20, time=self.clock+20)
        self.app.update()
        self.clock += elapsed

    def test_icon_and_label_select_all_completed_pdfs_then_clear(self):
        self.click_heading(20)
        self.assertEqual(self.app.checked, self.eligible)
        self.assertEqual(self.app.tree.heading("#0", "image"), (str(self.app.check_icons["on"]),))
        self.click_heading(60)
        self.assertEqual(self.app.checked, set())
        self.assertEqual(self.app.tree.heading("#0", "image"), (str(self.app.check_icons["off"]),))
        self.app.revisions.add("REVIEW.pdf")
        self.click_heading(20)
        self.assertEqual(self.app.checked, self.eligible | {"REVIEW.pdf"})

    def test_rapid_repeat_click_is_not_swallowed(self):
        self.click_heading(20, elapsed=100)
        self.assertEqual(self.app.checked, self.eligible)
        self.click_heading(20)
        self.assertEqual(self.app.checked, set())

    def test_header_does_nothing_while_working(self):
        before = self.app.checked.copy()
        self.app._set_busy(True, "Reading PDFs...")
        self.click_heading(20)
        self.assertEqual(self.app.checked, before)


if __name__ == "__main__":
    unittest.main()
