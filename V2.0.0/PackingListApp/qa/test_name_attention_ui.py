"""Missing product names are visible reminders, never a generation gate."""
import copy
from pathlib import Path
import sys
import time
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from app import PackingApp, NameAttentionPrompt
from product_attention import catalog_attention
from products_ui import ProductsDialog


ROWS = [dict(id=key, sku="" if key == "sample" else "UNSEEN-SKU",
             original_product="TEST MOONLIGHT VANILLA " + key, original_variation="50 ML",
             alternative_product="", alternative_variation="", locked=True)
        for key in ("sample", "unrelated")]
REPORT = catalog_attention({"rows": ROWS[:1]})


class NameAttentionTests(unittest.TestCase):
    def setUp(self):
        self.app = PackingApp(data_dir=Path(__file__).parent / "unused-attention-data", auto_scan=False)
        self.app.geometry("1100x688")
        self.app.update()

    def tearDown(self):
        for child in self.app.winfo_children():
            if isinstance(child, (NameAttentionPrompt, ProductsDialog)):
                child.destroy()
        self.app.busy = False
        self.app.close()

    def wait_for(self, condition, timeout=3):
        until = time.monotonic() + timeout
        while time.monotonic() < until:
            self.app.update()
            if condition():
                return
            time.sleep(.01)
        self.fail("Timed out waiting for UI")

    def prompts(self):
        return [child for child in self.app.winfo_children() if isinstance(child, NameAttentionPrompt)]

    def test_scan_later_suppresses_repeat_prompt_and_generation_still_runs(self):
        entry = dict(path="FLASH/test.pdf", platform="FLASH", status="READY", sha256="hash",
                     is_shopee=True, courier="Auto", orders=1, new_orders=1, message="Ready",
                     order_totals=[{"key": "FLASH:TEST1"}])
        response = dict(ok=True, entries=[entry], outputs=[], name_scan_attention=REPORT,
                        name_scan_ids=["sample"], name_source_issues=[])
        self.app.engine_handler = lambda request: copy.deepcopy(response)
        self.app.scan()
        self.wait_for(lambda: bool(self.prompts()))
        self.assertEqual(self.app.name_scan_ids, {"sample"})
        self.prompts()[0].later_button.invoke()
        self.app.update()
        self.assertIsNone(self.app.grab_current())
        self.assertTrue(self.app.attention_button.winfo_viewable())
        self.assertFalse(self.app.generate_button.instate(["disabled"]))
        self.app.scan(preserve=True)
        self.wait_for(lambda: not self.app.busy)
        self.app._prompt_name_attention()
        self.assertEqual(self.prompts(), [])
        with patch.object(self.app, "call_engine") as engine_call:
            self.app.generate()
        self.assertEqual(engine_call.call_args.args[0]["action"], "generate")
        self.assertEqual(engine_call.call_args.args[0]["output_format"], "pdf")

    def test_pulse_changes_color_without_hiding_text_and_cancels_when_resolved(self):
        self.app._set_name_attention(REPORT)
        self.app.update()
        button = self.app.attention_button
        before = (button.cget("text"), button.winfo_width())
        self.wait_for(lambda: button.cget("style") == "AttentionPulse.TButton", timeout=2)
        self.assertEqual((button.cget("text"), button.winfo_width()), before)
        self.assertTrue(button.winfo_viewable())
        self.app._set_busy(True, "Working")
        self.wait_for(lambda: button.cget("style") == "Attention.TButton", timeout=2)
        self.app._set_busy(False, "Ready")
        self.app._set_name_attention({"count": 0, "items": []})
        self.app.update()
        self.assertFalse(button.winfo_viewable())
        self.assertIsNone(self.app._attention_timer)

    def test_review_button_opens_only_current_affected_products_locked(self):
        self.app.name_scan_ids = {"sample"}
        self.app.engine_handler = lambda request: dict(ok=True, revision=1, rows=copy.deepcopy(ROWS))
        self.app._set_name_attention(REPORT)
        self.app.attention_button.invoke()
        self.wait_for(lambda: any(isinstance(child, ProductsDialog) for child in self.app.winfo_children()))
        dialog = next(child for child in self.app.winfo_children() if isinstance(child, ProductsDialog))
        self.assertEqual(dialog.tree.get_children(), ("sample",))
        self.assertTrue(dialog.model.rows["sample"]["locked"])
        self.assertIn("attentionselected", dialog.tree.item("sample", "tags"))
        dialog._destroy()

    def test_new_reason_prompts_after_earlier_advisory_was_dismissed(self):
        report = copy.deepcopy(REPORT)
        report["items"][0]["reasons"] = ["missing_product_name"]
        self.app._set_name_attention(report)
        self.app._prompt_name_attention()
        self.prompts()[0].destroy()
        self.app._set_name_attention(REPORT)
        self.app._prompt_name_attention()
        self.assertEqual(len(self.prompts()), 1)


if __name__ == "__main__":
    unittest.main()
