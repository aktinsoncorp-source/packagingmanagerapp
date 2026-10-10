"""Focused regressions for the SKU editor and its Tk save/lock workflow."""
import copy
import json
from pathlib import Path
import sys
import tkinter as tk
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from products_ui import ProductsDialog, ProductsModel, single_line, size_key


def product(identity="rose", **overrides):
    row = {"id": identity, "sku": "PERF-G-ROSE-W-100ML", "sku_scope": "CATALOG", "original_product": "G. ROSE PERFUME", "original_variation": "100ML",
           "alternative_product": "GOLDEA ROSE", "alternative_variation": "100ML", "locked": False,
           "aliases": [{"original_product": "EARLIER ROSE LISTING", "original_variation": "100 ML"}],
           "sku_aliases": [{"platform": "LAZADA", "sku": "PERF-G-ROSE-W-100ML"}],
           "ecommerce_sources": [{"platform": "SHOPEE", "sku": "PERF-G-ROSE-W-100ML", "product": "G. ROSE PERFUME", "variation": "100ML"}],
           "future_metadata": {"retain": [1, 2, 3]}}
    row.update(overrides)
    return row


class ProductsModelTests(unittest.TestCase):
    def test_open_locks_only_complete_rows_without_mutating_catalog(self):
        rows = [product(), product("partial", alternative_variation=""), product("empty", alternative_product="", alternative_variation="", locked=True)]
        original = copy.deepcopy(rows)
        model = ProductsModel({"revision": 7, "rows": rows})
        self.assertEqual([row["locked"] for row in model.rows.values()], [True, False, False])
        self.assertEqual(rows, original)
        self.assertFalse(model.dirty)

    def test_edit_retains_all_links_and_unknown_metadata(self):
        model = ProductsModel({"rows": [product()]})
        model.set_lock("rose", False)
        model.edit("rose", {"alternative_product": " @ inventory rose  ", "alternative_variation": "100 ml "})
        row = model.payload()[0]
        self.assertEqual(row["alternative_product"], "INVENTORY ROSE")
        self.assertEqual(row["alternative_variation"], "100 ML")
        for key in ("aliases", "sku_aliases", "sku_scope", "ecommerce_sources", "future_metadata"):
            self.assertEqual(row[key], product()[key])
        row["future_metadata"]["retain"].append(4)
        self.assertEqual(model.rows["rose"]["future_metadata"]["retain"], [1, 2, 3])

    def test_locked_row_and_saved_sku_are_protected(self):
        model = ProductsModel({"rows": [product()]})
        with self.assertRaises(ValueError):
            model.edit("rose", {"alternative_product": "OTHER"})
        model.set_lock("rose", False)
        with self.assertRaises(ValueError):
            model.edit("rose", {"sku": "DIFFERENT"})
        self.assertEqual(model.rows["rose"]["sku"], product()["sku"])

    def test_changed_saved_row_cannot_relock_before_saving(self):
        model = ProductsModel({"rows": [product()]})
        model.set_lock("rose", False)
        model.edit("rose", {"alternative_product": "NEW NAME"})
        with self.assertRaises(ValueError):
            model.set_lock("rose", True)
        self.assertFalse(model.rows["rose"]["locked"])
        self.assertEqual(model.unlocked_ids, {"rose"})

    def test_delete_remembers_explicit_unlock_and_skips_locked_rows(self):
        model = ProductsModel({"rows": [product(), product("other")]})
        self.assertEqual(model.delete(["rose", "other"]), 0)
        model.set_lock("rose", False)
        self.assertEqual(model.delete(["rose", "other"]), 1)
        self.assertEqual(model.unlocked_ids, {"rose"})
        self.assertTrue(model.dirty)

    def test_search_finds_sku_platform_and_linked_originals(self):
        model = ProductsModel({"rows": [product()]})
        for query in ("PERF-G-ROSE", "earlier rose", "lazada", "goldea"):
            self.assertEqual(len(model.sorted_rows(query)), 1)
        self.assertEqual(model.sorted_rows("not present"), [])
        self.assertEqual(model.sorted_rows(needs_names=True), [])

    def test_sort_families_are_independent_and_preserve_edits(self):
        model = ProductsModel({"rows": [product("large", alternative_variation="1L"), product("small", alternative_variation="30ML"), product("middle", alternative_variation="100ML")]})
        self.assertEqual([r["id"] for r in model.sorted_rows()], ["small", "middle", "large"])
        model.toggle_sort("alternative_variation")
        model.toggle_sort("original_product")
        self.assertTrue(model.size_desc)
        self.assertEqual([r["id"] for r in model.sorted_rows()], ["large", "middle", "small"])
        model.toggle_sort("alternative_variation")
        self.assertFalse(model.size_desc)
        model.set_lock("small", False)
        model.edit("small", {"alternative_product": "CHANGED"})
        model.toggle_sort("sku")
        self.assertTrue(model.changed("small"))
        self.assertFalse(model.rows["small"]["locked"])

    def test_sizes_normalize_within_unit_families(self):
        self.assertLess(size_key("30ML"), size_key("0.1L"))
        self.assertLess(size_key("100G"), size_key("1KG"))
        self.assertLess(size_key("1KG"), size_key("30ML"))

    def test_details_include_sources_legacy_links_and_unseen_seller_skus(self):
        row = product()
        row["sku_aliases"].append({"platform": "LAZADA", "sku": "OLD-ROSE-100"})
        details = ProductsModel.ecommerce_details(row)
        self.assertIn(("SHOPEE", row["sku"], "G. ROSE PERFUME", "100ML"), details)
        self.assertIn(("Earlier listing", "", "EARLIER ROSE LISTING", "100 ML"), details)
        self.assertIn(("LAZADA", "OLD-ROSE-100", "", ""), details)
        self.assertEqual(len(details), len(set(details)))

    def test_add_requires_sku_and_single_line_edits_are_atomic(self):
        model = ProductsModel({"rows": []})
        with self.assertRaises(ValueError):
            model.add({"sku": ""})
        identity = model.add({"sku": "perf-new-w-30ml", "alternative_product": "New", "alternative_variation": "30ML"})
        self.assertEqual(model.rows[identity]["sku"], "PERF-NEW-W-30ML")
        before = copy.deepcopy(model.rows[identity])
        with self.assertRaises(ValueError):
            model.edit(identity, {"alternative_product": "VALID", "alternative_variation": "Bad\nValue"})
        self.assertEqual(model.rows[identity], before)
        self.assertEqual(single_line(" @_ Name  ", alternative=True), "NAME")


class ProductsTkTests(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.requests = []
        self.fail_save = False

        def call(request, success, _message, on_error=None):
            self.requests.append(copy.deepcopy(request))
            if request["action"] == "names_prepare":
                success({"rows": request["rows"], "merges": []})
            elif request["action"] == "names_save":
                if self.fail_save:
                    on_error(RuntimeError("Simulated save failure"))
                    return
                rows = copy.deepcopy(request["rows"])
                if request["lock_all"]:
                    for row in rows:
                        row["locked"] = True
                success({"revision": 2, "rows": rows})
        self.dialog = ProductsDialog(self.root, {"revision": 1, "rows": [product(), product("partial", sku="PERF-NEW-W-30ML", alternative_product="", alternative_variation="")]}, call)
        self.root.update()

    def tearDown(self):
        self.dialog._destroy()
        self.root.destroy()

    def test_locked_on_open_search_details_and_global_row_lock_workflow(self):
        self.assertTrue(self.dialog.edit_locked)
        self.assertIn("disabled", self.dialog.save_button.state())
        self.assertFalse(self.dialog.save_names())
        self.assertEqual(tuple(self.dialog.tree.cget("displaycolumns")), ("sku", "alternative_product", "alternative_variation", "details"))
        self.dialog.search_var.set("earlier rose")
        self.root.update()
        self.assertEqual(len(self.dialog.tree.get_children()), 2)
        self.assertTrue(self.dialog.set_edit_locked(False))
        self.assertTrue(self.dialog.set_row_locked("rose", False))
        self.dialog.model.edit("rose", {"alternative_product": "INVENTORY ROSE"})
        self.dialog._refresh_rows()
        self.assertIn("dirty", self.dialog.tree.item("rose", "tags"))
        self.assertFalse(self.dialog.set_edit_locked(True))
        self.dialog.save_names()
        self.assertTrue(all(row["locked"] for row in self.dialog.model.rows.values()))
        self.assertFalse(self.dialog.model.dirty)
        self.assertTrue(self.dialog.set_edit_locked(True))
        self.assertEqual(self.requests[-1]["unlocked_ids"], ["rose"])

    def test_failed_save_keeps_edits_and_locks(self):
        self.dialog.set_edit_locked(False)
        self.dialog.set_row_locked("rose", False)
        self.dialog.model.edit("rose", {"alternative_product": "UNSAVED"})
        before = copy.deepcopy(self.dialog.model.rows)
        self.fail_save = True
        with patch("products_ui.messagebox.showwarning"):
            self.dialog.save_names()
        self.assertEqual(self.dialog.model.rows, before)
        self.assertTrue(self.dialog.model.dirty)
        self.assertFalse(self.dialog.busy)

    def test_pre_export_save_preserves_unlocks_and_hidden_rows(self):
        self.dialog.set_edit_locked(False)
        self.dialog.set_row_locked("rose", False)
        self.dialog.search_var.set("no matching row")
        self.dialog.save_names(lock_all=False)
        self.assertEqual(len(self.requests[-1]["rows"]), 2)
        self.assertFalse(self.dialog.model.rows["rose"]["locked"])
        self.assertFalse(self.dialog.model.rows["partial"]["locked"])

    def test_125_percent_small_window_keeps_actions_and_primary_columns_visible(self):
        call = self.dialog.call_engine
        self.dialog._destroy()
        self.root.tk.call("tk", "scaling", float(self.root.tk.call("tk", "scaling"))*1.25)
        self.root.deiconify()
        self.root.update()
        self.dialog = ProductsDialog(self.root, {"revision": 1, "rows": [product(str(index)) for index in range(1278)]}, call)
        self.dialog.geometry("1100x688")
        self.root.update()
        for name in ("save_button", "edit_button", "add_button", "details_button", "close_button", "export_button", "import_button", "delete_button", "lock_button", "unlock_button", "search_entry", "ecommerce_check", "status_label"):
            widget = getattr(self.dialog, name)
            x, y = widget.winfo_rootx()-self.dialog.winfo_rootx(), widget.winfo_rooty()-self.dialog.winfo_rooty()
            self.assertTrue(widget.winfo_viewable(), name)
            self.assertGreaterEqual(x, 0, name)
            self.assertGreaterEqual(y, 0, name)
            self.assertLessEqual(x+widget.winfo_width(), self.dialog.winfo_width(), name)
            self.assertLessEqual(y+widget.winfo_height(), self.dialog.winfo_height(), name)
        self.assertEqual(self.dialog.tree.xview(), (0.0, 1.0))
        self.assertGreaterEqual(self.dialog.tree.winfo_height(), 150)
        self.dialog.show_ecommerce.set(True)
        self.dialog._display_columns()
        self.assertEqual(self.dialog.tree.cget("displaycolumns")[1], "original_product")


if __name__ == "__main__":
    unittest.main()
