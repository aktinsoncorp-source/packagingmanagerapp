"""Focused regressions for the SKU editor and its Tk save/lock workflow."""
import copy
import json
from pathlib import Path
import sys
import tkinter as tk
from tkinter import ttk
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
    def test_open_unlocks_undefined_rows_without_mutating_catalog(self):
        rows = [product(), product("partial", alternative_variation=""), product("empty", alternative_product="", alternative_variation="", locked=True)]
        original = copy.deepcopy(rows)
        model = ProductsModel({"revision": 7, "rows": rows})
        self.assertEqual([row["locked"] for row in model.rows.values()], [True, False, False])
        self.assertEqual(rows, original)
        self.assertFalse(model.dirty)
        for identity in ("partial", "empty"):
            model.edit(identity, {"alternative_product": "CHANGE"})
            with self.assertRaisesRegex(ValueError, "stay unlocked"):
                model.set_lock(identity, True)
        self.assertEqual(model.unlocked_ids, {"partial", "empty"})

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

    def test_attention_filter_includes_missing_sku_even_with_finished_names(self):
        model = ProductsModel({"rows": [product(), product("no-sku", sku=""), product("partial", alternative_variation="")]})
        self.assertEqual({row["id"] for row in model.sorted_rows(needs_names=True)}, {"no-sku", "partial"})
        self.assertEqual({row["id"] for row in model.sorted_rows(needs_names=True, attention_ids={"rose"})}, {"rose", "no-sku", "partial"})
        before = copy.deepcopy(model.rows["no-sku"])
        model.set_lock("no-sku", False)
        model.edit("no-sku", {"sku": "", "alternative_product": "Standard rose", "alternative_variation": "100 ml"})
        self.assertEqual(model.rows["no-sku"]["sku"], "")
        self.assertEqual(model.rows["no-sku"]["alternative_product"], "STANDARD ROSE")
        self.assertEqual(model.rows["no-sku"]["ecommerce_sources"], before["ecommerce_sources"])

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

    def test_spelled_sizes_sort_numerically_in_both_directions(self):
        for sizes in (["100 GRAMS", "250 GRAMS", "1 KG"], ["250 ML", "1 LIT", "3 LIT", "1 GAL", "4 LIT"],
                      ["20 KG", "1 LIT", "1 GAL"]):
            model = ProductsModel({"rows": [product(str(i), alternative_variation=size)
                                           for i, size in enumerate(reversed(sizes))]})
            before = copy.deepcopy(model.rows)
            self.assertEqual([r["alternative_variation"] for r in model.sorted_rows()], sizes)
            model.toggle_sort("alternative_variation")
            self.assertEqual([r["alternative_variation"] for r in model.sorted_rows()], list(reversed(sizes)))
            self.assertEqual(model.rows, before)

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

    def test_review_auto_unlock_search_and_completed_row_lock_workflow(self):
        self.assertFalse(self.dialog.edit_locked)
        self.assertNotIn("disabled", self.dialog.save_button.state())
        self.assertFalse(self.dialog.set_row_locked("partial", True))
        self.assertFalse(self.dialog.set_all_locks(True))
        self.assertTrue(self.dialog.model.rows["rose"]["locked"])
        self.assertFalse(self.dialog.model.rows["partial"]["locked"])
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
        self.assertTrue(self.dialog.model.rows["rose"]["locked"])
        self.assertFalse(self.dialog.model.rows["partial"]["locked"])
        self.assertFalse(self.dialog.model.dirty)
        self.assertTrue(self.dialog.set_edit_locked(True))
        self.assertEqual(self.requests[-1]["unlocked_ids"], ["partial", "rose"])
        self.assertEqual(self.requests[-1]["review_source_issues"], [])

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

    def test_export_available_while_view_locked(self):
        self.dialog.set_edit_locked(True)
        self.assertNotIn('disabled',self.dialog.export_button.state())
        self.assertIn('disabled',self.dialog.save_button.state())

    def test_sku_name_choice_requires_selection_and_cancel_preserves_edits(self):
        before=copy.deepcopy(self.dialog.model.rows)
        conflict={'sku':'SAME-SKU','choices':[['FIRST','100 ML'],['SECOND','250 ML']]}
        def choose():
            dialog=next(w for w in self.dialog.winfo_children() if isinstance(w,tk.Toplevel))
            frame=next(w for w in dialog.winfo_children() if isinstance(w,ttk.Frame))
            combo=next(w for w in frame.winfo_children() if isinstance(w,ttk.Combobox))
            actions=next(w for w in frame.winfo_children() if isinstance(w,ttk.Frame))
            button=next(w for w in actions.winfo_children() if w.cget('text')=='Keep selected names')
            button.invoke()
            self.assertTrue(dialog.winfo_exists())
            combo.current(1);button.invoke()
        self.root.after(40,choose)
        self.assertEqual(self.dialog.choose_sku_names(conflict),['SECOND','250 ML'])
        def cancel():
            dialog=next(w for w in self.dialog.winfo_children() if isinstance(w,tk.Toplevel))
            dialog.destroy()
        self.root.after(40,cancel)
        self.assertIsNone(self.dialog.choose_sku_names(conflict))
        self.assertEqual(self.dialog.model.rows,before)

    def test_attention_highlight_survives_selection_and_dirty_edits_take_priority(self):
        self.assertIn("attention", self.dialog.tree.item("partial", "tags"))
        self.assertIn("!", self.dialog.tree.item("partial", "text"))
        values = self.dialog.tree.item("partial", "values")
        self.assertEqual(values[1:3], ("[Ecommerce] G. ROSE PERFUME", "[Ecommerce] 100ML"))
        self.dialog.tree.selection_set("partial")
        self.root.update()
        self.assertIn("attentionselected", self.dialog.tree.item("partial", "tags"))
        self.assertFalse(self.dialog.model.rows["partial"]["locked"])
        self.dialog.model.edit("partial", {"alternative_product": "INVENTORY NAME"})
        self.dialog._refresh_rows()
        self.root.update()
        self.assertIn("dirtyselected", self.dialog.tree.item("partial", "tags"))
        self.assertIn("! *", self.dialog.tree.item("partial", "text"))

    def test_attention_only_includes_missing_pdf_sku_for_a_complete_saved_definition(self):
        call = self.dialog.call_engine
        self.dialog._destroy()
        issue = {"id": "rose", "reasons": ["missing_sku"], "sources": [
            {"path": "SPX/orders.pdf", "platform": "SHOPEE", "sku": "", "product": "New ecommerce title", "variation": "100 ml", "order_id": "ORDER-1"}]}
        self.dialog = ProductsDialog(self.root, {"rows": [product(), product("complete-other")]}, call,
                                     attention_only=True, source_issues=[issue])
        self.root.update()
        self.assertEqual(self.dialog.tree.get_children(), ("rose",))
        self.assertFalse(self.dialog.edit_locked)
        self.assertFalse(self.dialog.model.rows["rose"]["locked"])
        self.assertTrue(self.dialog.model.rows["complete-other"]["locked"])
        self.assertIn("SKU missing in PDF", self.dialog._attention_text("rose"))
        self.assertIn(("SHOPEE PDF", "Missing in PDF", "New ecommerce title", "100 ml"), self.dialog._details_for_row(self.dialog.model.rows["rose"]))
        self.assertIn("Generation stays available", self.dialog.info_var.get())
        self.dialog.needs_names.set(False)
        self.dialog._refresh_rows()
        self.root.update()
        self.assertEqual(len(self.dialog.tree.get_children()), 2)
        self.assertIn("attentionselected", self.dialog.tree.item("rose", "tags"))

    def test_current_pdf_attention_scope_hides_unrelated_catalog_deficiencies(self):
        call = self.dialog.call_engine
        self.dialog._destroy()
        self.dialog = ProductsDialog(self.root, {"rows": [product("current", alternative_variation=""), product("old", sku="")]}, call,
                                     attention_only=True, attention_ids=["current"])
        self.root.update()
        self.assertEqual(self.dialog.tree.get_children(), ("current",))
        self.assertIn("! 1 need attention", self.dialog.count_var.get())
        self.dialog.needs_names.set(False)
        self.dialog._refresh_rows()
        self.assertEqual(set(self.dialog.tree.get_children()), {"current", "old"})
        self.assertIn("attention", self.dialog.tree.item("old", "tags"))

    def test_no_sku_product_can_save_standard_fields_and_notify_parent_without_unlocking_others(self):
        call = self.dialog.call_engine
        self.dialog._destroy()
        saved = []
        self.dialog = ProductsDialog(self.root, {"rows": [product("no-sku", sku="", alternative_product="", alternative_variation=""), product()]},
                                     call, attention_only=True, on_catalog_changed=saved.append)
        self.root.update()
        self.assertEqual(self.dialog.tree.get_children(), ("no-sku",))
        self.assertFalse(self.dialog.edit_locked)
        self.assertEqual(self.dialog.edit_button.cget("text"), "Edit selected")
        self.dialog.edit_selected()
        self.root.update()
        edit = next(widget for widget in self.dialog.winfo_children() if isinstance(widget, tk.Toplevel))
        frame = next(widget for widget in edit.winfo_children() if isinstance(widget, ttk.Frame))
        entries = {int(widget.grid_info()["row"]): widget for widget in frame.winfo_children() if isinstance(widget, ttk.Entry)}
        self.assertEqual(entries[1].get(), "")
        self.assertEqual(entries[4].get(), "G. ROSE PERFUME")
        self.assertIn("readonly", entries[4].state())
        entries[2].insert(0, "standard rose")
        entries[3].insert(0, "100 ml")
        actions = next(widget for widget in frame.winfo_children() if isinstance(widget, ttk.Frame))
        next(widget for widget in actions.winfo_children() if widget.cget("text") == "Apply change").invoke()
        self.root.update()
        self.dialog.save_names()
        self.root.update()
        self.assertEqual(len(saved), 1)
        row = self.dialog.model.rows["no-sku"]
        self.assertEqual((row["sku"], row["alternative_product"], row["alternative_variation"]), ("", "STANDARD ROSE", "100 ML"))
        self.assertFalse(row["locked"])
        self.assertTrue(self.dialog.model.rows["rose"]["locked"])
        self.assertEqual(self.dialog.attention_items["no-sku"]["reasons"], ["missing_sku"])
        self.assertFalse(self.dialog.model.dirty)

    def test_pre_export_save_preserves_unlocks_and_hidden_rows(self):
        self.dialog.set_edit_locked(False)
        self.dialog.set_row_locked("rose", False)
        self.dialog.set_row_locked("partial", False)
        self.dialog.search_var.set("no matching row")
        self.dialog.save_names(lock_all=False)
        self.assertEqual(len(self.requests[-1]["rows"]), 2)
        self.assertFalse(self.dialog.model.rows["rose"]["locked"])
        self.assertFalse(self.dialog.model.rows["partial"]["locked"])

    @patch("tkinter.Misc.clipboard_clear")
    @patch("tkinter.Misc.clipboard_append")
    def test_locked_rows_can_copy_fields_and_visible_columns_without_edits(self, append, clear):
        self.dialog.set_edit_locked(True)
        before = copy.deepcopy(self.dialog.model.rows)
        self.assertTrue(before["rose"]["locked"])
        self.dialog.tree.selection_set("rose")
        self.root.update()
        self.assertIn("disabled", self.dialog.edit_button.state())
        for field in ("sku", "alternative_product", "alternative_variation"):
            self.dialog.copy_selected(field)
            append.assert_called_with(before["rose"][field])
        self.root.deiconify()
        self.root.update()
        self.dialog.tree.focus_force()
        self.dialog.tree.event_generate("<Control-c>")
        self.root.update()
        append.assert_called_with("PERF-G-ROSE-W-100ML\tGOLDEA ROSE\t100ML")
        self.dialog.show_ecommerce.set(True)
        self.dialog._display_columns()
        self.dialog.copy_selected("original_product")
        append.assert_called_with("G. ROSE PERFUME")
        self.assertEqual(clear.call_count, 5)
        self.assertEqual(self.dialog.model.rows, before)
        self.assertTrue(self.dialog.edit_locked)
        self.assertFalse(self.dialog.model.dirty)

    @patch("tkinter.Misc.clipboard_clear")
    @patch("tkinter.Misc.clipboard_append")
    def test_locked_row_details_stay_viewable_and_copyable(self, append, clear):
        self.dialog.set_edit_locked(True)
        before = copy.deepcopy(self.dialog.model.rows)
        self.dialog.tree.selection_set("rose")
        self.root.update()
        self.assertNotIn("disabled", self.dialog.details_button.state())
        self.dialog.details_button.invoke()
        self.root.update()
        detail = next(widget for widget in self.dialog.winfo_children() if isinstance(widget, tk.Toplevel))

        def descendants(widget):
            for child in widget.winfo_children():
                yield child
                yield from descendants(child)

        widgets = list(descendants(detail))
        tree = next(widget for widget in widgets if isinstance(widget, ttk.Treeview))
        first = tree.get_children()[0]
        tree.selection_set(first)
        self.root.update()
        copy_button = next(widget for widget in widgets if isinstance(widget, ttk.Button) and widget.cget("text") == "Copy selected listing")
        copy_button.invoke()
        clear.assert_called_once_with()
        append.assert_called_once_with("\t".join(tree.item(first, "values")))
        self.assertIn("G. ROSE PERFUME", append.call_args.args[0])
        self.assertEqual(self.dialog.model.rows, before)
        self.assertTrue(self.dialog.edit_locked)
        self.assertFalse(self.dialog.model.dirty)
        detail.destroy()

    def test_125_percent_small_window_keeps_actions_and_primary_columns_visible(self):
        call = self.dialog.call_engine
        self.dialog._destroy()
        self.root.tk.call("tk", "scaling", float(self.root.tk.call("tk", "scaling"))*1.25)
        self.root.deiconify()
        self.root.update()
        self.dialog = ProductsDialog(self.root, {"revision": 1, "rows": [product(str(index)) for index in range(1278)]}, call)
        self.assertTrue(self.dialog.edit_locked)
        self.dialog.geometry("1100x688")
        self.root.update()
        for name in ("save_button", "edit_button", "add_button", "details_button", "close_button", "export_button", "import_button", "delete_button", "lock_button", "unlock_button", "search_entry", "ecommerce_check", "needs_names_check", "status_label"):
            widget = getattr(self.dialog, name)
            x, y = widget.winfo_rootx()-self.dialog.winfo_rootx(), widget.winfo_rooty()-self.dialog.winfo_rooty()
            self.assertTrue(widget.winfo_viewable(), name)
            self.assertGreaterEqual(x, 0, name)
            self.assertGreaterEqual(y, 0, name)
            self.assertLessEqual(x+widget.winfo_width(), self.dialog.winfo_width(), name)
            self.assertLessEqual(y+widget.winfo_height(), self.dialog.winfo_height(), name)
        self.assertEqual(self.dialog.tree.xview(), (0.0, 1.0))
        self.assertGreaterEqual(self.dialog.tree.winfo_height(), 150)
        self.dialog.attention_scope_ids = {"0"}
        self.dialog.needs_names.set(True)
        self.dialog._refresh_rows()
        self.root.update()
        self.assertGreaterEqual(self.dialog.tree.winfo_height(), 150)
        self.assertLessEqual(self.dialog.status_label.winfo_rooty()+self.dialog.status_label.winfo_height(), self.dialog.winfo_rooty()+self.dialog.winfo_height())
        self.dialog.show_ecommerce.set(True)
        self.dialog._display_columns()
        self.assertEqual(self.dialog.tree.cget("displaycolumns")[1], "original_product")


if __name__ == "__main__":
    unittest.main()
