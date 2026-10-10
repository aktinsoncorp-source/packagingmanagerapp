"""SKU-first product editor for the Python Packing List Manager.

The editor keeps the engine's established payload keys and copies every additive
metadata field. Display labels are intentionally independent of those keys.
"""
from __future__ import annotations

import copy
from decimal import Decimal, InvalidOperation
from functools import cmp_to_key
import json
from pathlib import Path
import queue
import re
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import font as tkfont, messagebox, ttk
import uuid

from PIL import Image, ImageTk

from product_attention import catalog_attention, row_attention_reasons, remap_attention_ids


FIELDS = ("original_product", "original_variation", "alternative_product", "alternative_variation")
LABELS = {"sku": "SKU", "original_product": "Ecommerce Name", "original_variation": "Ecommerce Variation",
          "alternative_product": "Product Name", "alternative_variation": "Size"}
BG, TEXT, MUTED, PRIMARY = "#f1f2ef", "#24292c", "#505b61", "#42484b"
PANEL, LINE, ALT, SELECTED = "#fcfcf9", "#6b767c", "#e6e9e6", "#dce3e4"
MISSING_EXCEL = "Error: No file has been detected\nPlease do not exit excel file before clicking import"


def single_line(value: str, *, alternative: bool = False) -> str:
    if any(ord(c) < 32 or c in "\x7f\x85\u2028\u2029" for c in value):
        raise ValueError("Use one line only. Line breaks are not allowed in product fields.")
    value = value.rstrip(" ").upper()
    if alternative:
        value = re.sub(r"^[^\w]*", "", value, flags=re.UNICODE)
        while value and not value[0].isalnum():
            value = value[1:]
    return value


def size_key(value: str):
    value = value or ""
    match = re.search(r"(?<![\d.])(\d+(?:[.,]\d+)?)\s*(ML|KG|L|G)\b", value, re.I)
    if match:
        try:
            number, unit = Decimal(match[1].replace(",", ".")), match[2].upper()
            return (0, 1 if unit.endswith("L") else 0, number * (1000 if unit in ("L", "KG") else 1), value.casefold())
        except InvalidOperation:
            pass
    return (1, 0, Decimal(0), value.casefold())


class ProductsModel:
    """Editable snapshots; all engine metadata survives UI edits and sorting."""

    def __init__(self, catalog: dict, source_issues=None):
        self.source_issues = copy.deepcopy(source_issues or [])
        self.product_sort, self.size_sort = "alternative_product", "alternative_variation"
        self.product_desc, self.size_desc = False, False
        self.selection_order: list[str] = []
        self.load(catalog, opening=True)

    @staticmethod
    def signature(row):
        return json.dumps(row, sort_keys=True, ensure_ascii=False, separators=(",", ":"))

    def load(self, catalog: dict, *, opening=False):
        self.revision = int(catalog.get("revision", 0))
        self.rows: dict[str, dict] = {}
        for source in catalog.get("rows", []):
            row = copy.deepcopy(source)
            row.setdefault("id", uuid.uuid4().hex)
            for field in FIELDS:
                row.setdefault(field, "")
            if opening:
                row["locked"] = True
            self.rows[row["id"]] = row
        self._source_review_ids = remap_attention_ids(
            {"rows": self.source_issues}, {"rows": list(self.rows.values())},
            [issue["id"] for issue in self.source_issues])
        self.unlocked_ids = self.review_ids()
        for identity in self.unlocked_ids:
            self.rows[identity]["locked"] = False
        self.saved = copy.deepcopy(self.rows)
        self.selection_order = []

    def review_ids(self):
        return {identity for identity, row in self.rows.items()
                if row_attention_reasons(row) or identity in self._source_review_ids}

    @property
    def dirty(self):
        return self.rows != self.saved

    def changed(self, identity):
        return self.rows.get(identity) != self.saved.get(identity)

    def toggle_sort(self, field):
        if field.endswith("variation"):
            self.size_desc = not self.size_desc if field == self.size_sort else False
            self.size_sort = field
        else:
            self.product_desc = not self.product_desc if field == self.product_sort else False
            self.product_sort = field

    def sort_help(self, field):
        size = field.endswith("variation")
        active = field == (self.size_sort if size else self.product_sort)
        desc = active and (self.size_desc if size else self.product_desc)
        current = ("largest to smallest" if desc else "smallest to largest") if size else ("Z-A" if desc else "A-Z")
        following = (("smallest to largest" if desc else "largest to smallest") if size else ("A-Z" if desc else "Z-A")) if active else current
        return f"{LABELS[field]}: {'sorted ' + current + '. ' if active else ''}Click for {following}. The other sort stays unchanged."

    def sorted_rows(self, query="", needs_names=False, attention_ids=()):
        query = query.strip().casefold()

        def matches(row):
            if needs_names and not row_attention_reasons(row) and row.get("id") not in attention_ids:
                return False
            if not query:
                return True
            values = [row.get(field, "") for field in (*FIELDS, "sku")]
            for field in ("aliases", "sku_aliases", "ecommerce_sources"):
                for source in row.get(field, []):
                    values.extend(str(value) for value in source.values())
            return any(query in str(value).casefold() for value in values)

        def compare(a, b):
            a_name, b_name = str(a.get(self.product_sort, "")).casefold(), str(b.get(self.product_sort, "")).casefold()
            if a_name != b_name:
                return ((a_name > b_name) - (a_name < b_name)) * (-1 if self.product_desc else 1)
            a_size, b_size = size_key(a.get(self.size_sort, "")), size_key(b.get(self.size_sort, ""))
            if a_size != b_size:
                return ((a_size > b_size) - (a_size < b_size)) * (-1 if self.size_desc else 1)
            a_key = (a.get("original_product", "").casefold(), size_key(a.get("original_variation", "")), a.get("sku", ""))
            b_key = (b.get("original_product", "").casefold(), size_key(b.get("original_variation", "")), b.get("sku", ""))
            return (a_key > b_key) - (a_key < b_key)

        return sorted((row for row in self.rows.values() if matches(row)), key=cmp_to_key(compare))

    def set_lock(self, identity, locked):
        row = self.rows[identity]
        if locked and (row_attention_reasons(row) or identity in self._source_review_ids):
            raise ValueError("Products needing review stay unlocked. Define their Product Name and Size and correct any missing SKU first.")
        saved = self.saved.get(identity)
        if locked and saved and saved.get("locked"):
            before, after = copy.deepcopy(saved), copy.deepcopy(row)
            before.pop("locked", None)
            after.pop("locked", None)
            if before != after:
                raise ValueError("Save this product before locking its row again.")
        row["locked"] = bool(locked)
        if not locked:
            self.unlocked_ids.add(identity)

    def edit(self, identity, values):
        row = self.rows[identity]
        if row.get("locked"):
            raise ValueError("Unlock this product's row before editing it.")
        new_values = {}
        for key, value in values.items():
            if key not in (*FIELDS, "sku"):
                continue
            new_values[key] = single_line(value, alternative=key.startswith("alternative_"))
        if self.saved.get(identity, {}).get("sku") and new_values.get("sku", row.get("sku")) != row.get("sku"):
            raise ValueError("An existing SKU cannot be changed here. Its ecommerce links must stay with the same product.")
        row.update(new_values)

    def add(self, values):
        sku = single_line(values.get("sku", "")).strip()
        if not sku:
            raise ValueError("Enter the product SKU. You can find it in the ecommerce catalog.")
        identity = uuid.uuid4().hex
        self.rows[identity] = {"id": identity, "sku": sku, "sku_aliases": [], "ecommerce_sources": [],
                               "aliases": [], "locked": False, **dict.fromkeys(FIELDS, "")}
        try:
            self.edit(identity, values)
        except Exception:
            del self.rows[identity]
            raise
        return identity

    def delete(self, identities):
        removed = 0
        for identity in identities:
            if identity in self.rows and not self.rows[identity].get("locked"):
                self.unlocked_ids.add(identity)
                del self.rows[identity]
                removed += 1
        return removed

    def payload(self):
        return copy.deepcopy(list(self.rows.values()))

    @staticmethod
    def ecommerce_details(row):
        result, seen = [], set()

        def append(platform, sku, product, variation):
            key = (platform or "", sku or "", product or "", variation or "")
            if key not in seen:
                seen.add(key)
                result.append(key)

        for source in row.get("ecommerce_sources", []):
            append(source.get("platform"), source.get("sku"), source.get("product"), source.get("variation"))
        known = {(entry[2].casefold(), entry[3].casefold()) for entry in result}
        originals = [row, *row.get("aliases", [])]
        for original in originals:
            product, variation = original.get("original_product", ""), original.get("original_variation", "")
            if (product or variation) and (product.casefold(), variation.casefold()) not in known:
                append("Earlier listing", "", product, variation)
                known.add((product.casefold(), variation.casefold()))
        linked = {(entry[0], entry[1]) for entry in result}
        for alias in row.get("sku_aliases", []):
            if (alias.get("platform", ""), alias.get("sku", "")) not in linked:
                append(alias.get("platform"), alias.get("sku"), "", "")
        return result


class ExcelSession:
    """A private Excel instance held by a hidden STA PowerShell helper."""

    def __init__(self, temporary_parent=None):
        self.folder = Path(tempfile.mkdtemp(prefix="PackingNames-", dir=temporary_parent))
        self.path = self.folder / "Product Names.xlsx"
        self.process = None

    def open(self, visible=True):
        helper = Path(__file__).with_name("products_excel.ps1")
        self.process = subprocess.Popen(
            ["powershell.exe", "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden",
             "-File", str(helper), "-WorkbookPath", str(self.path)] + ([] if visible else ["-HiddenExcel"]),
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, encoding="utf-8", creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return self._response()

    def _response(self):
        line = self.process.stdout.readline() if self.process else ""
        if not line:
            raise RuntimeError(MISSING_EXCEL)
        try:
            reply = json.loads(line.lstrip("\ufeff"))
        except (ValueError, TypeError) as exc:
            raise RuntimeError("Excel could not start. Close any Excel dialog and try Edit in Excel again.") from exc
        if not reply.get("ok"):
            if reply.get("missing"):
                raise RuntimeError(MISSING_EXCEL)
            raise RuntimeError(reply.get("error", "Excel is unavailable."))
        return reply

    def request(self, action):
        if not self.process or self.process.poll() is not None:
            raise RuntimeError(MISSING_EXCEL)
        try:
            self.process.stdin.write(json.dumps({"action": action}) + "\n")
            self.process.stdin.flush()
        except (OSError, ValueError) as exc:
            raise RuntimeError(MISSING_EXCEL) from exc
        return self._response()

    def save(self):
        if not self.path.is_file():
            raise RuntimeError(MISSING_EXCEL)
        self.request("save")

    def close(self):
        if self.process:
            if self.process.poll() is None:
                try:
                    self.request("close")
                except Exception:
                    pass
                try:
                    self.process.wait(timeout=8)
                except subprocess.TimeoutExpired:
                    # Stop only our helper. Never terminate the Excel process or
                    # another workbook that the user opened in the same instance.
                    self.process.terminate()
            for stream in (self.process.stdin, self.process.stdout, self.process.stderr):
                if stream:
                    stream.close()
            self.process = None
        try:
            self.path.unlink(missing_ok=True)
            self.folder.rmdir()
        except OSError:
            pass


def add_column_dividers(tree, color=LINE):
    """Add visible column boundaries without changing Treeview input or data.

    Call after assigning xscrollcommand so the existing scrollbar keeps working.
    The one-pixel child frames move with horizontal scrolling and header drags.
    """
    lines, pending = [], [None]
    previous_scroll = tree.cget("xscrollcommand")

    def refresh():
        pending[0] = None
        if not tree.winfo_exists():
            return
        columns = tuple(str(value) for value in tree.tk.splitlist(tree.cget("displaycolumns")))
        if columns == ("#all",):
            columns = tree.cget("columns")
        columns = list(columns)
        if "tree" in [str(value) for value in tree.tk.splitlist(tree.cget("show"))]:
            columns.insert(0, "#0")
        widths = [int(tree.column(column, "width")) for column in columns]
        count = max(0, len(widths)-1)
        while len(lines) < count:
            lines.append(tk.Frame(tree, background=color, width=1, takefocus=False, borderwidth=0))
        for extra in lines[count:]:
            extra.place_forget()
        offset = int(sum(widths) * tree.xview()[0]) if widths else 0
        x = 1-offset
        for line, width in zip(lines, widths[:-1]):
            x += width
            if 2 < x < tree.winfo_width()-3:
                line.place(x=x, y=2, width=1, height=max(0, tree.winfo_height()-4))
                line.lift()
            else:
                line.place_forget()

    def schedule(_event=None):
        if pending[0] is None and tree.winfo_exists():
            pending[0] = tree.after_idle(refresh)

    def scrolled(first, last):
        if previous_scroll:
            tree.tk.call(*tree.tk.splitlist(previous_scroll), first, last)
        schedule()

    def disposed(_event=None):
        if pending[0] is not None:
            tree.after_cancel(pending[0])
            pending[0] = None

    tree.configure(xscrollcommand=scrolled)
    tree.bind("<Configure>", schedule, add="+")
    tree.bind("<B1-Motion>", schedule, add="+")
    tree.bind("<ButtonRelease-1>", schedule, add="+")
    tree.bind("<Destroy>", disposed, add="+")
    schedule()
    return schedule


class ProductsDialog(tk.Toplevel):
    """Modal editor. call_engine(request, on_success, message, on_error=None)."""

    def __init__(self, parent, catalog, call_engine, *, attention_only=False, source_issues=None, on_catalog_changed=None, attention_ids=None):
        super().__init__(parent)
        self.parent, self.call_engine = parent, call_engine
        self.model = ProductsModel(catalog, source_issues)
        self.attention_only = attention_only
        self.attention_scope_ids = None if attention_ids is None else set(attention_ids)
        self.source_issues = copy.deepcopy(source_issues or [])
        self.on_catalog_changed = on_catalog_changed
        self.attention_items = {}
        self.edit_locked, self.busy = not bool(self.model.review_ids()), False
        self.excel = None
        self._results = queue.Queue()
        self._poll_id = None
        self._closing_after_save = False
        self.title("Products / Define Names")
        self.configure(background=BG)
        self.transient(parent)
        width, height = min(1240, self.winfo_screenwidth() - 60), min(790, self.winfo_screenheight() - 90)
        self.geometry(f"{width}x{height}+{max(10, (self.winfo_screenwidth()-width)//2)}+{max(10, (self.winfo_screenheight()-height)//2)}")
        self.minsize(940, 620)
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.closed_icon, self.open_icon = self._lock_icon(False), self._lock_icon(True)
        self.disabled_closed_icon = self._lock_icon(False, disabled=True)
        self.disabled_open_icon = self._lock_icon(True, disabled=True)
        self.both_icons = self._both_lock_icons()
        self._build()
        self._refresh()
        if attention_only and self.tree.get_children():
            self.tree.selection_set(self.tree.get_children()[0])
        self.grab_set()
        self.search_entry.focus_set()
        self._poll_id = self.after(70, self._poll_workers)

    def _lock_icon(self, opened, disabled=False):
        name = "lock-open" if opened else "lock-closed"
        if disabled:
            name += "-disabled"
        with Image.open(Path(__file__).resolve().parent / "assets" / "ui" / f"{name}.png") as icon:
            source = icon.convert("RGBA")
        return ImageTk.PhotoImage(source, master=self)

    def _both_lock_icons(self):
        source = Image.new("RGBA", (84, 34))
        for name, x in (("closed", 2), ("open", 48)):
            with Image.open(Path(__file__).resolve().parent / "assets" / "ui" / f"lock-{name}-header.png") as icon:
                source.alpha_composite(icon.convert("RGBA"), (x, 0))
        return ImageTk.PhotoImage(source, master=self)

    def _build(self):
        style = ttk.Style(self)
        style.configure("Products.TFrame", background=BG)
        style.configure("Products.TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 12))
        style.configure("Products.Muted.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 11))
        style.configure("Products.State.TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 11, "bold"))
        style.configure("Products.Title.TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 22, "bold"))
        style.configure("Products.Treeview", font=("Segoe UI", 12), rowheight=46, background=PANEL, fieldbackground=PANEL, foreground=TEXT,
                        borderwidth=2, relief="solid", bordercolor=LINE, lightcolor=LINE, darkcolor=LINE)
        # Paint selection with row tags so selecting a changed row keeps its
        # yellow highlight instead of covering it with the theme's blue fill.
        style.map("Products.Treeview", background=[("disabled", ALT)], foreground=[("disabled", "#646e73")])
        style.configure("Products.Treeview.Heading", font=("Segoe UI", 11, "bold"), padding=(8, 10), background=PRIMARY, foreground=PANEL,
                        borderwidth=2, relief="flat", bordercolor=LINE, lightcolor=LINE, darkcolor=LINE)
        style.map("Products.Treeview.Heading", background=[("active", "#30383d")])
        style.configure("Products.TButton", padding=(12, 9), font=("Segoe UI", 11), foreground=TEXT, background=PANEL,
                        borderwidth=1, relief="solid", bordercolor=LINE, lightcolor=LINE, darkcolor=LINE)
        style.map("Products.TButton", background=[("disabled", ALT), ("active", ALT)], foreground=[("disabled", "#646e73")])
        style.configure("Products.TEntry", fieldbackground=PANEL, foreground=TEXT, bordercolor=LINE, lightcolor=LINE, darkcolor=LINE)
        style.configure("Products.TCheckbutton", background=BG, foreground=TEXT, font=("Segoe UI", 11))
        style.map("Products.TCheckbutton", background=[("active", BG)], foreground=[("disabled", "#646e73")])
        style.configure("Products.Primary.TButton", padding=(14, 10), font=("Segoe UI", 12, "bold"), foreground=PANEL, background=PRIMARY,
                        borderwidth=1, relief="solid", bordercolor=PRIMARY, lightcolor=PRIMARY, darkcolor=PRIMARY)
        style.map("Products.Primary.TButton", background=[("disabled", ALT), ("active", "#30383d")], foreground=[("disabled", "#646e73")])
        style.configure("Loading.Horizontal.TProgressbar", background=PRIMARY, troughcolor="#cbd1d0",
                        bordercolor=LINE, lightcolor=PRIMARY, darkcolor=PRIMARY, thickness=12)
        outer = ttk.Frame(self, padding=20, style="Products.TFrame")
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(3, weight=1)
        heading = ttk.Frame(outer, style="Products.TFrame")
        heading.grid(row=0, column=0, sticky="ew")
        ttk.Label(heading, text="Products", style="Products.Title.TLabel").pack(anchor="w")
        tk.Frame(heading, height=4, background=PRIMARY).pack(fill="x", pady=(5, 3))
        self.info_var = tk.StringVar()
        self.info_label = ttk.Label(outer, textvariable=self.info_var, style="Products.TLabel", wraplength=1120)
        self.info_label.grid(row=1, column=0, sticky="ew", pady=(3, 16))
        toolbar = ttk.Frame(outer, style="Products.TFrame")
        toolbar.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        toolbar.columnconfigure(1, weight=1)
        ttk.Label(toolbar, text="Find product", style="Products.TLabel").grid(row=0, column=0, padx=(0, 10))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(toolbar, textvariable=self.search_var, width=25, font=("Segoe UI", 12), style="Products.TEntry")
        self.search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 16))
        self.search_var.trace_add("write", lambda *_: self._refresh_rows())
        self.show_ecommerce = tk.BooleanVar(value=False)
        self.ecommerce_check = ttk.Checkbutton(toolbar, text="Show Ecommerce columns", variable=self.show_ecommerce, command=self._display_columns, style="Products.TCheckbutton")
        self.ecommerce_check.grid(row=0, column=2, padx=(0, 14))
        self.lock_button = ttk.Button(toolbar, image=(self.closed_icon, "disabled", self.disabled_closed_icon), width=3, command=lambda: self.set_edit_locked(True), style="Products.TButton")
        self.unlock_button = ttk.Button(toolbar, image=(self.open_icon, "disabled", self.disabled_open_icon), width=3, command=lambda: self.set_edit_locked(False), style="Products.TButton")
        self.edit_state_var = tk.StringVar(value="Editing locked")
        ttk.Label(toolbar, textvariable=self.edit_state_var, style="Products.State.TLabel").grid(row=0, column=3, padx=(0, 10))
        self.lock_button.grid(row=0, column=4, padx=(0, 4))
        self.unlock_button.grid(row=0, column=5)
        self._tip(self.lock_button, "Lock product editing")
        self._tip(self.unlock_button, "Unlock editing. Then unlock a product row to change it.")
        table = ttk.Frame(outer, style="Products.TFrame")
        table.grid(row=3, column=0, sticky="nsew")
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(table, columns=("sku", "alternative_product", "alternative_variation", "details", "original_product", "original_variation"), selectmode="extended", style="Products.Treeview")
        self.tree.heading("#0", text="", image=self.both_icons, anchor="center")
        self.tree.column("#0", width=106, minwidth=106, stretch=False, anchor="w")
        for field, width in (("sku", 310), ("alternative_product", 350), ("alternative_variation", 165), ("original_product", 440), ("original_variation", 240)):
            self.tree.heading(field, text=LABELS[field], anchor="w", command=lambda key=field: self.sort(key))
            self.tree.column(field, width=width, minwidth=140 if field.endswith("variation") else 220, stretch=True)
        self.tree.heading("details", text="Ecommerce Details", anchor="w")
        self.tree.column("details", width=180, minwidth=180, stretch=False)
        self.tree.tag_configure("odd", background=ALT)
        self.tree.tag_configure("dirty", background="#ffe599", foreground="#162334")
        self.tree.tag_configure("selected", background=SELECTED, foreground=TEXT)
        self.tree.tag_configure("dirtyselected", background="#ffd966", foreground="#162334")
        self.tree.tag_configure("attention", background="#f6d6ae", foreground="#39250c")
        self.tree.tag_configure("attentionselected", background="#e8bb87", foreground="#39250c")
        self.tree.grid(row=0, column=0, sticky="nsew")
        vertical = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        horizontal = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        self._update_dividers = add_column_dividers(self.tree)
        self.tree.bind("<Button-1>", self._tree_click)
        self.tree.bind("<Double-1>", self._double_click)
        self.tree.bind("<<TreeviewSelect>>", self._selected)
        self.tree.bind("<Return>", lambda _e: self.edit_selected())
        self.tree.bind("<F2>", lambda _e: self.edit_selected())
        self.tree.bind("<Control-c>", lambda _e: self.copy_selected())
        self.tree.bind("<Control-C>", lambda _e: self.copy_selected())
        self.tree.bind("<Button-3>", self._copy_menu)
        self.tree.bind("<Configure>", lambda _e: self._fit_columns(), add="+")
        self._tip(self.tree, self._tree_tip)
        under = ttk.Frame(outer, style="Products.TFrame")
        under.grid(row=4, column=0, sticky="ew", pady=(10, 5))
        under.columnconfigure(0, weight=1)
        self.count_var = tk.StringVar()
        ttk.Label(under, textvariable=self.count_var, style="Products.Muted.TLabel").grid(row=0, column=0, sticky="w")
        self.needs_names = tk.BooleanVar(value=self.attention_only)
        self.needs_names_check = ttk.Checkbutton(under, text="Needs SKU, Name or Size", variable=self.needs_names, command=self._refresh_rows, style="Products.TCheckbutton")
        self.needs_names_check.grid(row=0, column=1, sticky="e")
        actions = ttk.Frame(outer, style="Products.TFrame")
        actions.grid(row=5, column=0, sticky="ew", pady=(8, 8))
        self.save_button = ttk.Button(actions, text="Save names", command=self.save_names, style="Products.Primary.TButton")
        self.edit_button = ttk.Button(actions, text="Edit selected", command=lambda: self.edit_selected(unlock=True), style="Products.TButton")
        self.add_button = ttk.Button(actions, text="Add product", command=self.add_product, style="Products.TButton")
        self.details_button = ttk.Button(actions, text="View Ecommerce Details", command=self.show_details, style="Products.TButton")
        self.save_button.pack(side="left", padx=(0, 8))
        self.edit_button.pack(side="left", padx=(0, 8))
        self.add_button.pack(side="left", padx=(0, 8))
        self.details_button.pack(side="left", padx=(0, 8))
        self.close_button = ttk.Button(actions, text="Close", command=self.close, style="Products.TButton")
        self.close_button.pack(side="right")
        advanced = ttk.Frame(outer, style="Products.TFrame")
        advanced.grid(row=6, column=0, sticky="ew", pady=(0, 10))
        self.export_button = ttk.Button(advanced, text="Edit in Excel", command=self.export_excel, style="Products.TButton")
        self.import_button = ttk.Button(advanced, text="Apply Excel edits", command=self.import_excel, style="Products.TButton")
        self.delete_button = ttk.Button(advanced, text="Delete selected", command=self.delete_selected, style="Products.TButton")
        self.export_button.pack(side="left", padx=(0, 8))
        self.import_button.pack(side="left", padx=(0, 8))
        self.delete_button.pack(side="left", padx=(0, 14))
        ttk.Label(advanced, text="* Unsaved changes (yellow rows)", style="Products.Muted.TLabel").pack(side="left")
        self.status_var = tk.StringVar(value=("Review rows are unlocked. Select one and choose Edit selected. Completed products stay locked."
                                            if not self.edit_locked else "Viewing products. Ctrl+C or right-click copies names. Use the open padlock above to edit."))
        self.status_label = ttk.Label(outer, textvariable=self.status_var, style="Products.Muted.TLabel", wraplength=1120)
        self.status_label.grid(row=7, column=0, sticky="ew")
        self.progress = ttk.Progressbar(outer, mode="indeterminate", style="Loading.Horizontal.TProgressbar")
        self.progress.grid(row=8, column=0, sticky="ew", pady=(5, 0))
        self.progress.grid_remove()
        self.bind("<Configure>", self._resize_notes)
        self._display_columns()

    def _resize_notes(self, event):
        if event.widget is self:
            width = max(850, self.winfo_width()-60)
            self.status_label.configure(wraplength=width)
            self.info_label.configure(wraplength=width)

    def _attention_text(self, identity):
        row = self.model.rows[identity]
        reasons = self.attention_items.get(identity, {}).get("reasons", [])
        sku_issue = "SKU missing or needs correction" if "missing_sku" in row_attention_reasons(row) else "SKU missing in PDF"
        labels = {"missing_sku": sku_issue,
                  "missing_product_name": "Product Name missing", "missing_size": "Size missing"}
        return "; ".join(labels[reason] for reason in reasons if reason in labels)

    def _details_for_row(self, row):
        details = self.model.ecommerce_details(row)
        for source in self.attention_items.get(row["id"], {}).get("sources", []):
            if not source.get("path"):
                continue
            item = (source.get("platform", "") + " PDF", source.get("sku") or "Missing in PDF",
                    source.get("product", ""), source.get("variation", ""))
            if item not in details:
                details.append(item)
        return details

    def _row_tag(self, identity, number, selected=False):
        if self.model.changed(identity):
            return "dirtyselected" if selected else "dirty"
        if identity in self.attention_items:
            return "attentionselected" if selected else "attention"
        return "selected" if selected else ("odd" if number % 2 == 0 else "")

    def _tip(self, widget, text):
        state = {"window": None, "timer": None}

        def hide(_event=None):
            if state["timer"]:
                self.after_cancel(state["timer"])
                state["timer"] = None
            if state["window"]:
                state["window"].destroy()
                state["window"] = None

        def motion(event):
            hide()
            content = text(event) if callable(text) else text
            if not content:
                return

            def show():
                state["timer"] = None
                tip = tk.Toplevel(self)
                tip.overrideredirect(True)
                tip.geometry(f"+{event.x_root+12}+{event.y_root+18}")
                tk.Label(tip, text=content, background="#fff8d8", foreground="#162334", padx=8, pady=6, wraplength=450, justify="left", font=("Segoe UI", 11)).pack()
                state["window"] = tip
            state["timer"] = self.after(650, show)
        widget.bind("<Motion>", motion, add="+")
        widget.bind("<Leave>", hide, add="+")
        widget.bind("<ButtonPress>", hide, add="+")
        widget.bind("<Destroy>", hide, add="+")

    def _tree_tip(self, event):
        column = self.tree.identify_column(event.x)
        region = self.tree.identify_region(event.x, event.y)
        if column == "#0":
            return "Lock all rows (left padlock) / unlock all rows (right padlock)" if region == "heading" else "Click this row's padlock to lock or unlock it."
        field = self._column_field(column)
        if region == "heading" and field in LABELS:
            return self.model.sort_help(field)
        identity = self.tree.identify_row(event.y)
        if identity and identity in self.attention_items:
            if field == "details":
                return self._attention_text(identity) + ". Click to view ecommerce names, variations and seller SKUs."
            if field in ("sku", "alternative_product", "alternative_variation"):
                return self._attention_text(identity) + ". Blank standard fields print the ecommerce text. Add missing SKUs on your ecommerce website."
        if identity and field == "sku" and not self.model.rows[identity].get("sku"):
            return "This earlier definition has no confirmed SKU. Its saved ecommerce names remain available for matching."
        if identity and field == "details":
            return "Click to view linked seller SKUs, ecommerce names and variations."
        if identity and field in LABELS:
            value = self.model.rows[identity].get(field, "")
            if value:
                return LABELS[field] + ": " + value
        return None

    def _column_field(self, display_column):
        if display_column == "#0":
            return "#0"
        columns = self.tree.cget("displaycolumns")
        try:
            return columns[int(display_column[1:])-1]
        except (ValueError, IndexError):
            return None

    def _display_columns(self):
        columns = ("sku", "alternative_product", "alternative_variation", "details")
        if self.show_ecommerce.get():
            # Put the requested fields into view immediately. The inventory
            # columns remain available to the right in this expanded view.
            columns = ("sku", "original_product", "original_variation", "alternative_product", "alternative_variation", "details")
        self.tree.configure(displaycolumns=columns)
        self.tree.xview_moveto(0)
        self._fit_columns()
        if hasattr(self, "_update_dividers"):
            self._update_dividers()

    def _fit_columns(self):
        if self.show_ecommerce.get() or self.tree.winfo_width() < 200:
            return
        heading = tkfont.Font(self, font=("Segoe UI", 11, "bold"))
        body = tkfont.Font(self, font=("Segoe UI", 12))
        row_number_width = max(106, body.measure("  " + str(max(1, len(self.model.rows))) + " ! *")+58)
        self.tree.column("#0", width=row_number_width, minwidth=row_number_width)
        details = max(180, heading.measure("Ecommerce Details")+20, body.measure("View details (999)")+18)
        size = max(145, heading.measure("Size ↑ Smallest")+18)
        product_min = max(220, heading.measure("Product Name ↑ A-Z")+20)
        space = max(220+product_min, self.tree.winfo_width()-row_number_width-details-size-5)
        sku = max(220, min(int(space*.46), space-product_min))
        for field, width in (("sku", sku), ("alternative_product", space-sku), ("alternative_variation", size), ("details", details)):
            self.tree.column(field, width=int(width))
        if hasattr(self, "_update_dividers"):
            self._update_dividers()

    def _refresh(self):
        self._refresh_rows()
        self._state()

    def _refresh_rows(self):
        if not hasattr(self, "tree"):
            return
        selected = set(self.tree.selection())
        yview = self.tree.yview()[0] if self.tree.get_children() else 0
        children = self.tree.get_children()
        if children:
            self.tree.delete(*children)
        needs_names = self.needs_names.get() if hasattr(self, "needs_names") else False
        attention = catalog_attention({"rows": list(self.model.rows.values())}, self.source_issues)
        self.attention_items = {item["id"]: item for item in attention["items"] if item["id"] in self.model.rows}
        rows = self.model.sorted_rows(self.search_var.get(), needs_names, self.attention_items)
        if needs_names and self.attention_scope_ids is not None:
            rows = [row for row in rows if row["id"] in self.attention_scope_ids]
        scope = " for these PDFs" if self.attention_scope_ids is not None else " below"
        self.info_var.set(f"! Define Product Name and Size{scope}. Review rows are automatically unlocked.\nGeneration stays available. Add missing seller SKUs on your ecommerce website."
                          if needs_names else "SKU keeps your Product Name and Size linked when ecommerce text changes.\n! Amber rows need SKU, Product Name or Size. Blank fields print ecommerce text.")
        for number, row in enumerate(rows, 1):
            identity = row["id"]
            details_count = len(self._details_for_row(row))
            product_name = row.get("alternative_product", "").strip() or "[Ecommerce] " + (row.get("original_product") or "Name not supplied")
            size = row.get("alternative_variation", "").strip() or "[Ecommerce] " + (row.get("original_variation") or "Variation not supplied")
            values = [row.get("sku") or "! SKU missing", product_name, size,
                      f"View details ({details_count})", row.get("original_product", ""), row.get("original_variation", "")]
            changed = self.model.changed(identity)
            tag = self._row_tag(identity, number)
            marker = (" !" if identity in self.attention_items else "") + (" *" if changed else "")
            self.tree.insert("", "end", iid=identity, text=f"  {number}{marker}", image=self.closed_icon if row.get("locked") else self.open_icon, values=values, tags=(tag,) if tag else ())
        self.tree.selection_set([identity for identity in selected if self.tree.exists(identity)])
        self.tree.yview_moveto(yview)
        for field in LABELS:
            size = field.endswith("variation")
            active = field == (self.model.size_sort if size else self.model.product_sort)
            desc = self.model.size_desc if size else self.model.product_desc
            suffix = (("Largest" if desc else "Smallest") if size else ("Z-A" if desc else "A-Z")) if active else ""
            arrow = " ↓ " if desc else " ↑ "
            self.tree.heading(field, text=LABELS[field] + (arrow + suffix if active else ""))
        if hasattr(self, "count_var"):
            locked = sum(bool(row.get("locked")) for row in self.model.rows.values())
            text = f"{len(rows):,} / {len(self.model.rows):,} products  ·  {locked:,} locked"
            attention_count = len(self.attention_items.keys() & self.attention_scope_ids) if needs_names and self.attention_scope_ids is not None else len(self.attention_items)
            if attention_count:
                text += f"  ·  ! {attention_count:,} need attention"
            self.count_var.set(text)
        self._state()

    def _state(self):
        if not hasattr(self, "save_button"):
            return
        pending = self.excel is not None
        editable = not self.busy and not pending and not self.edit_locked
        selected = self.tree.selection()
        selected_unlocked = any(not self.model.rows[identity].get("locked") for identity in selected if identity in self.model.rows)
        enabled = {
            self.save_button: editable, self.add_button: editable,
            self.edit_button: editable and len(selected) == 1,
            self.delete_button: editable and selected_unlocked,
            self.export_button: editable, self.import_button: not self.busy and pending and not self.edit_locked,
            self.lock_button: not self.busy and not pending and not self.edit_locked,
            self.unlock_button: not self.busy and not pending and self.edit_locked,
            self.details_button: not self.busy and len(selected) == 1,
            self.close_button: not self.busy,
        }
        for widget, active in enabled.items():
            widget.state(["!disabled"] if active else ["disabled"])
        self.edit_button.configure(text="Unlock & edit" if len(selected) == 1 and not selected_unlocked else "Edit selected")
        self.edit_state_var.set("Editing locked" if self.edit_locked else "Editing unlocked")
        self.configure(cursor="watch" if self.busy else "")
        if self.busy:
            if not self.progress.winfo_manager():
                self.progress.grid()
                self.progress.start(18)
        else:
            self.progress.stop()
            self.progress.grid_remove()

    def _selected(self, _event=None):
        selection = self.tree.selection()
        self.model.selection_order = [identity for identity in self.model.selection_order if identity in selection]
        self.model.selection_order.extend(identity for identity in selection if identity not in self.model.selection_order)
        for number, identity in enumerate(self.tree.get_children(), 1):
            tag = self._row_tag(identity, number, identity in selection)
            self.tree.item(identity, tags=(tag,) if tag else ())
        self._state()

    def _tree_click(self, event):
        if self.busy:
            return "break"
        column, region = self.tree.identify_column(event.x), self.tree.identify_region(event.x, event.y)
        if column == "#0" and region == "heading":
            self.set_all_locks(event.x < 53)
            return "break"
        identity = self.tree.identify_row(event.y)
        if identity and column == "#0" and event.x < 58:
            self.set_row_locked(identity, not self.model.rows[identity].get("locked"))
            return "break"
        if identity and self._column_field(column) == "details":
            self.tree.selection_set(identity)
            self.show_details()
            return "break"

    def _double_click(self, event):
        if self.tree.identify_region(event.x, event.y) not in ("tree", "cell"):
            return
        if self._column_field(self.tree.identify_column(event.x)) == "details":
            return "break"
        self.edit_selected()
        return "break"

    def copy_selected(self, field=None):
        fields = (field,) if field else tuple(key for key in self.tree.cget("displaycolumns") if key != "details")
        selected = set(self.tree.selection())
        rows = [self.model.rows[identity] for identity in self.tree.get_children() if identity in selected]
        if rows:
            self.clipboard_clear()
            self.clipboard_append("\n".join("\t".join(str(row.get(key, "")) for key in fields) for row in rows))
            self.status_var.set(f"Copied {LABELS[field] if field else 'selected rows'} to the clipboard.")
        return "break"

    def _copy_menu(self, event):
        identity = self.tree.identify_row(event.y)
        if not identity:
            return
        if identity not in self.tree.selection():
            self.tree.selection_set(identity)
        self.tree.focus_set()
        menu = getattr(self, "copy_context_menu", None)
        if menu is None:
            menu = self.copy_context_menu = tk.Menu(self, tearoff=False)
        else:
            menu.delete(0, "end")
        for field in self.tree.cget("displaycolumns"):
            if field in LABELS:
                menu.add_command(label="Copy " + LABELS[field], command=lambda key=field: self.copy_selected(key))
        menu.add_separator()
        menu.add_command(label="Copy selected rows", accelerator="Ctrl+C", command=self.copy_selected)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def sort(self, field):
        if self.busy:
            return
        self.model.toggle_sort(field)
        self._refresh_rows()
        self.status_var.set(self.model.sort_help(field))

    def set_edit_locked(self, locked):
        if self.busy or self.excel:
            return False
        if locked and self.model.dirty:
            self.status_var.set("Save your changes before locking product editing.")
            return False
        self.edit_locked = locked
        self._state()
        self.status_var.set("Viewing products. Use the open padlock above to edit." if locked else "Editing unlocked. Review rows are ready to edit. Save names locks completed products.")
        return True

    def set_row_locked(self, identity, locked):
        if self.busy or self.excel or self.edit_locked:
            self.status_var.set("Use the open padlock above to unlock editing first.")
            return False
        try:
            self.model.set_lock(identity, locked)
        except ValueError as exc:
            self.status_var.set(str(exc))
            return False
        self._refresh_rows()
        self.status_var.set("Row lock changed. Save names locks completed products; review rows stay unlocked.")
        return True

    def set_all_locks(self, locked):
        if self.busy or self.excel or self.edit_locked:
            self.status_var.set("Use the open padlock above to unlock editing first.")
            return False
        failures = []
        for identity in self.model.rows:
            try:
                self.model.set_lock(identity, locked)
            except ValueError as exc:
                failures.append(str(exc))
        self._refresh_rows()
        self.status_var.set(failures[0] if failures else ("All rows locked." if locked else "All rows unlocked. You can edit them here or in Excel."))
        return not failures

    def edit_selected(self, unlock=False):
        selected = self.tree.selection()
        if self.busy or self.excel or self.edit_locked or len(selected) != 1:
            if self.edit_locked:
                self.status_var.set("Use the open padlock above to unlock editing first.")
            return
        identity = selected[0]
        if self.model.rows[identity].get("locked"):
            if not unlock:
                self.status_var.set("Choose Unlock & edit, or click this row's padlock, to change the product.")
                return
            self.set_row_locked(identity, False)
        self._edit_dialog(identity)

    def add_product(self):
        if not self.busy and not self.excel and not self.edit_locked:
            self._edit_dialog(None)

    def _edit_dialog(self, identity):
        row = self.model.rows[identity] if identity else {}
        dialog = tk.Toplevel(self)
        dialog.title("Edit product" if identity else "Add product")
        dialog.configure(background=BG)
        dialog.transient(self)
        dialog.resizable(False, False)
        frame = ttk.Frame(dialog, padding=24, style="Products.TFrame")
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)
        ttk.Label(frame, text="Edit product" if identity else "Add product", style="Products.Title.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))
        fields = ["sku", "alternative_product", "alternative_variation"]
        show_source_evidence = bool(identity and identity in self.attention_items)
        if self.show_ecommerce.get() or show_source_evidence:
            fields += ["original_product", "original_variation"]
        entries = {}
        focus = None
        for number, field in enumerate(fields, 1):
            ttk.Label(frame, text=LABELS[field], style="Products.TLabel").grid(row=number, column=0, sticky="w", padx=(0, 18), pady=7)
            entry = ttk.Entry(frame, width=54, font=("Segoe UI", 12))
            entry.insert(0, row.get(field, ""))
            entry.grid(row=number, column=1, sticky="ew", pady=7)
            if (field == "sku" and identity and self.model.saved.get(identity, {}).get("sku")) or (show_source_evidence and field.startswith("original_") and not self.show_ecommerce.get()):
                entry.configure(state="readonly")
            elif focus is None:
                focus = entry
            entries[field] = entry
        hint = "Saved SKUs keep their existing ecommerce links. Blank Product Name or Size uses the ecommerce text."
        if identity and not row.get("sku"):
            hint = "You may leave SKU blank and save Product Name and Size for this ecommerce listing. Add its SKU on your ecommerce website to keep future name changes linked. Blank standard fields print the ecommerce text."
            focus = entries["alternative_product"] if not row.get("alternative_product", "").strip() else entries["alternative_variation"]
        elif show_source_evidence:
            hint = self._attention_text(identity) + ". " + hint
        ttk.Label(frame, text=hint, wraplength=600, style="Products.Muted.TLabel").grid(row=len(fields)+1, column=0, columnspan=2, sticky="w", pady=(12, 4))
        error = tk.StringVar()
        ttk.Label(frame, textvariable=error, wraplength=600, style="Products.Muted.TLabel").grid(row=len(fields)+2, column=0, columnspan=2, sticky="w", pady=(0, 8))
        actions = ttk.Frame(frame, style="Products.TFrame")
        actions.grid(row=len(fields)+3, column=0, columnspan=2, sticky="e")

        def done():
            values = {key: entry.get() for key, entry in entries.items() if "readonly" not in entry.state()}
            try:
                if identity:
                    self.model.edit(identity, values)
                    edited_id = identity
                else:
                    edited_id = self.model.add(values)
            except ValueError as exc:
                error.set(str(exc))
                return
            dialog.destroy()
            self.grab_set()
            self._refresh_rows()
            if self.tree.exists(edited_id):
                self.tree.selection_set(edited_id)
                self.tree.see(edited_id)
            self.status_var.set("Product updated in this edit. Click Save names to use it on future packing lists.")

        def cancel():
            dialog.destroy()
            self.grab_set()

        ttk.Button(actions, text="Cancel", command=cancel, style="Products.TButton").pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Apply change", command=done, style="Products.Primary.TButton").pack(side="left")
        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.bind("<Escape>", lambda _e: cancel())
        dialog.bind("<Return>", lambda _e: done())
        dialog.grab_set()
        if focus:
            focus.focus_set()
            focus.selection_range(0, "end")

    def delete_selected(self):
        if self.busy or self.excel or self.edit_locked:
            return
        count = self.model.delete(self.tree.selection())
        self._refresh_rows()
        self.status_var.set(f"{count} unlocked product(s) removed from this edit. Save names to apply. A later PDF scan may add them again.")

    def show_details(self):
        selected = self.tree.selection()
        if self.busy or len(selected) != 1:
            return
        row = self.model.rows[selected[0]]
        dialog = tk.Toplevel(self)
        dialog.title("Ecommerce Details — " + (row.get("sku") or "SKU needs review"))
        dialog.configure(background=BG)
        dialog.transient(self)
        dialog.geometry(f"{min(1160, self.winfo_screenwidth()-80)}x540")
        dialog.minsize(850, 420)
        frame = ttk.Frame(dialog, padding=20, style="Products.TFrame")
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)
        ttk.Label(frame, text=row.get("sku") or "SKU needs review", style="Products.Title.TLabel").grid(row=0, column=0, sticky="w")
        title = "  ·  ".join(value for value in (row.get("alternative_product"), row.get("alternative_variation")) if value)
        ttk.Label(frame, text=title or "Product Name and Size have not been set.", style="Products.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 8))
        note = "These seller SKUs and earlier ecommerce listings use this product definition. Excel includes the linked originals."
        if row["id"] in self.attention_items:
            note = "! " + self._attention_text(row["id"]) + ". Blank standard fields print ecommerce text."
            if "missing_sku" in self.attention_items[row["id"]]["reasons"]:
                note += " Add the missing SKU on your ecommerce website, then use a new PDF. You may save Product Name and Size here now."
        if row.get("review_notes"):
            note += "\nReview: " + "; ".join(row["review_notes"])
        ttk.Label(frame, text=note, wraplength=1050, style="Products.Muted.TLabel").grid(row=2, column=0, sticky="ew", pady=(0, 12))
        body = ttk.Frame(frame, style="Products.TFrame")
        body.grid(row=3, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)
        tree = ttk.Treeview(body, columns=("platform", "sku", "product", "variation"), show="headings", selectmode="browse", style="Products.Treeview")
        details = self._details_for_row(row)
        direction = {}

        def show_sorted(field):
            descending = not direction.get(field, True)
            direction.clear()
            direction[field] = descending
            index = ("platform", "sku", "product", "variation").index(field)
            details.sort(key=lambda data: size_key(data[index]) if field == "variation" else data[index].casefold(), reverse=descending)
            populate()

        for field, label, width in (("platform", "Platform", 120), ("sku", "Seller SKU", 240), ("product", "Ecommerce Name", 470), ("variation", "Ecommerce Variation", 220)):
            tree.heading(field, text=label, command=lambda key=field: show_sorted(key))
            tree.column(field, width=width, minwidth=100)
        tree.tag_configure("odd", background=ALT)

        def populate():
            if tree.get_children():
                tree.delete(*tree.get_children())
            for index, detail in enumerate(details):
                tree.insert("", "end", values=detail, tags=("odd",) if index % 2 else ())
        populate()
        tree.grid(row=0, column=0, sticky="nsew")
        vertical = ttk.Scrollbar(body, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(body, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        add_column_dividers(tree)
        full_text = tk.StringVar(value="Select a listing to read its full name and variation. Ctrl+C copies the listing.")
        tree.tag_configure("selected", background=SELECTED, foreground=TEXT)

        def select_detail(_event=None):
            selected = tree.selection()
            for index, item in enumerate(tree.get_children()):
                tree.item(item, tags=("selected",) if item in selected else (("odd",) if index % 2 else ()))
            if selected:
                values = tree.item(selected[0], "values")
                full_text.set("\n".join((str(values[2]), str(values[3]))))
        tree.bind("<<TreeviewSelect>>", select_detail)
        def copy_detail(_event=None):
            selected = tree.selection()
            if selected:
                dialog.clipboard_clear()
                dialog.clipboard_append("\t".join(str(value) for value in tree.item(selected[0], "values")))
            return "break"
        tree.bind("<Control-c>", copy_detail)
        tree.bind("<Control-C>", copy_detail)
        ttk.Label(frame, textvariable=full_text, wraplength=1000, style="Products.Muted.TLabel").grid(row=4, column=0, sticky="ew", pady=10)

        def done():
            dialog.destroy()
            self.grab_set()
        detail_actions = ttk.Frame(frame, style="Products.TFrame")
        detail_actions.grid(row=5, column=0, sticky="ew")
        ttk.Button(detail_actions, text="Copy selected listing", command=copy_detail, style="Products.TButton").pack(side="left")
        ttk.Button(detail_actions, text="Close details", command=done, style="Products.TButton").pack(side="right")
        dialog.protocol("WM_DELETE_WINDOW", done)
        dialog.bind("<Escape>", lambda _e: done())
        dialog.grab_set()

    def _engine(self, request, success, progress, on_error=None):
        self.busy = True
        self.status_var.set(progress)
        self._state()

        def complete(result):
            self.busy = False
            self._state()
            success(result)

        def failed(error):
            self.busy = False
            self._state()
            (on_error or self._error)(error)
        self.call_engine(request, complete, progress, on_error=failed)

    def _error(self, error):
        text = str(error)
        self.status_var.set(text)
        messagebox.showwarning("Check products", text, parent=self)

    def save_names(self, lock_all=True, on_saved=None, prepared_rows=None, priority=None):
        if self.busy or self.edit_locked:
            return False
        if self.excel is not None and prepared_rows is None:
            return False
        rows = self.model.payload() if prepared_rows is None else prepared_rows
        priority = list(self.model.selection_order) if priority is None else priority

        def prepared(preview):
            approvals = []
            for group in preview.get("merges", []):
                text = ("These products have the same Product Name and Size:\n\n"
                        f"{group.get('alternative_product', '')}\n{group.get('alternative_variation', '')}\n\n"
                        f"Merge them and keep all {len(group.get('originals', []))} linked ecommerce originals and seller SKUs?\n\n"
                        f"First selected base: {group.get('base_product', '')} / {group.get('base_variation', '')}\n\n"
                        "Choose No to keep separate products.")
                if messagebox.askyesno("Merge matching products?", text, parent=self, default="no"):
                    approvals.append(group["ids"])

            def saved(result):
                if self.attention_scope_ids is not None:
                    self.attention_scope_ids = remap_attention_ids(
                        {"rows": list(self.model.rows.values())}, result, self.attention_scope_ids)
                self.model.load(result)
                self._refresh()
                self.status_var.set("Names saved. Completed products are locked; review rows stay unlocked." if lock_all else "Names saved. Explicit row unlocks are kept for Excel editing.")
                if self.on_catalog_changed:
                    self.on_catalog_changed(result)
                if on_saved:
                    on_saved(result)

            self._engine({"action": "names_save", "rows": preview["rows"], "revision": self.model.revision,
                          "merge_groups": approvals, "lock_all": lock_all, "unlocked_ids": sorted(self.model.unlocked_ids),
                          "review_source_issues": self.source_issues}, saved, "Saving product names…")
        self._engine({"action": "names_prepare", "rows": rows, "priority": priority}, prepared, "Checking product names…")
        return True

    def _work(self, action, success, progress, failure=None):
        self.busy = True
        self.status_var.set(progress)
        self._state()

        def run():
            try:
                result = action()
                self._results.put((success, result, None, failure))
            except Exception as exc:
                self._results.put((success, None, exc, failure))
        threading.Thread(target=run, daemon=True).start()

    def _poll_workers(self):
        try:
            while True:
                success, result, error, failure = self._results.get_nowait()
                self.busy = False
                self._state()
                if error:
                    (failure or self._error)(error)
                else:
                    success(result)
        except queue.Empty:
            pass
        if self.winfo_exists():
            self._poll_id = self.after(70, self._poll_workers)
        else:
            self._poll_id = None

    def export_excel(self):
        if self.busy or self.edit_locked or self.excel:
            return

        def begin(_saved):
            self.excel = ExcelSession()

            def export_failed(error):
                session = self.excel

                def reset(_result):
                    self.excel = None
                    self._state()
                    self._error(error)
                self._work(session.close, reset, "Clearing the unfinished Excel workbook…")

            def exported(_result):
                session = self.excel

                def opened(_unused):
                    self.status_var.set("Edit the Excel workbook. Keep it open, then click Apply Excel edits here. Locked rows stay protected.")
                    self._state()

                def failed(error):
                    self._work(session.close, lambda _r: reset(error), "Closing the unfinished Excel session…")

                def reset(error):
                    self.excel = None
                    self._state()
                    self._error(error)
                self._work(session.open, opened, "Opening product names in Microsoft Excel…", failed)
            self._engine({"action": "names_export", "path": str(self.excel.path)}, exported, "Preparing the Excel workbook…", on_error=export_failed)
        self.save_names(lock_all=False, on_saved=begin)

    def import_excel(self):
        if self.busy or self.edit_locked or not self.excel:
            return
        session = self.excel

        def saved_excel(_result):
            def imported(preview):
                def saved(_result):
                    def finished(_unused):
                        self.excel = None
                        self._state()
                        text = f"Excel edits saved: {preview.get('added', 0)} added, {preview.get('updated', 0)} updated. Review rows stay unlocked."
                        if preview.get("separated"):
                            text += f" {preview['separated']} linked products separated."
                        self.status_var.set(text)
                    self._work(session.close, finished, "Finishing Excel import…")
                self.save_names(prepared_rows=preview["rows"], priority=preview.get("priority", []), on_saved=saved)
            self._engine({"action": "names_import", "path": str(session.path), "rows": self.model.payload()}, imported, "Checking Excel edits…")

        def failed(error):
            if str(error) == MISSING_EXCEL:
                def reset(_result):
                    self.excel = None
                    self._state()
                    self._error(error)
                self._work(session.close, reset, "Clearing the closed Excel session…")
            else:
                self._error(error)
        self._work(session.save, saved_excel, "Saving the open Excel workbook…", failed)

    def close(self):
        if self.busy:
            return
        if self.excel:
            if not messagebox.askyesno("Excel edit pending", "Close Products and discard the Excel edits that have not been applied?", parent=self, default="no"):
                return
            session = self.excel

            def reset(_result):
                self.excel = None
                self.close()
            self._work(session.close, reset, "Closing the product Excel workbook…")
            return
        if self.model.dirty:
            answer = messagebox.askyesnocancel("Unsaved products", "Save your product changes before closing?", parent=self)
            if answer is None:
                return
            if answer:
                self.save_names(on_saved=lambda _result: self._destroy())
                return
        self._destroy()

    def _destroy(self):
        if self._poll_id:
            self.after_cancel(self._poll_id)
            self._poll_id = None
        self.grab_release()
        self.destroy()


def open_products(parent, catalog, call_engine):
    return ProductsDialog(parent, catalog, call_engine)
