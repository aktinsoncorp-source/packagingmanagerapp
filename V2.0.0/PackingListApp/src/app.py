"""Packing List Manager 2.0.0: an accessible, local Python desktop interface.

The UI never changes processing rules. All source checks, price extraction,
history, deduplication and PDF generation are delegated to the existing engine.
"""
from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timedelta
import json
import os
from pathlib import Path
import queue
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont


VERSION = "2.0.0"
_SOURCE_DIR = Path(__file__).resolve().parent
_PACKAGED = _SOURCE_DIR.name == "Source" and _SOURCE_DIR.parent.name == "App Files"
from app_paths import default_data_dir
DATA_DIR = default_data_dir()
BG = "#f1f2ef"
INK = "#24292c"
MUTED = "#505b61"
BLUE = "#42484b"
LINE = "#6b767c"
PANEL = "#fcfcf9"
ROW = "#fcfcf9"
HEADER = "#42484b"
SELECTED = "#dce3e4"
ALTERNATE = "#e6e9e6"
PLATFORMS = ("SPX", "FLASH", "LAZADA", "TIKTOK", "SHOPEE")


def can_select(entry, revisions=()):
    return entry.get("status") not in ("ERROR", "COURIER") and (
        not entry.get("requires_revision") or entry["path"] in revisions
    )


def selection_counts(entries, checked, revisions=()):
    """Count the same platform/order key once, with reviewed revisions first."""
    result, seen, seen_orders = {}, set(), set()
    for entry in sorted(entries, key=lambda row: row["path"] not in revisions):
        if entry["path"] not in checked or not can_select(entry, revisions):
            continue
        platform = entry.get("platform", "")
        values = result.setdefault(platform, {"pdfs": 0, "orders": 0, "shipments": 0})
        values["pdfs"] += 1
        for order in entry.get("order_totals", []):
            if order["key"] not in seen:
                values["shipments"] += 1
                order_key = order["key"].split(":TRACKING:", 1)[0] if platform == "LAZADA" else order["key"]
                if order_key not in seen_orders:
                    values["orders"] += 1
                    seen_orders.add(order_key)
                seen.add(order["key"])
    for values in result.values():
        if values["shipments"] == values["orders"]:
            values.pop("shipments")
    return result


def generation_selection(entries, checked, revisions=()):
    return [dict(path=row["path"], sha256=row.get("sha256", ""),
                 revision=row["path"] in revisions, courier=row.get("courier", "Auto"),
                 platform=row.get("platform", ""))
            for row in entries if row["path"] in checked and can_select(row, revisions)]


def _single_instance():
    """Keep the V2 editor singular, while allowing the separate V1 app to run."""
    if os.name != "nt":
        return None
    import ctypes
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.CreateMutexW.argtypes = (ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p)
    kernel.CreateMutexW.restype = ctypes.c_void_p
    handle = kernel.CreateMutexW(None, False, "Local\\PackingListManagerV2Python")
    if ctypes.get_last_error() == 183:
        kernel.CloseHandle.argtypes = (ctypes.c_void_p,)
        kernel.CloseHandle(handle)
        raise RuntimeError("Packing List Manager 2.0.0 is already open. Use the existing window.")
    if not handle:
        raise ctypes.WinError(ctypes.get_last_error())
    return handle


def configure_styles(root):
    root.option_add("*TCombobox*Listbox.font", "{Segoe UI} 11")
    root.option_add("*TCombobox*Listbox.background", ROW)
    root.option_add("*TCombobox*Listbox.foreground", INK)
    root.option_add("*TCombobox*Listbox.selectBackground", SELECTED)
    root.option_add("*TCombobox*Listbox.selectForeground", INK)
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".", font=("Segoe UI", 11), foreground=INK, background=BG)
    style.configure("App.TFrame", background=BG)
    style.configure("Card.TFrame", background=PANEL, bordercolor=HEADER, borderwidth=2, relief="solid")
    style.configure("Paper.TFrame", background=PANEL)
    style.configure("TLabel", background=BG)
    style.configure("Card.TLabel", background=PANEL)
    style.configure("Muted.TLabel", foreground=MUTED)
    style.configure("CardMuted.TLabel", background=PANEL, foreground=MUTED, font=("Segoe UI", 10))
    style.configure("Title.TLabel", font=("Segoe UI", 21, "bold"))
    style.configure("Version.TLabel", foreground=MUTED, font=("Consolas", 9))
    style.configure("Section.TLabel", font=("Segoe UI", 13, "bold"))
    style.configure("CardHeading.TLabel", background=PANEL, font=("Segoe UI", 11, "bold"))
    style.configure("RailHeading.TLabel", font=("Segoe UI", 10, "bold"))
    style.configure("PaperSection.TLabel", background=PANEL, font=("Segoe UI", 13, "bold"))
    style.configure("TButton", padding=(12, 9), background=PANEL, bordercolor=LINE,
                    lightcolor=PANEL, darkcolor=PANEL, focuscolor=INK, borderwidth=1)
    style.map("TButton", background=[("disabled", ALTERNATE), ("active", ALTERNATE)],
              foreground=[("disabled", MUTED)], bordercolor=[("focus", INK)])
    style.configure("Primary.TButton", background=BLUE, foreground="white", bordercolor=BLUE,
                    font=("Segoe UI", 11, "bold"), padding=(16, 13), lightcolor=BLUE, darkcolor=BLUE)
    style.map("Primary.TButton", background=[("disabled", ALTERNATE), ("active", "#30383d")],
              foreground=[("disabled", MUTED), ("!disabled", "white")])
    style.configure("Date.TButton", background=HEADER, foreground=PANEL, bordercolor=HEADER,
                    lightcolor=HEADER, darkcolor=HEADER, font=("Consolas", 12, "bold"), padding=(12, 9))
    style.map("Date.TButton", background=[("disabled", ALTERNATE), ("active", "#30383d")],
              foreground=[("disabled", MUTED), ("!disabled", PANEL)])
    for name, color in (("Attention", "#ffe2a0"), ("AttentionPulse", "#efbc58")):
        style.configure(name + ".TButton", background=color, foreground=INK,
                        bordercolor=INK, lightcolor=color, darkcolor=color,
                        font=("Segoe UI", 11, "bold"), padding=(9, 9))
        style.map(name + ".TButton", background=[("disabled", ALTERNATE), ("active", "#efbc58")],
                  foreground=[("disabled", MUTED), ("!disabled", INK)])
    style.configure("Treeview", background=ROW, fieldbackground=ROW, rowheight=40,
                    borderwidth=1, bordercolor=LINE, lightcolor=LINE, darkcolor=LINE, font=("Segoe UI", 11), foreground=INK, relief="solid")
    style.configure("Treeview.Heading", background=HEADER, foreground=PANEL,
                    font=("Segoe UI", 11, "bold"), padding=(7, 10), relief="flat", bordercolor=LINE, lightcolor=LINE, darkcolor=LINE)
    style.map("Treeview", background=[("selected", SELECTED)], foreground=[("selected", INK)])
    style.map("Treeview.Heading", background=[("active", "#30383d")])
    style.configure("TEntry", padding=(5, 5), fieldbackground=ROW, foreground=INK, insertcolor=INK,
                    bordercolor=LINE, lightcolor=LINE, darkcolor=LINE, font=("Segoe UI", 11))
    style.map("TEntry", fieldbackground=[("readonly", ROW), ("disabled", ALTERNATE)],
              foreground=[("disabled", MUTED)], bordercolor=[("focus", INK)])
    style.configure("TCombobox", padding=(4, 5), fieldbackground=ROW, background=PANEL, foreground=INK,
                    arrowcolor=INK, bordercolor=LINE, lightcolor=LINE, darkcolor=LINE)
    style.map("TCombobox", fieldbackground=[("readonly", ROW)], foreground=[("readonly", INK)],
              selectbackground=[("readonly", SELECTED)], selectforeground=[("readonly", INK)])
    style.configure("TCheckbutton", background=BG, foreground=INK)
    style.map("TCheckbutton", background=[("active", BG)], foreground=[("disabled", MUTED)])
    style.configure("TScrollbar", background="#b9c1c2", troughcolor=PANEL, arrowcolor=INK,
                    bordercolor=LINE, lightcolor="#b9c1c2", darkcolor="#b9c1c2")
    style.map("TScrollbar", background=[("active", "#909b9f")])
    style.configure("Loading.Horizontal.TProgressbar", background=BLUE, troughcolor="#cbd1d0",
                    borderwidth=0, lightcolor=BLUE, darkcolor=BLUE, thickness=12)
    return style


def center_window(window, parent, width, height):
    window.update_idletasks()
    width = min(width, window.winfo_screenwidth() - 40)
    height = min(height, window.winfo_screenheight() - 80)
    x = max(0, parent.winfo_rootx() + (parent.winfo_width() - width) // 2)
    y = max(0, parent.winfo_rooty() + (parent.winfo_height() - height) // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")


class RepeatWarning(tk.Toplevel):
    """Enter, Escape and the window close button all choose the safe No path."""
    def __init__(self, parent, count):
        super().__init__(parent)
        self.title("Repeated orders")
        self.configure(background=BG)
        self.resizable(False, False)
        self.transient(parent)
        self.result = "skip"
        center_window(self, parent, 620, 340)
        ttk.Label(self, text="Some orders have been generated before", style="Section.TLabel",
                  wraplength=562).pack(anchor="w", padx=28, pady=(24, 12))
        ttk.Label(self, text=f"{count} repeated order(s) found in earlier generation history or other source PDFs.\n\nDo you want to generate the repeated orders again?",
                  wraplength=562, justify="left").pack(anchor="w", padx=28)
        ttk.Label(self, text="No will generate only orders without an existing matching output.",
                  style="Muted.TLabel", wraplength=562).pack(anchor="w", padx=28, pady=(12, 12))
        area = ttk.Frame(self, style="App.TFrame", width=500)
        area.pack(fill="x", padx=60)
        self.no_button = tk.Button(area, text="No", command=lambda: self.choose("skip"),
                                   font=("Segoe UI", 12, "bold"), background=BLUE, foreground="white",
                                   activebackground="#30383d", activeforeground="white", relief="flat",
                                   height=2, cursor="hand2", default="active", bd=0)
        self.no_button.pack(fill="x")
        self.continue_link = tk.Label(area, text="Continue to Generate the Repeated Orders",
                                      font=("Segoe UI", 12, "bold"), background=BG, foreground=INK, activebackground=ALTERNATE, activeforeground=INK,
                                      relief="flat", borderwidth=0, cursor="hand2", pady=12, takefocus=1)
        self.continue_link.pack(fill="x")
        self.continue_link.bind("<Button-1>", lambda event: self.choose("include"))
        self.continue_link.bind("<space>", lambda event: self.choose("include"))
        self.bind("<Return>", lambda event: self.choose("skip"))
        self.bind("<Escape>", lambda event: self.choose("skip"))
        self.protocol("WM_DELETE_WINDOW", lambda: self.choose("skip"))
        self.update_idletasks()
        center_window(self, parent, 620, max(340, self.winfo_reqheight()+12))
        self.grab_set()
        self.no_button.focus_set()

    def choose(self, result):
        self.result = result
        self.destroy()


class NameAttentionPrompt(tk.Toplevel):
    """An advisory with an explicit Later path; packing remains available."""
    def __init__(self, parent, report, review):
        super().__init__(parent)
        self.title("Products need attention")
        self.configure(background=BG)
        self.transient(parent)
        self.resizable(False, False)
        area = ttk.Frame(self, padding=24)
        area.pack(fill="both", expand=True)
        ttk.Label(area, text="Define these products", style="Section.TLabel").pack(anchor="w")
        counts = []
        if report.get("missing_name_count"):
            counts.append(f'{report["missing_name_count"]} product(s) need a standard Product Name or Size.')
        if report.get("missing_sku_count"):
            counts.append(f'{report["missing_sku_count"]} product(s) have a missing or unusable SKU. Check the website listing and its PDF.')
        ttk.Label(area, text="\n".join(counts), wraplength=555, justify="left").pack(fill="x", pady=(12, 10))
        examples = [item.get("original_product") or item.get("sku") or "Unnamed product"
                    for item in report.get("items", [])[:3]]
        tk.Label(area, text="\n".join("!  " + text[:90] for text in examples),
                 background="#ffe2a0", foreground=INK, font=("Segoe UI", 11, "bold"),
                 justify="left", anchor="w", wraplength=535, padx=10, pady=10).pack(fill="x")
        ttk.Label(area, text="You can define a Product Name and Size even without a SKU.\n\n"
                  "Choose Later to keep packing. Any blank standard field will print its Ecommerce Name or Ecommerce Variation instead.",
                  wraplength=555, justify="left").pack(fill="x", pady=(12, 16))
        actions = ttk.Frame(area)
        actions.pack(fill="x")
        def open_names():
            self.destroy()
            review()
        self.later_button = ttk.Button(actions, text="Later — keep packing", command=self.destroy)
        self.later_button.pack(side="left")
        self.review_button = ttk.Button(actions, text="Define now", style="Primary.TButton", command=open_names)
        self.review_button.pack(side="right")
        self.bind("<Escape>", lambda event: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.update_idletasks()
        center_window(self, parent, max(610, self.winfo_reqwidth()), self.winfo_reqheight()+4)
        self.grab_set()
        self.review_button.focus_set()


class DateChooser(tk.Toplevel):
    def __init__(self, parent, current, callback):
        import calendar
        super().__init__(parent)
        self.title("Choose packing date")
        self.configure(background=PANEL)
        self.resizable(False, False)
        self.transient(parent)
        self.current = current.replace(day=1)
        self.callback, self.calendar = callback, calendar.Calendar(firstweekday=0)
        top = ttk.Frame(self, style="Card.TFrame", padding=10)
        top.pack(fill="x")
        ttk.Button(top, text="‹", width=3, command=lambda: self.move(-1)).pack(side="left")
        self.caption = ttk.Label(top, style="CardHeading.TLabel", anchor="center")
        self.caption.pack(side="left", fill="x", expand=True)
        ttk.Button(top, text="›", width=3, command=lambda: self.move(1)).pack(side="right")
        self.days = ttk.Frame(self, style="Card.TFrame", padding=(10, 0, 10, 10))
        self.days.pack()
        ttk.Button(self, text="Today", command=lambda: self.choose(date.today())).pack(fill="x", padx=10, pady=(0, 10))
        self.render()
        self.bind("<Escape>", lambda event: self.destroy())
        self.grab_set()
        center_window(self, parent, max(400, self.winfo_reqwidth()), max(355, self.winfo_reqheight()))

    def move(self, amount):
        month = self.current.month - 1 + amount
        self.current = self.current.replace(year=self.current.year + month // 12, month=month % 12 + 1)
        self.render()

    def render(self):
        for child in self.days.winfo_children():
            child.destroy()
        self.caption.configure(text=self.current.strftime("%B %Y"))
        for col, text in enumerate(("Mo", "Tu", "We", "Th", "Fr", "Sa", "Su")):
            ttk.Label(self.days, text=text, anchor="center", style="CardMuted.TLabel").grid(row=0, column=col, sticky="ew")
        for row, week in enumerate(self.calendar.monthdatescalendar(self.current.year, self.current.month), 1):
            for col, day in enumerate(week):
                button = ttk.Button(self.days, text=str(day.day), width=3, command=lambda value=day: self.choose(value))
                button.grid(row=row, column=col, padx=1, pady=1)
                if day.month != self.current.month:
                    button.state(["disabled"])

    def choose(self, value):
        self.destroy()
        self.callback(value)


class PackingApp(tk.Tk):
    def __init__(self, data_dir=DATA_DIR, source_folder=None, engine_handler=None, auto_scan=True, ui_scale=1.0):
        super().__init__()
        if ui_scale != 1.0:
            self.tk.call("tk", "scaling", float(self.tk.call("tk", "scaling")) * ui_scale)
        self.title(f"Packing List Manager {VERSION}")
        self.configure(background=BG)
        self.minsize(1050, 680)
        width, height = min(1220, self.winfo_screenwidth()-32), min(850, self.winfo_screenheight()-80)
        self.geometry(f"{width}x{height}+{max(0,(self.winfo_screenwidth()-width)//2)}+20")
        icon = Path(__file__).parent / "assets" / "packing-list-icon.ico"
        if icon.exists():
            try:
                from PIL import Image, ImageTk
                with Image.open(icon) as artwork:
                    self._app_icons = [ImageTk.PhotoImage(artwork.ico.getimage(size), master=self)
                                       for size in sorted(artwork.ico.sizes())]
                self.iconphoto(True, *self._app_icons)
            except (tk.TclError, OSError):
                pass
        configure_styles(self)
        self.data_dir = Path(data_dir)
        self.catalog_path = self.data_dir / "ProductNames.json"
        self.settings_path = self.data_dir / "Settings.json"
        self.engine_handler = engine_handler
        self.executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="packing-engine")
        self.events = queue.Queue()
        self.busy = False
        self.entries, self.outputs = [], []
        self.checked, self.revisions = set(), set()
        self.row_entries = {}
        self.buttons = []
        self.name_attention = {"items": [], "count": 0}
        self.name_source_issues, self.name_scan_ids = [], set()
        self._attention_seen = set()
        self._attention_timer = self._attention_prompt_timer = None
        self._attention_bright = False
        self.root_folder = Path.home() / "Documents" if _PACKAGED else Path(__file__).resolve().parents[3]
        self.packing_date = date.today()
        try:
            saved = json.loads(self.settings_path.read_text(encoding="utf-8-sig"))
            if Path(saved.get("folder", "")).is_dir():
                self.root_folder = Path(saved["folder"])
        except (OSError, ValueError, KeyError):
            pass
        if source_folder:
            self._set_folder(Path(source_folder))
        self._build_ui()
        self._compact = None
        self.bind("<Configure>", self._responsive_layout, add="+")
        self._update_paths()
        self._refresh_selection()
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.bind("<F5>", lambda event: self.scan() if not self.busy else None)
        self._poll_id = self.after(50, self._poll)
        if auto_scan:
            self.after(100, self.scan)

    @property
    def date_folder(self):
        return self.root_folder / self.packing_date.strftime("%Y%m%d")

    @property
    def output_folder(self):
        return self.date_folder / "000 Output"

    def _button(self, parent, text, command, primary=False, **kwargs):
        button = ttk.Button(parent, text=text, command=command,
                            style="Primary.TButton" if primary else "TButton", **kwargs)
        self.buttons.append(button)
        return button

    def _build_ui(self):
        outer = ttk.Frame(self, style="App.TFrame", padding=(22, 17, 22, 12))
        self.outer_panel = outer
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1)
        icon_dir = Path(__file__).resolve().parent / "assets" / "ui"
        self.nav_lock_icon = tk.PhotoImage(master=self, file=str(icon_dir / "lock-closed.png"))
        self.nav_disabled_lock_icon = tk.PhotoImage(master=self, file=str(icon_dir / "lock-closed-disabled.png"))
        header = ttk.Frame(outer, style="App.TFrame")
        self.header_panel = header
        header.grid(row=0, column=0, sticky="ew", pady=(0, 17))
        header.columnconfigure(0, weight=1)
        self.subtitle = ttk.Label(header, text=f"V{VERSION}", style="Version.TLabel")
        self.subtitle.grid(row=0, column=0, sticky="w", pady=(0, 3))
        ttk.Label(header, text="PACKING LIST MANAGER", style="Title.TLabel").grid(row=1, column=0, sticky="w")
        self.products_button = self._button(header, "Products / Define Names", self.open_products,
                                          image=(self.nav_lock_icon, "disabled", self.nav_disabled_lock_icon), compound="left")
        self.products_button.grid(row=0, column=1, rowspan=2, padx=(12, 9))
        self.attention_button = self._button(header, "! Review (0)", lambda: self.open_products(attention_only=True))
        self.attention_button.configure(style="Attention.TButton")
        self.attention_button.grid(row=0, column=2, rowspan=2, padx=(0, 9))
        self.attention_button.grid_remove()
        self._button(header, "Help", self.show_help, width=5).grid(row=0, column=3, rowspan=2)
        self.header_rule = tk.Frame(header, background=HEADER, height=6)
        self.header_rule.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(12, 0))

        folder = ttk.Frame(outer, style="App.TFrame")
        folder.grid(row=1, column=0, sticky="ew", pady=(0, 16))
        self.folder_panel = folder
        folder.columnconfigure(0, weight=1)
        ttk.Label(folder, text="Source folder", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Label(folder, text="Packing date", font=("Segoe UI", 11, "bold")).grid(row=0, column=2, sticky="w", padx=(22, 0), pady=(0, 6))
        self.folder_var = tk.StringVar()
        self.folder_entry = ttk.Entry(folder, textvariable=self.folder_var, state="readonly", font=("Segoe UI", 11))
        self.folder_entry.grid(row=1, column=0, sticky="ew", padx=(0, 10))
        self._button(folder, "Choose folder", self.choose_folder).grid(row=1, column=1)
        self.date_var = tk.StringVar()
        self.date_button = self._button(folder, "", self.choose_date, textvariable=self.date_var, width=16)
        self.date_button.configure(style="Date.TButton")
        self.date_button.grid(row=1, column=2, padx=(22, 0))
        self.path_var = tk.StringVar()
        self.path_label = ttk.Label(folder, textvariable=self.path_var, style="Muted.TLabel", font=("Segoe UI", 10))
        self.path_label.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(7, 0))
        folder.bind("<Configure>", lambda event: self.path_label.configure(wraplength=max(300, event.width-35)))

        body = ttk.Frame(outer, style="App.TFrame")
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, minsize=246)
        body.rowconfigure(0, weight=1)
        workspace = ttk.Frame(body, style="Card.TFrame", padding=1)
        workspace.grid(row=0, column=0, sticky="nsew")
        workspace.columnconfigure(0, weight=1)
        workspace.rowconfigure(1, weight=1)
        self.workspace_panel = workspace
        rail = ttk.Frame(body, style="App.TFrame")
        rail.grid(row=0, column=1, sticky="nsew", padx=(16, 0))
        rail.columnconfigure(0, weight=1)
        self.action_rail = rail

        heading = ttk.Frame(workspace, style="Paper.TFrame", padding=(12, 10))
        self.files_heading = heading
        heading.grid(row=0, column=0, sticky="ew")
        heading.columnconfigure(0, weight=1)
        self.summary_var = tk.StringVar(value="Packing PDFs")
        self.files_summary_label = ttk.Label(heading, textvariable=self.summary_var, style="PaperSection.TLabel", wraplength=330)
        self.files_summary_label.grid(row=0, column=0, sticky="w")
        self.refresh_button = self._button(heading, "Refresh", self.scan)
        self.refresh_button.grid(row=0, column=1, padx=(10, 8))
        self.add_button = self._button(heading, "Add PDFs", self.add_pdfs)
        self.add_button.grid(row=0, column=2)
        heading.bind("<Configure>", lambda event: self.files_summary_label.configure(
            wraplength=max(140, event.width-self.refresh_button.winfo_reqwidth()-self.add_button.winfo_reqwidth()-50)))

        files = ttk.Frame(workspace, style="Paper.TFrame")
        files.grid(row=1, column=0, sticky="nsew")
        files.columnconfigure(0, weight=1)
        files.rowconfigure(0, weight=1)
        self.check_icons = {state: tk.PhotoImage(master=self, file=str(icon_dir / f"check-{state}.png"))
                            for state in ("off", "on", "mixed", "disabled")}
        # The native tree column supports images; omit its unused expand arrow.
        ttk.Style(self).layout("Files.Treeview.Item", [("Treeitem.padding", {
            "sticky": "nswe", "children": [("Treeitem.image", {"sticky": ""})]})])
        ttk.Style(self).layout("Files.Treeview.Heading", [
            ("Treeheading.cell", {"sticky": "nswe"}),
            ("Treeheading.border", {"sticky": "nswe", "children": [
                ("Treeheading.padding", {"sticky": "nswe", "children": [
                    ("Treeheading.image", {"side": "left", "sticky": ""}),
                    ("Treeheading.text", {"sticky": "we"})]})]})])
        columns = ("platform", "file", "courier", "orders", "status")
        self.tree = ttk.Treeview(files, columns=columns, show="tree headings", style="Files.Treeview", selectmode="browse", height=6)
        heading_font = tkfont.Font(self, font=("Segoe UI", 11, "bold"))
        for key, label, width, minimum, stretch in (
            ("#0", "Use", 90, 90, False), ("platform", "Platform", 95, 82, False),
            ("file", "PDF file", 250, 140, True), ("courier", "Courier", 94, 86, False),
            ("orders", "Orders", 70, 62, False), ("status", "Status", 230, 160, True)):
            if not stretch:
                minimum = max(minimum, heading_font.measure(label) + 26)
                width = max(width, minimum)
            self.tree.heading(key, text=label, anchor="center" if key == "#0" else "w", command=self.toggle_all if key == "#0" else lambda: None)
            self.tree.column(key, width=width, minwidth=minimum, stretch=stretch, anchor="center" if key in ("#0", "orders", "courier") else "w")
        self.tree.tag_configure("even", background=ROW)
        self.tree.tag_configure("odd", background=ALTERNATE)
        self.tree.tag_configure("issue", foreground="#843d0a")
        self.tree.tag_configure("ready", foreground=INK)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(files, command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
        horizontal = ttk.Scrollbar(files, orient="horizontal", command=self.tree.xview)
        horizontal.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.tree.configure(xscrollcommand=horizontal.set)
        self.tree.bind("<Configure>", self._fit_file_columns, add="+")
        from products_ui import add_column_dividers
        self.refresh_file_dividers = add_column_dividers(self.tree, color=LINE)
        self.tree.bind("<Button-1>", self._tree_click)
        self.tree.bind("<space>", self._toggle_current)
        self.tree.bind("<Return>", lambda event: self.file_details())
        self.tree.bind("<Double-1>", self._tree_double_click)
        self.tree.bind("<<TreeviewSelect>>", lambda event: self._refresh_selection())
        self.empty_label = ttk.Label(files, text="No packing PDFs for this date\nChoose the correct folder and date, or click Add PDFs.",
                                     style="CardMuted.TLabel", anchor="center", justify="center")
        totals = tk.Frame(files, background=ALTERNATE, padx=12, pady=8, highlightbackground=HEADER, highlightthickness=1)
        totals.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.totals_var = tk.StringVar()
        self.platform_totals_var = tk.StringVar()
        tk.Label(totals, textvariable=self.totals_var, background=ALTERNATE, foreground=INK,
                 font=("Segoe UI", 11, "bold"), anchor="w").pack(fill="x")
        self.platform_totals_label = tk.Label(totals, textvariable=self.platform_totals_var, background=ALTERNATE, foreground=MUTED,
                                             font=("Segoe UI", 10), anchor="w", justify="left", wraplength=650)
        self.platform_totals_label.pack(fill="x", pady=(3, 0))
        totals.bind("<Configure>", lambda event: self.platform_totals_label.configure(wraplength=max(200, event.width-28)))

        actions = ttk.Frame(workspace, style="Paper.TFrame", padding=(12, 10))
        actions.grid(row=2, column=0, sticky="ew")
        actions.columnconfigure(3, weight=1)
        self.source_button = self._button(actions, "Open PDF", self.open_source)
        self.source_button.grid(row=0, column=0, padx=(0, 8))
        self.details_button = self._button(actions, "File details", self.file_details)
        self.details_button.grid(row=0, column=1, padx=(0, 8))
        self.review_button = self._button(actions, "Review changed file", self.review)
        self.review_button.grid(row=0, column=2)

        generate = ttk.Frame(rail, style="Card.TFrame", padding=12)
        generate.grid(row=0, column=0, sticky="ew")
        generate.columnconfigure(0, weight=1)
        self.generate_panel = generate
        self.create_heading = ttk.Label(generate, text="CREATE DOCUMENTS", style="CardHeading.TLabel")
        self.create_heading.grid(row=0, column=0, sticky="w", pady=(0, 10))
        self.generate_button = self._button(generate, "Generate\nPacking Lists", self.generate, primary=True)
        self.generate_button.grid(row=1, column=0, sticky="ew")
        self.generate_hint = tk.StringVar()
        self.generate_hint_label = ttk.Label(generate, textvariable=self.generate_hint, style="CardMuted.TLabel", justify="left", wraplength=205)
        self.generate_hint_label.grid(row=2, column=0, sticky="w", pady=10)
        self.sales_button = self._button(generate, "Export Item Sales", self.export_sales)
        self.sales_button.grid(row=3, column=0, sticky="ew")
        saved = ttk.Frame(rail, style="App.TFrame")
        self.saved_panel = saved
        saved.grid(row=1, column=0, sticky="ew", pady=(15, 0))
        saved.columnconfigure(0, weight=1)
        self.saved_rule = tk.Frame(saved, background=HEADER, height=3)
        self.saved_rule.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.saved_heading = ttk.Label(saved, text="SAVED DOCUMENTS", style="RailHeading.TLabel")
        self.saved_heading.grid(row=1, column=0, sticky="w", pady=(0, 8))
        self.history_button = self._button(saved, "Completed documents", self.show_history)
        self.history_button.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        self.output_button = self._button(saved, "Open Output Folder", self.open_output)
        self.output_button.grid(row=3, column=0, sticky="ew")
        footer = ttk.Frame(outer, style="App.TFrame")
        footer.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        footer.columnconfigure(0, weight=1)
        tk.Frame(footer, background=HEADER, height=2).grid(row=0, column=0, sticky="ew", pady=(0, 7))
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(footer, textvariable=self.status_var, style="Muted.TLabel", font=("Segoe UI", 9))
        self.status_label.grid(row=1, column=0, sticky="ew")
        footer.bind("<Configure>", lambda event: self.status_label.configure(wraplength=max(300, event.width)))
        self.progress = ttk.Progressbar(footer, mode="indeterminate", style="Loading.Horizontal.TProgressbar")
        self.progress.grid(row=2, column=0, sticky="ew", pady=(5, 0))
        self.progress.grid_remove()

    def _responsive_layout(self, event=None):
        if event is not None and event.widget is not self:
            return
        compact = self.winfo_height() < 780
        if compact == self._compact:
            return
        self._compact = compact
        self.outer_panel.configure(padding=(20, 10 if compact else 17, 20, 8 if compact else 12))
        self.header_panel.grid_configure(pady=(0, 7 if compact else 17))
        self.header_rule.grid_configure(pady=(6 if compact else 12, 0))
        self.folder_panel.grid_configure(pady=(0, 6 if compact else 16))
        self.files_heading.configure(padding=(12, 7 if compact else 10))
        self.generate_panel.configure(padding=8 if compact else 12)
        self.create_heading.grid_configure(pady=(0, 7 if compact else 10))
        self.generate_button.configure(padding=(16, 10 if compact else 13))
        self.generate_hint_label.grid_configure(pady=6 if compact else 10)
        self.saved_panel.grid_configure(pady=(12 if compact else 15, 0))
        self.saved_rule.grid_configure(pady=(0, 7 if compact else 10))
        self.saved_heading.grid_configure(pady=(0, 6 if compact else 8))
        self.history_button.grid_configure(pady=(0, 6 if compact else 8))
        if compact:
            self.path_label.grid_remove()
        else:
            self.path_label.grid()
        self._refresh_selection()

    def _fit_file_columns(self, event=None):
        """Keep headings visible; scroll only if the readable minima cannot fit."""
        available = self.tree.winfo_width() - 4 - sum(
            self.tree.column(key, "width") for key in ("#0", "platform", "courier", "orders"))
        file_width = max(140, int(available * .55))
        status_width = max(160, available-file_width)
        if file_width+status_width > available and available >= 300:
            file_width = available-status_width
        self.tree.column("file", width=file_width)
        self.tree.column("status", width=status_width)

    def _set_folder(self, path):
        try:
            selected_date = datetime.strptime(path.name, "%Y%m%d").date()
            if len(path.name) != 8:
                raise ValueError()
            self.root_folder, self.packing_date = path.parent, selected_date
        except ValueError:
            self.root_folder = path

    def _update_paths(self):
        self.folder_var.set(str(self.root_folder))
        self.date_var.set(self.packing_date.strftime("%d %b %Y").upper() + "  ▾")
        self.path_var.set("Reading: " + str(self.date_folder))

    def choose_folder(self):
        selected = filedialog.askdirectory(parent=self, title="Choose a main packing folder or a YYYYMMDD dated folder", initialdir=str(self.root_folder))
        if not selected:
            return
        self._set_folder(Path(selected))
        self._update_paths()
        try:
            self.data_dir.mkdir(parents=True, exist_ok=True)
            settings = {"folder": str(self.root_folder)}
            temp = self.settings_path.with_suffix(".tmp")
            temp.write_text(json.dumps(settings, indent=2), encoding="utf-8")
            os.replace(temp, self.settings_path)
        except OSError as exc:
            self.show_error("The folder is selected, but the preference could not be saved: " + str(exc))
        self.scan()

    def choose_date(self):
        def chosen(value):
            self.packing_date = value
            self._update_paths()
            self.scan()
        DateChooser(self, self.packing_date, chosen)

    def call_engine(self, request, on_success, message="Working...", on_error=None):
        if self.busy:
            error = RuntimeError("Please wait for the current task to finish.")
            if on_error:
                on_error(error)
            return
        request = dict(request, folder=str(self.date_folder), catalog_path=str(self.catalog_path))
        self._set_busy(True, message)
        def run():
            if self.engine_handler:
                return self.engine_handler(request)
            from engine import handle
            return handle(request)
        future = self.executor.submit(run)
        future.add_done_callback(lambda completed: self.events.put((completed, on_success, on_error)))

    def _poll(self):
        try:
            while True:
                future, success, failure = self.events.get_nowait()
                self._set_busy(False, "Ready")
                try:
                    result = future.result()
                    if not result.get("ok", False):
                        raise RuntimeError(result.get("error", "The task could not be completed."))
                    success(result)
                except Exception as exc:
                    if failure:
                        failure(exc)
                    else:
                        self.show_error(exc)
        except queue.Empty:
            pass
        if self.winfo_exists():
            self._poll_id = self.after(50, self._poll)

    def _set_busy(self, value, message):
        self.busy = value
        self.status_var.set(message)
        for button in self.buttons:
            button.state(["disabled"] if value else ["!disabled"])
        if value:
            self.progress.grid()
            self.progress.start(18)
        else:
            self.progress.stop()
            self.progress.grid_remove()
            self._refresh_selection()

    def show_error(self, error):
        text = str(error)
        self.status_var.set("Needs attention: " + text.replace("\n", " ")[:190])
        messagebox.showwarning("Needs attention", text, parent=self)

    def scan(self, preserve=False, after=None):
        previous = set(self.checked) if preserve else None
        focus = self.current_entry()
        focus_path = focus["path"] if focus else None
        def done(result):
            self.entries, self.outputs = result.get("entries", []), result.get("outputs", [])
            self.revisions.clear()
            self.checked = {row["path"] for row in self.entries if can_select(row) and (
                row["path"] in previous if previous is not None else row.get("status") in ("READY", "MISSING"))}
            self._populate_files(focus_path)
            self.status_var.set("New orders are checked automatically. Tick Use to include a PDF, or select a row for details." if self.entries else "No supported packing PDFs found. Check the folder and date, or use Add PDFs.")
            self.name_source_issues = result.get("name_source_issues", [])
            self.name_scan_ids = set(result.get("name_scan_ids", []))
            self._set_name_attention(result.get("name_scan_attention", {"items": [], "count": 0}), prompt=after is None)
            if after:
                after(result)
        def failed(exc):
            self.entries, self.outputs, self.checked, self.revisions = [], [], set(), set()
            self.name_source_issues, self.name_scan_ids = [], set()
            self._populate_files()
            self._set_name_attention({"items": [], "count": 0})
            self.show_error(exc)
        self.call_engine({"action": "scan"}, done, "Reading PDFs and checking completed documents...", failed)

    def _populate_files(self, focus_path=None):
        self.tree.delete(*self.tree.get_children())
        self.row_entries.clear()
        issues = sum(row.get("status") in ("REVIEW", "ERROR", "COURIER") for row in self.entries)
        fresh = sum(row.get("new_orders", 0) for row in self.entries)
        unit = "shipments" if any(row.get("shipments", row.get("orders", 0)) > row.get("orders", 0) for row in self.entries) else "orders"
        text = f"Packing PDFs · {fresh} new {unit}"
        if issues:
            text += f" · {issues} need attention"
        self.summary_var.set(text)
        for number, row in enumerate(self.entries):
            key = str(number)
            tags = ["even" if number % 2 == 0 else "odd"]
            if row.get("status") in ("REVIEW", "ERROR", "COURIER"):
                tags.append("issue")
            elif row.get("status") in ("READY", "MISSING"):
                tags.append("ready")
            self.tree.insert("", "end", iid=key, image=self._row_check_icon(row), values=self._row_values(row), tags=tags)
            self.row_entries[key] = row
            if row["path"] == focus_path:
                self.tree.selection_set(key)
                self.tree.see(key)
        if not self.entries:
            self.empty_label.place(relx=.5, rely=.44, anchor="center")
        else:
            self.empty_label.place_forget()
            if not self.tree.selection():
                self.tree.selection_set("0")
        self._refresh_selection()

    def _row_values(self, row):
        return (row.get("platform", ""), row["path"],
                (row.get("courier", "Auto") + " ▾") if row.get("is_shopee") else "—",
                row.get("orders", 0), row.get("message", ""))

    def _row_check_icon(self, row):
        state = "on" if row["path"] in self.checked else "off" if can_select(row, self.revisions) else "disabled"
        return self.check_icons[state]

    def current_entry(self):
        selection = self.tree.selection()
        return self.row_entries.get(selection[0]) if selection else None

    def _tree_click(self, event):
        if self.busy:
            return "break"
        if self.tree.identify_region(event.x, event.y) == "heading" and self.tree.identify_column(event.x) == "#0":
            self.toggle_all()
            return "break"
        item, column = self.tree.identify_row(event.y), self.tree.identify_column(event.x)
        if not item:
            return
        self.tree.selection_set(item)
        if column == "#0":
            self._toggle_current()
            return "break"
        if column == "#3":
            self.choose_courier(item, event.x_root, event.y_root)
            return "break"

    def _tree_double_click(self, event):
        if self.tree.identify_region(event.x, event.y) == "heading":
            return self._tree_click(event)
        if self.tree.identify_column(event.x) not in ("#0", "#3"):
            self.file_details()
        return "break"

    def _toggle_current(self, event=None):
        row = self.current_entry()
        if self.busy or row is None or not can_select(row, self.revisions):
            return "break"
        if row["path"] in self.checked:
            self.checked.remove(row["path"])
        else:
            self.checked.add(row["path"])
        self._refresh_selection()
        return "break"

    def toggle_all(self):
        if self.busy:
            return
        eligible = {row["path"] for row in self.entries if can_select(row, self.revisions)}
        self.checked = set() if eligible <= self.checked else eligible
        self._refresh_selection()

    def _refresh_selection(self):
        for key, row in self.row_entries.items():
            self.tree.item(key, image=self._row_check_icon(row), values=self._row_values(row))
        counts = selection_counts(self.entries, self.checked, self.revisions)
        pdfs = sum(value["pdfs"] for value in counts.values())
        orders = sum(value["orders"] for value in counts.values())
        shipments = sum(value.get("shipments", value["orders"]) for value in counts.values())
        self.totals_var.set(f"Selected: {pdfs} PDFs   |   {orders} distinct orders" + (f" / {shipments} shipments" if shipments != orders else ""))
        self.platform_totals_var.set("     ".join(f'{platform}: {counts[platform]["pdfs"]} PDFs / {counts[platform]["orders"]} orders' +
            (f' / {counts[platform]["shipments"]} shipments' if "shipments" in counts[platform] else "")
            for platform in PLATFORMS if platform in counts) or "Tick Use beside the PDFs you want to generate.")
        if getattr(self, "_compact", False):
            self.generate_hint.set("A4 PDF · 000 Output" if pdfs else "Tick Use beside a PDF.")
        else:
            self.generate_hint.set("A4 PDF · Selected orders\nSaved in 000 Output" if pdfs else "Select a PDF to begin.\nTick its Use checkbox.")
        eligible = {row["path"] for row in self.entries if can_select(row, self.revisions)}
        state = "on" if eligible and eligible <= self.checked else "mixed" if self.checked else "off"
        self.tree.heading("#0", image=self.check_icons[state])
        row = self.current_entry()
        for button, allowed in ((self.generate_button, pdfs > 0), (self.source_button, row is not None),
                                (self.details_button, row is not None), (self.review_button, bool(row and row.get("requires_revision"))),
                                (self.history_button, bool(self.outputs))):
            button.state(["!disabled"] if allowed and not self.busy else ["disabled"])

    def choose_courier(self, key, x, y):
        row = self.row_entries[key]
        if not row.get("is_shopee"):
            return
        menu = tk.Menu(self, tearoff=False, font=("Segoe UI", 11), background=PANEL, foreground=INK,
                       activebackground=SELECTED, activeforeground=INK)
        for courier in ("Auto", "SPX", "FLASH"):
            menu.add_command(label=courier + ("  ✓" if courier == row.get("courier", "Auto") else ""),
                             command=lambda value=courier: self.set_courier(row, value))
        menu.tk_popup(x, y)

    def set_courier(self, row, mode):
        if mode == row.get("courier", "Auto"):
            return
        self.call_engine({"action": "courier", "carrier": mode, "selected": {"path": row["path"], "sha256": row["sha256"]}},
                         lambda result: self.scan(preserve=True), "Saving courier selection...")

    def open_path(self, value):
        path = Path(value)
        try:
            if not path.exists():
                raise FileNotFoundError(f"This file or folder is no longer available:\n\n{path}")
            os.startfile(str(path))
        except OSError as exc:
            self.show_error(exc)

    def open_source(self):
        row = self.current_entry()
        if row:
            self.open_path(self.date_folder / row["path"])

    def open_output(self):
        if self.output_folder.exists():
            self.open_path(self.output_folder)
        else:
            messagebox.showinfo("Output folder", f"Your documents will be saved here:\n\n{self.output_folder}\n\nGenerate packing lists or export item sales to create this folder.", parent=self)

    def add_pdfs(self):
        paths = filedialog.askopenfilenames(parent=self, title="Choose packing-list PDFs", filetypes=[("PDF files", "*.pdf")])
        if not paths:
            return
        def done(result):
            text = f'{len(result.get("imported", []))} PDF(s) added. {len(result.get("skipped", []))} identical file(s) already present.'
            if result.get("errors"):
                text += "\n\n" + "\n".join(f'{row["file"]}: {row["message"]}' for row in result["errors"])
                messagebox.showwarning("Add PDFs", text, parent=self)
            else:
                messagebox.showinfo("Add PDFs", text, parent=self)
            self.scan()
        self.call_engine({"action": "import", "files": list(paths)}, done, "Reading and adding the selected PDFs...")

    def file_details(self):
        row = self.current_entry()
        if not row or self.busy:
            return
        dialog = self._text_dialog("Packing PDF details", row["path"], row.get("details", row.get("message", "")))
        self._dialog_button(dialog.actions, "Open source PDF", lambda: self.open_path(self.date_folder / row["path"])).pack(side="left")
        if row.get("is_shopee"):
            ttk.Label(dialog.actions, text="Courier: " + row.get("courier", "Auto") + "  ·  Change it in the main list.", style="Muted.TLabel").pack(side="left", padx=14)

    def _dialog_button(self, parent, text, command, primary=False):
        return ttk.Button(parent, text=text, command=command, style="Primary.TButton" if primary else "TButton")

    def _text_dialog(self, title, heading, text):
        dialog = tk.Toplevel(self)
        dialog.title(title)
        dialog.configure(background=BG)
        dialog.transient(self)
        dialog.minsize(650, 420)
        center_window(dialog, self, 980, 660)
        container = ttk.Frame(dialog, padding=18, style="App.TFrame")
        container.pack(fill="both", expand=True)
        ttk.Label(container, text=heading, style="Section.TLabel", wraplength=880).pack(anchor="w", pady=(0, 14))
        area = ttk.Frame(container, style="Card.TFrame")
        area.pack(fill="both", expand=True)
        box = tk.Text(area, wrap="word", background=ROW, foreground=INK, selectbackground=SELECTED, selectforeground=INK,
                      relief="flat", padx=12, pady=12, font=("Segoe UI", 11))
        box.insert("1.0", text)
        box.configure(state="disabled")
        box.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(area, command=box.yview)
        scroll.pack(side="right", fill="y")
        box.configure(yscrollcommand=scroll.set)
        dialog.actions = ttk.Frame(container, style="App.TFrame")
        dialog.actions.pack(fill="x", pady=(12, 0))
        ttk.Button(dialog.actions, text="Close", command=dialog.destroy).pack(side="right")
        dialog.bind("<Escape>", lambda event: dialog.destroy())
        dialog.grab_set()
        return dialog

    def review(self):
        row = self.current_entry()
        if not row or not row.get("requires_revision"):
            return
        dialog = self._text_dialog("Review changed packing file", "Compare the earlier packing record with the current PDF",
                                   "Check whether these orders have already been packed before creating a revision.\n\n" + row.get("details", ""))
        def revise():
            self.revisions.add(row["path"])
            self.checked.add(row["path"])
            row["message"] = "Reviewed — revision selected"
            self._refresh_selection()
            dialog.destroy()
        def keep():
            dialog.destroy()
            self.call_engine({"action": "acknowledge", "selected": {"path": row["path"], "sha256": row["sha256"]}},
                             lambda result: self.scan(), "Saving the changed-file review...")
        ttk.Button(dialog.actions, text="Create revision", command=revise).pack(side="left", padx=(0, 8))
        ttk.Button(dialog.actions, text="Keep earlier output", command=keep).pack(side="left")

    def generate(self):
        selected = generation_selection(self.entries, self.checked, self.revisions)
        if not selected or self.busy:
            return
        request = {"action": "generate", "selected": selected, "output_format": "pdf"}
        def completed(result):
            if result.get("confirmation_required"):
                warning = RepeatWarning(self, result.get("repeat_count", 0))
                self.wait_window(warning)
                request["repeat_action"] = warning.result
                request["repeat_confirmation"] = result["repeat_token"]
                self.call_engine(request, completed, "Generating confirmed packing lists. Please wait...")
                return
            made = result.get("outputs", [])
            self.scan(after=lambda ignored: self._generation_done(result, made))
        self.call_engine(request, completed, "Preparing PDF packing lists and checking A4 pages. Please wait...")

    def _generation_done(self, result, made):
        self.status_var.set(result.get("message", "Packing lists ready"))
        if not made:
            messagebox.showinfo("No new orders", result.get("message", "All selected orders already have a matching output. No new document was created."), parent=self)
            return
        dialog = tk.Toplevel(self)
        dialog.title("Packing lists ready")
        dialog.configure(background=BG)
        dialog.transient(self)
        center_window(dialog, self, 820, 390)
        frame = ttk.Frame(dialog, padding=22, style="App.TFrame")
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Your packing lists are ready", style="Section.TLabel").pack(anchor="w")
        text = "Open a document below and print on A4."
        if result.get("skipped_repeat_orders"):
            text += f' {result["skipped_repeat_orders"]} already-generated order(s) skipped.'
        ttk.Label(frame, text=text, style="Muted.TLabel", wraplength=750).pack(anchor="w", pady=(6, 14))
        files = tk.Listbox(frame, font=("Segoe UI", 11), relief="flat", activestyle="none", background=ROW, foreground=INK,
                           selectbackground=SELECTED, selectforeground=INK, highlightbackground=LINE)
        files.pack(fill="both", expand=True)
        for row in made:
            files.insert("end", row["filename"] + f'  ({row.get("pages", "")} pages)')
        files.selection_set(0)
        def open_selected(event=None):
            if files.curselection():
                self.open_path(made[files.curselection()[0]]["path"])
        files.bind("<Double-1>", open_selected)
        actions = ttk.Frame(frame, style="App.TFrame")
        actions.pack(fill="x", pady=(14, 0))
        ttk.Button(actions, text="Open document", command=open_selected).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Open Output Folder", command=self.open_output).pack(side="left")
        close = ttk.Button(actions, text="Done", command=dialog.destroy)
        close.pack(side="right")
        dialog.bind("<Return>", lambda event: dialog.destroy())
        dialog.bind("<Escape>", lambda event: dialog.destroy())
        dialog.grab_set()
        close.focus_set()

    def export_sales(self):
        def done(result):
            self.status_var.set(result.get("message", "Item sales exported"))
            messagebox.showinfo("Item sales exported", result.get("message", "Export complete.") +
                                "\n\nThis report includes every source PDF for the selected date. The Use checkboxes do not limit it.\n\nSaved to:\n" + result["path"], parent=self)
            self.open_path(result["path"])
        self.call_engine({"action": "sales_export"}, done, "Exporting item quantities for the full selected date...")

    def _set_name_attention(self, report, prompt=False):
        self.name_attention = report
        for attribute in ("_attention_timer", "_attention_prompt_timer"):
            timer = getattr(self, attribute)
            if timer:
                self.after_cancel(timer)
                setattr(self, attribute, None)
        self._attention_bright = False
        self.attention_button.configure(style="Attention.TButton", text=f'! Review ({report.get("count", 0)})')
        if report.get("count"):
            self.attention_button.grid()
            self._attention_timer = self.after(1000, self._pulse_attention)
            if prompt:
                self._attention_prompt_timer = self.after(150, self._prompt_name_attention)
        else:
            self.attention_button.grid_remove()

    def _pulse_attention(self):
        self._attention_timer = None
        if not self.name_attention.get("count"):
            return
        self._attention_bright = not self._attention_bright if not self.busy and self.grab_current() is None else False
        self.attention_button.configure(style="AttentionPulse.TButton" if self._attention_bright else "Attention.TButton")
        self._attention_timer = self.after(1000, self._pulse_attention)

    def _prompt_name_attention(self):
        self._attention_prompt_timer = None
        keys = {(str(self.date_folder), item["id"], reason)
                for item in self.name_attention.get("items", []) for reason in item["reasons"]}
        if not keys - self._attention_seen:
            return
        if self.busy or self.grab_current() is not None:
            self._attention_prompt_timer = self.after(1000, self._prompt_name_attention)
            return
        self._attention_seen.update(keys)
        NameAttentionPrompt(self, self.name_attention, lambda: self.open_products(attention_only=True))

    def open_products(self, attention_only=False):
        def done(catalog):
            from products_ui import ProductsDialog
            from product_attention import catalog_attention, remap_attention_ids
            saved = False
            def changed(result):
                nonlocal saved, catalog
                saved = True
                self.name_scan_ids = remap_attention_ids(catalog, result, self.name_scan_ids)
                catalog = result
                report = catalog_attention({"rows": [row for row in result.get("rows", [])
                                                    if row["id"] in self.name_scan_ids]}, self.name_source_issues)
                self._set_name_attention(report)
            dialog = ProductsDialog(self, catalog, self.call_engine, attention_only=attention_only,
                                    source_issues=self.name_source_issues,
                                    attention_ids=self.name_scan_ids if attention_only else None,
                                    on_catalog_changed=changed)
            def closed(event):
                if event.widget is dialog and saved:
                    self.after_idle(lambda: self.scan(preserve=True))
            dialog.bind("<Destroy>", closed, add="+")
        self.call_engine({"action": "names_list"}, done, "Loading products and SKU links...")

    def show_history(self):
        dialog = tk.Toplevel(self)
        dialog.title("Completed packing documents")
        dialog.configure(background=BG)
        dialog.transient(self)
        dialog.minsize(900, 480)
        center_window(dialog, self, 1100, 680)
        area = ttk.Frame(dialog, padding=18, style="App.TFrame")
        area.pack(fill="both", expand=True)
        ttk.Label(area, text="Completed documents", style="Section.TLabel").pack(anchor="w")
        ttk.Label(area, text="Expand a document to see its source PDFs, orders and printed products.", style="Muted.TLabel").pack(anchor="w", pady=(5, 12))
        frame = ttk.Frame(area, style="Card.TFrame")
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, show="tree", selectmode="browse")
        tree.column("#0", width=980, minwidth=780, stretch=True)
        tree.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(frame, command=tree.yview)
        scroll.pack(side="right", fill="y")
        tree.configure(yscrollcommand=scroll.set)
        links = {}
        actions = ttk.Frame(area, style="App.TFrame")
        actions.pack(fill="x", pady=(12, 0))
        def populate():
            tree.delete(*tree.get_children())
            links.clear()
            groups = {}
            for output in self.outputs:
                if not (self.output_folder / output["filename"]).is_file():
                    continue
                group = output.get("group", "Other")
                if group not in groups:
                    groups[group] = tree.insert("", "end", text="Lazada + TikTok" if group == "LAZADA_TIKTOK" else group, open=True)
                suffix = "" if output.get("status") == "active" else f'  [{output.get("status", "")}]'
                count = output.get("order_count", 0)
                shipments = output.get("shipments", count)
                shipment_label = f' / {shipments} shipments' if shipments != count else ""
                document = tree.insert(groups[group], "end", text=f'{output["filename"]}  |  {count} orders{shipment_label}{suffix}')
                links[document] = output, None
                for source in output.get("included_sources", []):
                    source_orders = source.get("orders", [])
                    count = len({(o.get("platform"), o.get("order_id")) for o in source_orders})
                    shipment_label = f' / {len(source_orders)} shipments' if len(source_orders) != count else ""
                    node = tree.insert(document, "end", text=f'{source["path"]}  |  {count} orders{shipment_label}')
                    links[node] = output, source["path"]
                    for order in source.get("orders", []):
                        label = "Order Number" if order.get("platform") == "LAZADA" else "Order ID"
                        tracking = f' | Tracking: {order["tracking_id"]}' if order.get("tracking_id") else ""
                        item = tree.insert(node, "end", text=f'{label}: {order.get("order_id", "")}{tracking}')
                        links[item] = output, source["path"]
                        for product in order.get("items", []):
                            row = tree.insert(item, "end", text=f'{product.get("pieces", 0)} pcs | {product.get("product", "")} | {product.get("variation", "")}')
                            links[row] = output, source["path"]
                        if not order.get("items"):
                            tree.insert(item, "end", text="Item details were not recorded for this earlier document.")
            selected()
        def chosen():
            return links.get(tree.selection()[0]) if tree.selection() else None
        def selected(event=None):
            value = chosen()
            open_button.state(["!disabled"] if value else ["disabled"])
            source_button.state(["!disabled"] if value and value[1] else ["disabled"])
        def open_document():
            value = chosen()
            if not value:
                return
            output = value[0]
            if output.get("status") != "active":
                messagebox.showinfo("Earlier document", "This document is " + output.get("status", "superseded") + ". Use its replacement or generate another copy from the main screen.", parent=dialog)
                return
            self.open_path(self.output_folder / output["filename"])
        def open_source():
            value = chosen()
            if value and value[1]:
                self.open_path(self.date_folder / value[1])
        open_button = ttk.Button(actions, text="Open document", command=open_document)
        open_button.pack(side="left", padx=(0, 8))
        source_button = ttk.Button(actions, text="Open source PDF", command=open_source)
        source_button.pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Refresh list", command=populate).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Open processing log", command=lambda: self.open_path(self.output_folder / ".PackingData" / "Processed_Files.txt")).pack(side="left")
        ttk.Button(actions, text="Close", command=dialog.destroy).pack(side="right")
        tree.bind("<<TreeviewSelect>>", selected)
        dialog.bind("<Escape>", lambda event: dialog.destroy())
        populate()
        dialog.grab_set()

    def show_help(self):
        text = ("Choose the folder that contains your dated packing folders, then select the packing date. You can also choose a YYYYMMDD folder directly.\n\n"
                "New orders are checked automatically. Tick or untick Use beside a PDF. The Use heading selects all PDFs that are ready.\n\n"
                "For Shopee, click the Courier cell to choose Auto, SPX or FLASH. Auto reads the nearest courier folder first, then the filename.\n\n"
                "Choose Generate Packing Lists to create A4 PDFs. Finished documents go into 000 Output inside the selected date. Existing documents are never overwritten.\n\n"
                "If repeated orders are found, No generates only orders without an existing matching output. The continuation link includes earlier generated orders. Enter and Escape choose No.\n\n"
                "Select a PDF and choose File details to inspect it. Changed files must be reviewed before they can be included.\n\n"
                "Products / Define Names lets you link stable SKUs to the Product Name and Size used in your inventory. You can also inspect the Ecommerce Name and Ecommerce Variation underneath each SKU. Review rows unlock automatically; completed products require explicit unlocking.\n\n"
                "The flashing Review button highlights products in these PDFs with missing SKUs or standard names. You can choose Later and still generate; blank standard fields use the ecommerce text. Correct missing seller SKUs on the website and refresh with updated PDFs.\n\n"
                "Completed documents shows the PDFs, orders and products recorded in each earlier document. Open Output Folder opens all generated documents together.\n\n"
                "Export Item Sales creates an Excel quantity report for the entire selected date, regardless of the Use checkboxes. It does not mark orders processed.")
        self._text_dialog("Packing List Manager help", "Your daily packing workflow", text)

    def close(self):
        if self.busy:
            messagebox.showinfo("Working", "Please wait for the current scan, export or batch to finish.", parent=self)
            return
        self.executor.shutdown(wait=False, cancel_futures=True)
        if getattr(self, "_poll_id", None):
            self.after_cancel(self._poll_id)
        for timer in (self._attention_timer, self._attention_prompt_timer):
            if timer:
                self.after_cancel(timer)
        self.destroy()


def ui_self_test(report_path):
    """Exercise UI selection and the async repeat-No flow without real file writes."""
    import time
    rows = [
        dict(path="SPX/first.pdf", platform="SPX", status="READY", is_shopee=True,
             courier="Auto", sha256="first", orders=2, new_orders=2, message="2 new orders ready",
             order_totals=[{"key": "SPX:1", "pieces": 2}, {"key": "SPX:2", "pieces": 1}]),
        dict(path="SPX/shared.pdf", platform="SPX", status="DUPLICATE", is_shopee=True,
             courier="FLASH", sha256="shared", orders=1, message="Process again",
             order_totals=[{"key": "SPX:1", "pieces": 2}]),
        dict(path="Lazada/lazada.pdf", platform="LAZADA", status="READY", is_shopee=False,
             sha256="lazada", orders=1, new_orders=1, message="1 new order ready",
             order_totals=[{"key": "LAZADA:1", "pieces": 1}]),
        dict(path="SPX/error.pdf", platform="SPX", status="ERROR", is_shopee=True,
             sha256="error", orders=0, message="Missing continuation section", order_totals=[]),
        dict(path="SPX/changed.pdf", platform="SPX", status="REVIEW", is_shopee=True,
             sha256="changed", orders=1, requires_revision=True, message="Review changed file",
             order_totals=[{"key": "SPX:3", "pieces": 1}]),
    ]
    calls, completed = [], []
    def fake_engine(request):
        calls.append(request)
        if request["action"] == "scan":
            return {"ok": True, "entries": rows, "outputs": []}
        if request["action"] == "generate":
            if "repeat_action" not in request:
                return {"ok": True, "confirmation_required": True, "repeat_count": 1, "repeat_token": "selection-format-token"}
            return {"ok": True, "outputs": [], "message": "No new orders. No document created."}
        raise AssertionError("Unexpected action: " + request["action"])
    app = PackingApp(engine_handler=fake_engine, auto_scan=False)
    def wait_until(predicate):
        until = time.monotonic() + 5
        while not predicate() and time.monotonic() < until:
            app.update()
            time.sleep(.01)
        assert predicate(), "UI action did not finish"
    original_warning = globals()["RepeatWarning"]
    try:
        app.scan()
        wait_until(lambda: len(app.entries) == 5 and not app.busy)
        assert app.checked == {"SPX/first.pdf", "Lazada/lazada.pdf"}
        app.toggle_all()
        assert app.checked == {"SPX/first.pdf", "SPX/shared.pdf", "Lazada/lazada.pdf"}
        assert selection_counts(rows, app.checked) == {"SPX": {"pdfs": 2, "orders": 2}, "LAZADA": {"pdfs": 1, "orders": 1}}
        assert "pieces" not in app.totals_var.get().lower()
        assert app.tree.set("2", "courier") == "—"
        app.toggle_all()
        assert not app.checked and app.generate_button.instate(["disabled"])
        app.revisions.add("SPX/changed.pdf")
        app.toggle_all()
        request_rows = generation_selection(rows, app.checked, app.revisions)
        assert len(request_rows) == 4 and request_rows[-1]["revision"]
        assert request_rows[1]["courier"] == "FLASH"
        app.geometry("1100x688")
        app.update()
        assert app.tree.winfo_height() >= 120
        assert app.generate_button.winfo_rooty() + app.generate_button.winfo_height() < app.winfo_rooty() + app.winfo_height()
        warning = original_warning(app, 2)
        warning.update()
        assert warning.result == "skip"
        assert warning.no_button.winfo_width() == warning.continue_link.winfo_width()
        assert warning.no_button.cget("font") == warning.continue_link.cget("font")
        assert warning.bind("<Return>") and warning.bind("<Escape>")
        warning.choose("skip")
        class AutomaticNo(original_warning):
            def __init__(self, parent, count):
                super().__init__(parent, count)
                self.after(20, lambda: self.choose("skip"))
        globals()["RepeatWarning"] = AutomaticNo
        app._generation_done = lambda result, made: completed.append((result, made))
        app.generate()
        wait_until(lambda: bool(completed))
        requests = [request for request in calls if request["action"] == "generate"]
        assert len(requests) == 2
        assert requests[-1]["repeat_action"] == "skip"
        assert requests[-1]["repeat_confirmation"] == "selection-format-token"
        assert requests[-1]["selected"] == requests[0]["selected"]
        assert requests[-1]["output_format"] == "pdf"
        assert completed[-1][1] == []
        report = {"ok": True, "selection_and_dedup": True, "review_gating": True,
                  "courier_preserved": True, "responsive_layout": True,
                  "repeat_no_default": True, "async_repeat_skip": True,
                  "no_empty_output": True, "source_files_written": False}
    finally:
        globals()["RepeatWarning"] = original_warning
        app.close()
    report_path = Path(report_path)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    return report


def main():
    parser = argparse.ArgumentParser(description="Packing List Manager 2.0.0 Python desktop app")
    parser.add_argument("--folder", help="Main packing folder or YYYYMMDD date folder")
    parser.add_argument("--data-dir", type=Path, default=DATA_DIR, help="Independent application data folder")
    parser.add_argument("--self-test", type=Path, metavar="REPORT_JSON", help="Run isolated UI checks and save their report")
    args = parser.parse_args()
    if args.self_test:
        print(json.dumps(ui_self_test(args.self_test)))
        return
    try:
        instance = _single_instance()
    except RuntimeError as exc:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Already open", str(exc), parent=root)
        root.destroy()
        return
    app = PackingApp(data_dir=args.data_dir, source_folder=args.folder)
    try:
        app.mainloop()
    finally:
        if instance:
            import ctypes
            kernel = ctypes.WinDLL("kernel32")
            kernel.CloseHandle.argtypes = (ctypes.c_void_p,)
            kernel.CloseHandle(instance)


if __name__ == "__main__":
    main()
