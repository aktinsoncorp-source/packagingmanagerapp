"""Check the native Station Board layout with isolated, synthetic PDF rows."""
import argparse
from contextlib import nullcontext
import ctypes
import json
from pathlib import Path
import sys
import time

from PIL import ImageGrab

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from app import PackingApp
from products_ui import ProductsDialog


def capture(window, destination):
    window.update()
    handle = ctypes.windll.user32.GetParent(window.winfo_id())
    ImageGrab.grab(window=handle).save(destination)


def verify(output):
    output.mkdir(parents=True, exist_ok=True)
    reports = []
    for scale, size in ((1.0, "1220x850"), (1.0, "1100x688"), (1.25, "1100x688")):
        # auto_scan is off and no save commands run; this path stays unwritten.
        with nullcontext(output / "isolated-ui-data") as temporary:
            app = PackingApp(data_dir=temporary, auto_scan=False, ui_scale=scale)
            try:
                app.geometry(size)
                app.folder_var.set("Packing / 20261010")
                app.entries = [
                    dict(path=f"{platform}/packing-list-{index+1:02}.pdf", platform=platform,
                         status="READY", is_shopee=platform in ("SPX", "FLASH"), courier="Auto",
                         orders=count, new_orders=count, message="Ready",
                         order_totals=[{"key": f"{platform}:{i}", "pieces": 1} for i in range(count)])
                    for index, (platform, count) in enumerate((("SPX", 8), ("FLASH", 4), ("LAZADA", 6)))
                ]
                app.checked = {row["path"] for row in app.entries[:2]}
                app._populate_files()
                app.summary_var.set("Packing PDFs · 190 new orders · 15 need attention")
                app.update()
                windowsize = (app.winfo_width(), app.winfo_height())
                for button in [*app.buttons, app.folder_entry, app.tree, app.status_label]:
                    x = button.winfo_rootx()-app.winfo_rootx()
                    y = button.winfo_rooty()-app.winfo_rooty()
                    label = button.cget("text") if "text" in button.keys() else button.winfo_class()
                    assert button.winfo_viewable(), (scale, size, label, "not visible")
                    assert 0 <= x and x+button.winfo_width() <= windowsize[0], (scale, size, label, "horizontal clipping")
                    assert 0 <= y and y+button.winfo_height() <= windowsize[1], (scale, size, label, "vertical clipping")
                    parent = button.master
                    while parent is not app:
                        assert button.winfo_rootx() >= parent.winfo_rootx(), (scale, label, "clipped left by parent")
                        assert button.winfo_rooty() >= parent.winfo_rooty(), (scale, label, "clipped above by parent")
                        assert button.winfo_rootx()+button.winfo_width() <= parent.winfo_rootx()+parent.winfo_width(), (scale, label, "clipped right by parent")
                        assert button.winfo_rooty()+button.winfo_height() <= parent.winfo_rooty()+parent.winfo_height(), (scale, label, "clipped below by parent")
                        parent = parent.master
                assert app.tree.winfo_height() >= 150, (scale, size, "PDF list too short", app.tree.winfo_height())
                assert app.tree.xview() == (0.0, 1.0), (scale, size, "Primary columns do not fit")

                # Use/courier hit targets must stay correctly routed after moving the tree.
                app.tree.see("0")
                app.update()
                x, y, width, height = app.tree.bbox("0", "#0")
                app.tree.event_generate("<Button-1>", x=x+width//2, y=y+height//2, time=1000)
                app.tree.event_generate("<ButtonRelease-1>", x=x+width//2, y=y+height//2, time=1010)
                app.update()
                assert app.entries[0]["path"] not in app.checked
                app.tree.event_generate("<Button-1>", x=x+width//2, y=y+height//2, time=2000)
                app.tree.event_generate("<ButtonRelease-1>", x=x+width//2, y=y+height//2, time=2010)
                app.update()
                assert app.entries[0]["path"] in app.checked
                courier_calls = []
                app.choose_courier = lambda *args: courier_calls.append(args)
                x, y, width, height = app.tree.bbox("0", "courier")
                app.tree.event_generate("<Button-1>", x=x+width//2, y=y+height//2, time=3000)
                app.update()
                assert courier_calls and courier_calls[-1][0] == "0"
                assert app.tree.set("2", "courier") == "—"

                app.summary_var.set("Packing PDFs · 18 new orders")
                app.status_var.set("Ready to generate packing lists")
                key = f"station-{size}-{int(scale*100)}"
                capture(app, output / f"{key}.png")
                app._set_busy(True, "Reading PDFs and checking completed documents...")
                app.update()
                assert app.progress.winfo_viewable()
                assert all(button.instate(["disabled"]) for button in app.buttons)
                start = app.progress.cget("value")
                until = time.monotonic()+.2
                while time.monotonic() < until:
                    app.update()
                    time.sleep(.01)
                assert app.progress.cget("value") != start, "Loading bar is not moving"
                assert app.progress.winfo_rooty()+app.progress.winfo_height() <= app.winfo_rooty()+app.winfo_height()
                assert app.tree.winfo_height() >= 150
                assert app.output_button.winfo_rooty()+app.output_button.winfo_height() <= app.action_rail.winfo_rooty()+app.action_rail.winfo_height(), "Output button clipped during loading"
                capture(app, output / f"{key}-loading.png")
                app.status_var.set("Needs attention: " + ("This source file needs review before its orders can be included. " * 3)[:190])
                app.summary_var.set("Packing PDFs · 190 new orders · 15 need attention")
                app.update()
                assert app.output_button.winfo_rooty()+app.output_button.winfo_height() <= app.action_rail.winfo_rooty()+app.action_rail.winfo_height(), "Output button clipped with long status"
                assert app.tree.winfo_height() >= 150, (scale, size, "Long status hides PDF rows", app.tree.winfo_height(), app.files_summary_label.winfo_height(), app.status_label.winfo_height())
                app._set_busy(False, "Ready")
                app.update()
                assert not app.progress.winfo_viewable()

                if scale == 1.25:
                    dialog = ProductsDialog(app, {"revision": 1, "rows": [dict(
                        id="white-musk", sku="PERF-WHITE-MUSK-W-30ML", original_product="WHITE MUSK",
                        original_variation="30 ML", alternative_product="WHITE MUSK (W) P.O.",
                        alternative_variation="30 ML", locked=True)]}, app.call_engine)
                    dialog.geometry(size)
                    app.update()
                    capture(dialog, output / "station-products-125.png")
                    dialog._destroy()
                reports.append(dict(scale=scale, size=size, table_height=app.tree.winfo_height(),
                                    xview=app.tree.xview(), controls_visible=True,
                                    click_routing=True, animated_loading=True, long_status_visible=True))
            finally:
                if app.busy:
                    app._set_busy(False, "Ready")
                app.close()
    (output / "station-layout-verification.json").write_text(json.dumps(reports, indent=2), encoding="utf-8")
    return reports


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=Path(__file__).resolve().parent)
    print(json.dumps(verify(parser.parse_args().output)))
