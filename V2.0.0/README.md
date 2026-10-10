# Packing List Manager 2.0.0

This is the approved V2 release, with Python source preserved. The **Release** folder contains **Packing List Manager Setup 2.0.0.exe** and a portable ZIP. Run the installer for Windows shortcuts and an uninstaller, or extract the entire ZIP and open **Packing List Manager.exe**. The installed V1 app remains in its original folder.

For the source preview, double-click **Start Packing List Manager.cmd**. Its existing settings can point to `Test Data`, with dated copies of the reference PDFs. Packaged copies start without a preset source folder: choose your usual packing folder and date. They include the approved product definitions and SKU catalog, with no test orders, outputs or history.

The main screen provides the date/folder, PDF selections, courier choices and **Generate packing lists**. **Products** opens the SKU catalog; unlock it to edit **Product Name** and **Size**. **Ecommerce details** shows the marketplace names, variations and seller SKUs behind each definition. The existing packing, repeat-order, changed-file review, history and sales workflows are retained.

The approved Station Board design uses graphite gray and off-white, strong dividers, a prominent packing date, and document actions on the right. A full-width animated loading bar appears while work is running. Products uses the matching theme; an asterisk and yellow highlight identify unsaved rows. The interface is checked at 125% scaling in an 1100×688 window. This design update keeps version **2.0.0**.

Click the **Use** heading to select all eligible PDFs, including **Process again** rows; click again to clear them. Completed Define Names rows open locked; products needing review open unlocked. You can still view Ecommerce Details and copy with **Ctrl+C**, or right-click a row to copy its SKU, Product Name or Size. Editing completed rows requires explicit unlocking.

Shopee is the master SKU source. Approved earlier Lazada SKUs remain linked so older packing lists still work. A changed listing title with the same SKU keeps the same inventory Product Name and Size. Older PDFs without a usable SKU can still use their existing name/variation definitions; the app does not invent a SKU or combine generic size-only codes across products.

Flash packing headers and picking-summary headers use pastel yellow. SPX remains orange; Lazada and TikTok share blue outputs. All outputs retain A4 paper, whole-order tables, handwritten staff fields and quantity picking summaries.

Lazada shipments use **Order Number + Tracking Number**. One order with different tracking numbers prints as separate shipment tables, each showing its tracking number. Counts distinguish orders from shipments, and sales exports include every shipment's quantities once. Earlier outputs remain covered when unchanged source evidence proves their tracking number; uncertain earlier records require review.

## Folders

- `PackingListApp/src`: Python application and engine source.
- `PackingListApp/src/assets`: document templates and Windows Office helpers.
- `Packing List Manager`: this V2 copy of private settings, product definitions and SKU catalog. Keep it with the source folder and back it up separately from Git.
- `Reference`: your supplied test PDFs, preserved as provided.
- `Test Data/20261010`: ready-to-use copies for trying the V2 app.
- `PackingListApp/qa`: automated checks and local test results. Generated test documents remain inside this folder, separate from the reference PDFs and production dates.
- `Release`: packaged application, installer, portable ZIP and file hashes. These contain private catalog copies and are excluded from Git.

The launcher uses this computer's available Python runtime, or a `.venv` here, then falls back to the Python 3.12 launcher. On another Windows computer, install Python 3.12 (including Tkinter) and run `py -3.12 -m pip install -r PackingListApp/requirements.txt`. Microsoft Word desktop is required for verified PDF export. Microsoft Excel desktop is required only for the live product-editing workbook.

For development, run `python PackingListApp/src/app.py`. Optional `--data-dir` and `--folder` arguments support isolated testing. The packaged application needs no separately installed Python; Microsoft Word desktop is still required for verified PDF export.

The installer defaults to your account's `AppData/Local/Programs/Packing List Manager` folder. It creates **Packing List Manager V2** shortcuts, keeping earlier V1 shortcuts separate. Reinstall keeps existing product definitions, SKU catalogs and settings. Uninstall removes only known program files and retains your saved data and dated outputs.
