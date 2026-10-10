# Packing List Manager — Python source, V2.0.0

Start with `../Start Packing List Manager.cmd`, or `python src/app.py`. V2 uses Python/Tkinter for its desktop interface and retains the local Python engine.

## Source map

- `src/app.py`: main workflow, asynchronous engine calls, repeat and changed-source dialogs, courier choices and output history.
- `src/products_ui.py`: SKU-first product editor, ecommerce details, locks, sorting, linked definitions and Excel sessions.
- `src/names.py`: versioned catalog, SKU/name migration, independent Product Name/Size fallback, locked definitions and Excel round trips.
- `src/sku_registry.py`: platform-scoped seller aliases and Shopee-authoritative catalog evidence.
- `src/extractors.py`: PDF parsing, raw seller SKUs, continuation validation and unchanged source fingerprints.
- `src/engine.py`: date/history operations, repeat decisions, PDF-only generation and immutable audit snapshots.
- `src/lazada_history.py`: verified tracking links for earlier Lazada outputs without altering their audit snapshots.
- `src/documents.py`: A4 layouts, platform colors, prices, buyer comments and picking summaries.
- `src/sales.py`: full-date quantity Excel snapshot, with source prices when available.
- `src/assets/verify_word.ps1`: actual Word pagination verification and PDF export.
- `src/products_excel.ps1`: temporary Excel workbook lifecycle without a Python COM package.

`ProductNames.json`, `SkuCatalog.json` and `Settings.json` belong in `../Packing List Manager`. They are private and excluded from Git. Schema 3 adds SKU identity and ecommerce provenance without discarding original links, IDs, row locks or saved inventory names. Existing immutable history remains with each dated folder.

An unknown supplied SKU must not silently borrow a different SKU because titles match. Missing or generic size-only codes retain legacy name/variation fallback. Conflicting finished definitions remain visible for review and cannot silently determine output names.

`qa/import_sku_catalog.py` regenerates the private catalog from approved ecommerce workbooks and historical Lazada mappings. It refuses to overwrite an existing destination. Reference business data and QA results are deliberately not tracked.

## Checks

Run `python -m unittest discover -s qa -p "test_*.py" -v`. Inherited private-data checks accept `PACKING_TEST_DATA` pointing at the parent V1 workspace's `PackingListApp/qa/names-regression-fixtures`. `python src/app.py --self-test qa/ui-verification.json` checks the desktop workflow. `qa/verify_reference.py` uses an isolated dated copy to check immutable source records, SKU rename stability, real Office PDF layout and full-date sales totals.

`python qa/verify_ui_layout.py` checks the Station Board interface at normal and 125% scaling, verifies checkbox/courier routing and animated loading, and saves local screenshots. It uses synthetic rows and does not write product settings or packing outputs.

Keep work local. The checkpoint is on `codex/v2.0.0`; a verified bundle belongs in the workspace's `Git/v2.0.0`. No executable is built at this stage.
