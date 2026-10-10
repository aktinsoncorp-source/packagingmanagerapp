# V2.0.0 release checkpoint

Approved branch: `codex/v2.0.0`, merged into `main` for release. V2 lives under `V2.0.0`; the root V1 Python/C# sources and installed application are retained. The desktop interface remains Python/Tkinter, packaged with a private CPython/Tk runtime and a small Windows launcher. Windows Office helpers verify PDF export and support live Excel editing. `Release` contains the installer and portable package; the Python source remains editable.

Release checks are recorded in `PackingListApp/qa/release-2.0.0-regression.txt` (281 passing cases), `packaged-v2-ui.json`, and `installer-verification.txt`. The installer is exercised in a separate QA folder, including preservation of edited catalogs during reinstall/uninstall and real Lazada split-shipment PDF generation. Earlier source-preview verification remains below.

The private registry contains 1,384 master SKU records and 2,449 marketplace seller aliases from the approved local catalogs. Shopee supplies the current descriptions. The V2 definition copy retains every pre-existing definition ID, lock and inventory alternative; the three user-approved conflicting values were resolved in this copy. Platform-scoped unknown codes cannot silently become a different marketplace's known SKU. Older blank or generic size-only codes retain legacy matching.

Source fingerprints remain compatible with earlier history. Additional SKU evidence catches contradictory copies and changed historical SKUs before packing or sales deduplication. Approved old/new aliases are equivalent. When overlapping copies contain different amounts of SKU information, the app uses a real sufficiently complete source copy; it does not invent missing identifiers.

## Verification evidence

- `PackingListApp/qa/unit-tests-final.log`: all 221 regression tests passed.
- `PackingListApp/qa/ui-final-verification.json`: main-screen selection, deduplication, courier, review, PDF-only and repeat-No checks.
- `PackingListApp/qa/ui-preview/products-responsive-check.json`: full 1,278-row test catalog at 125% scaling and 1100x688, with all controls and primary columns visible.
- `PackingListApp/qa/v2-reference-v1-parity.json`: exact V1/V2 extraction parity across 15 reference PDFs, 190 orders and 415 pieces, including prices, comments and hashes.
- `PackingListApp/qa/reference-verification-r9fp03t5/verification.json`: real Word PDF verification, source preservation, saved-name stability after ecommerce renaming, immutable snapshots, repeated-order No and full-date sales reconciliation.
- `PackingListApp/qa/reference-verification-r9fp03t5/final-engine-verification.json`: final SKU-consistency integration against the completed test batch.

All 19 pages were rendered with Poppler and visually checked. Flash produces 4 pages, SPX 13 and Lazada/TikTok 2 for these references. Every order remains together; summaries start on a new page. Flash is pastel yellow, SPX orange, Lazada/TikTok blue, with more visible printed table borders. An explicit tiny final paragraph prevents Word adding an empty trailing page after a full summary.

The Products editor has focused model and actual Tk checks for metadata, locks, failed saves, sorting and responsive layout. A protected desktop Excel export/save/import was tested with linked originals and SKU metadata intact.

## Local data and recovery

The first launch points to `Test Data`, containing unchanged copies of the reference PDFs under `20261010`. Generated QA outputs are separate from these ready-to-use copies. Choose the normal shop folder in the app when moving to live use.

The Git bundle in the workspace's `Git/v2.0.0` contains source and sanitized templates. Private catalogs, settings, references, test PDFs and generated outputs are excluded. Preserve `Packing List Manager/ProductNames.json` and `SkuCatalog.json` separately when moving or restoring V2. The original V1 application and earlier `Git/vVERSION` bundles remain in place.
