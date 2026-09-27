# Packing List Manager

The employee application is in `../Packing List Manager/`. Launch `Packing List Manager.exe`.
The whole application folder is portable. It includes its Python runtime. Windows x64 and desktop Microsoft Word are required.

This directory contains development files, sanitized document templates, build dependencies and QA evidence.

## Components

- `src/PackingListManager.cs`: Windows Forms interface; invokes the packaged engine with JSON requests.
- `src/extractors.py`: layout-specific readers; quantities are validated without OCR or AI services.
- `src/documents.py`: clones the approved A4 order tables and appends a per-document product picking summary after a page break.
- `src/picking.py`: aggregate quantities only for matching original product names and variation text. No aliases, shortening, or inferred variation sizes are applied. The earlier alias catalog is retained as historical QA data and is no longer read by the application.
- `src/engine.py`: date-specific file and order history, explicit revisions, recoverable commits and duplicate handling.
- `src/assets/verify_word.ps1`: verifies A4 paper size and keeps every complete order table on one page using Word's actual pagination.
- `src/assets/orange.docx` and `blue.docx`: sanitized templates distilled from the approved documents.
- `Employee Guide.txt`: employee instructions, also distributed with the application.

The courier dropdown has been removed. Add PDFs automatically recognizes Shopee SPX/FLASH from the source filename or folder path; an ambiguous path asks the employee to label or move the PDF. Existing saved courier settings are ignored. Lazada/TikTok detection still uses PDF contents.

The Use header contains a select-all checkbox. Clicking it selects every compatible PDF (including processed/repeated files), and clicking again clears all selections. Error rows and unresolved changed-file reviews are excluded. A partial selection displays an indeterminate check. The header stays aligned when scrolling or resizing. `--self-test` checks all/none/partial selection with error and review fixtures.

The main window shows one PDF selection table. The former lower order-details panel has been removed and the table uses the available height. Full status messages are available on hover; changed-file comparisons remain in the Review changed file dialog.

Version 1.3 checks: six naming/summary tests passed; the installed executable preserved exact original names and variations in a seven-order test. Word verified all three full-day layouts: SPX 11 pages (two summary pages), FLASH 4, Lazada/TikTok 2; totals remain 228, 76 and 14 pieces. All changed summary pages were visually checked. The installed UI passed all/none/partial selection checks and excluded error/review rows. Evidence: `qa/select-all-ui-verification.json`, `qa/original-names-packaged-verification.json`, and `qa/original-names-layout/verification.json`.

## Regeneration and missing Word files

Processed and repeated PDF rows remain selectable and start unchecked. New PDFs and PDFs with uncovered orders due to missing Word outputs start checked. Generate includes all orders from the selected PDFs, deduplicating identical order IDs within the run. Another generation uses one above the highest suffix physically present in that output group and date, starting at 1 when none remain. Deleted trailing numbers can be reused; lower gaps are not filled. Regeneration keeps surviving earlier copies active; a reviewed revision supersedes all older affected copies.

Each output stores immutable `sources` (PDF path, SHA256, included order keys), `order_fingerprints`, `order_snapshots`, and a unique `record_id`. Before a deleted filename is reused, its old record is moved to `archived_outputs`; it cannot count as coverage for the replacement. The text log retains both generations and their record IDs. The readable log also lists selected PDFs and linked order IDs. Scan checks output existence, ignores superseded order versions, and finds any surviving current copy before marking an order missing. Renamed or moved Word files count as missing at their recorded paths. A partial PDF is regenerated in full when selected. Changed source contents still require explicit review.

Scan migrates earlier log-only batches and older app records into these associations while retaining their original Word files and text log. It saves the JSON associations so subsequent deletion of an old Word file can be detected. If a legacy Word file was already deleted, its source paths are recovered from the text log; no old order contents are invented. Deleting app-era JSON history is still an error because it removes the source/version record.

Regeneration checks: 28 automated tests passed, including overlap, duplicate selection, surviving copies, missing outputs, legacy migration, changed-file review and revision invalidation. Installed executable/UI evidence is saved in `qa/regeneration-*-verification.json`.

Summary layout templates (`summary_orange.docx`, `summary_blue.docx`) follow the September 26 approved summaries. Each summary totals only the orders included in its generated file, including a later batch or explicit revision. A long summary repeats its header on additional pages, keeps individual rows together, and keeps the last item with the grand total. Word counts only order tables as orders. Legacy history accepts and validates the separate summary table.

## Completed documents

Only existing files are returned to the Completed documents view. They are grouped SPX, FLASH, LAZADA_TIKTOK and sorted by integer suffix within each group. Missing records stay internal for regeneration tracking, and are excluded from the visible list. The TreeView expands document -> source PDF -> order -> original product/variation/pieces. Included data comes from the saved output snapshot, not from the source PDF's current contents. The list filters existence again when opened/refreshed. The main PDF selection screen stays a single table.

## Build

Use Python 3.12 with pdfplumber, python-docx and PyInstaller 6.22.3. The current build uses Codex's bundled Python for PDF and document dependencies and project-local PyInstaller in `.builddeps`.

Run `build.ps1`. It packages the engine and compiles the Windows Forms interface using the .NET Framework C# compiler. The program runs as the current user and has no network calls.

## Verification completed September 26 2026

Version 1.4: all 35 regression checks passed, including suffix reuse after deletion, missing completed-file filtering, numeric platform sorting, saved document contents, and protection against overwriting a file created during generation. The installed app scanned all nine production PDFs successfully; selecting processed PDFs and Select all worked. The Completed documents screen was visually checked with SPX, FLASH, and Lazada/TikTok groups and expandable source/order contents. Reports: `qa/completed-documents-tests.log`, `qa/completed-documents-ui-verification.json`, `qa/completed-documents-tree-verification.json`; preview: `qa/app-completed-documents.png`. The final installed Word-generation smoke test was interrupted by a computer crash and has not completed. No production packing documents were regenerated during these checks.

Summary update: all 18 regression checks passed. The rebuilt executable reproduced the approved per-file summaries (SPX 228 pieces/10 pages, FLASH 76 pieces/4 pages, Lazada/TikTok 14 pieces/2 pages). Word verified every order stays on one A4 page and each summary starts on a new page. A 75-product synthetic summary continued correctly across two pages with repeated headers. Updated UI loaded and scanned all nine production files as already processed; no production batch was regenerated. QA reports are `qa/summary-*-verification.json` and `qa/packaged-smoke/layout-verification.json`. Word exported the preview PDFs and Poppler rendered the visual checks; the skill renderer could not run because LibreOffice is not installed.

Original version checks:

- 13 workflow tests passed, including duplicate files, changed files, revisions, removed orders, history recovery, import routing, next suffixes, rejected input and no completed records after a failed Word check.
- 27 supplied PDFs across September 24, 25 and 26 parsed with complete Order ID coverage.
- All 139 current orders, their products, variations and 318 pieces matched the approved September 26 Word documents.
- A full run of the packaged executable produced SPX 9 pages, FLASH 3 pages and combined Lazada/TikTok 1 page, all A4, with no split order tables. A second scan marked all nine source files completed.
- All generated pages were rendered and visually reviewed. Poppler was used for the final raster checks because PDFium sometimes omitted very thin underscore glyphs at preview scale.
- The packaged Windows form loaded successfully and scanned the current dated folder.

Run `python -m unittest discover -s qa -p "test_*.py" -v` for regression tests (35 checks as of version 1.4). The document-layout validator is replaced only in the unit tests that test commit behavior; `qa/packaged_smoke.py` uses the actual executable and Microsoft Word. `qa/check_summary_layout.py` verifies an additional multipage summary and exports QA PDFs.
The smoke test requires a fresh `qa/packaged-smoke` directory to avoid overwriting test batches.

## Current scope

Duplicate history is per dated folder. Scanned PDFs, unknown layouts, repeated order labels inside one PDF, multipackage Shopee labels and continued order pages stop for review. A platform layout change may require updating the corresponding parser.

The employee workflow is one operator on one Windows computer. Processing history belongs with its dated source and output folder. The app reads the earlier human processing log and verifies it against completed documents when importing pre-app history; missing app-era JSON history is an error.
