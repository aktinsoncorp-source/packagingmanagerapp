# Packing List Manager

The employee application is in `../Packing List Manager/`. Launch `Packing List Manager.exe`.
The whole application folder is portable. It includes its Python runtime. Windows x64 and desktop Microsoft Word are required.

This directory contains development files, sanitized document templates, build dependencies and QA evidence.

## Components

- `src/PackingListManager.cs`: Windows Forms interface; invokes the packaged engine with JSON requests.
- `src/extractors.py`: layout-specific readers; quantities are validated without OCR or AI services.
- `src/documents.py`: clones the approved A4 order tables and appends a per-document product picking summary after a page break.
- `src/picking.py`: aggregate quantities for matching final product names and variations. `src/names.py` applies user-defined alternatives independently before building a document; blank fields retain the originals. The earlier fixed alias catalog is historical data and is not read by the application.
- `src/NameDefinitionsForm.cs` and `src/NameExcelSession.cs`: editable name catalog and a temporary desktop Excel session with alternating Export/Import buttons.
- `src/engine.py`: date-specific file and order history, explicit revisions, recoverable commits and duplicate handling.
- `src/assets/verify_word.ps1`: verifies A4 paper size and keeps every complete order table on one page using Word's actual pagination.
- `src/assets/orange.docx` and `blue.docx`: sanitized templates distilled from the approved documents.
- `Employee Guide.txt`: employee instructions, also distributed with the application.

Version 1.7 adds an individual Auto/SPX/FLASH courier cell for each Shopee PDF. Auto prioritizes the nearest exact SPX/FLASH folder, then a single filename label, then unambiguous remaining path labels. Explicit choices override Auto and persist per source path/hash in `000 Output/Courier_Selections.json`; changed contents revert to Auto. Lazada/TikTok courier cells are read-only dashes. Content recognition includes PDFs anywhere in the dated folder, excluding `000 Output`. Unassigned Shopee PDFs retain real counts and need a courier choice before generation. Add PDFs accepts them into the dated folder root.

The Use header contains a select-all checkbox. Clicking it selects every compatible PDF (including processed/repeated files), and clicking again clears all selections. Error rows and unresolved changed-file reviews are excluded. A partial selection displays an indeterminate check. The header stays aligned when scrolling or resizing. `--self-test` checks all/none/partial selection with error and review fixtures.

The main window shows one PDF selection table. The former lower order-details panel has been removed and the table uses the available height. Full status messages are available on hover; changed-file comparisons remain in the Review changed file dialog.

Version 1.3 checks: six naming/summary tests passed; the installed executable preserved exact original names and variations in a seven-order test. Word verified all three full-day layouts: SPX 11 pages (two summary pages), FLASH 4, Lazada/TikTok 2; totals remain 228, 76 and 14 pieces. All changed summary pages were visually checked. The installed UI passed all/none/partial selection checks and excluded error/review rows. Evidence: `qa/select-all-ui-verification.json`, `qa/original-names-packaged-verification.json`, and `qa/original-names-layout/verification.json`.

## Regeneration and missing Word files

Processed and repeated PDF rows remain selectable and start unchecked. New PDFs and PDFs with uncovered orders due to missing Word outputs start checked. Generate includes all orders from the selected PDFs, deduplicating identical order IDs within the run. Another generation uses one above the highest suffix physically present across PDF and Word files in that output group and date, starting at 1 when none remain. Deleted trailing numbers can be reused; lower gaps are not filled. Regeneration keeps surviving earlier copies active; a reviewed revision supersedes all older affected copies.

Each output stores immutable `sources` (PDF path, SHA256, included order keys), `order_fingerprints`, `order_snapshots`, and a unique `record_id`. Before a deleted filename is reused, its old record is moved to `archived_outputs`; it cannot count as coverage for the replacement. The text log retains both generations and their record IDs. The readable log also lists selected PDFs and linked order IDs. Scan checks output existence, ignores superseded order versions, and finds any surviving current copy before marking an order missing. Renamed or moved generated files count as missing at their recorded paths. A partial PDF is regenerated in full when selected. Changed source contents still require explicit review.

Scan migrates earlier log-only batches and older app records into these associations while retaining their original Word files and text log. It saves the JSON associations so subsequent deletion of an old Word file can be detected. If a legacy Word file was already deleted, its source paths are recovered from the text log; no old order contents are invented. Deleting app-era JSON history is still an error because it removes the source/version record.

Regeneration checks: 28 automated tests passed, including overlap, duplicate selection, surviving copies, missing outputs, legacy migration, changed-file review and revision invalidation. Installed executable/UI evidence is saved in `qa/regeneration-*-verification.json`.

Summary layout templates (`summary_orange.docx`, `summary_blue.docx`) follow the September 26 approved summaries. Each summary totals only the orders included in its generated file, including a later batch or explicit revision. A long summary repeats its header on additional pages, keeps individual rows together, and keeps the last item with the grand total. Word counts only order tables as orders. Legacy history accepts and validates the separate summary table.

## Completed documents

Only existing files are returned to the Completed documents view. They are grouped SPX, FLASH, LAZADA_TIKTOK and sorted by integer suffix within each group. Missing records stay internal for regeneration tracking, and are excluded from the visible list. The TreeView expands document -> source PDF -> order -> original product/variation/pieces. Included data comes from the saved output snapshot, not from the source PDF's current contents. The list filters existence again when opened/refreshed. The main PDF selection screen stays a single table.

## Build

Use Python 3.12 with pdfplumber, python-docx, openpyxl and PyInstaller 6.22.3. The current build uses Codex's bundled Python for PDF and document dependencies and project-local PyInstaller in `.builddeps`.

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

Duplicate history is per dated folder. Scanned, unknown, corrupt and textless PDFs are omitted. Recognized packing lists with invalid data remain error rows. Shopee continuation fragments are joined only with matching Order IDs and consecutive printed item numbers; missing, overlapping, duplicated or orphan fragments and multipackage labels stop for review. A platform layout change may require updating the corresponding parser.

The employee workflow is one operator on one Windows computer. Processing history belongs with its dated source and output folder. The app reads the earlier human processing log and verifies it against completed documents when importing pre-app history; missing app-era JSON history is an error.

## Define Name (version 1.5)

Compatible PDF scans remember exact original product/variation pairs in the installation's ProductNames.json. Definitions persist across dates and folders. The four editable columns support adding rows, changing originals and alternatives, search, and deletion. Deletion clears alternatives permanently for that row; a later PDF encounter re-adds the pair with a new ID and blank alternatives. Opening the editor or exporting does not rediscover deleted entries.

Alternatives must start with a Unicode letter or number. Internal spaces and the common symbols - ( ) / % . + , & apostrophe are allowed; control characters, leading spaces, and leading symbols are rejected. Each alternative field falls back independently. Final matching pairs share a picking-summary total; source order fingerprints and raw snapshots remain unchanged. Outputs additionally store immutable printed snapshots and the catalog revision so history is accurate after a later rename.

Excel editing uses a program-owned workbook in a unique Windows temporary folder. Export saves grid edits, opens Excel, disables Export, enables Import, and pauses grid editing. Import requires that exact workbook to remain open, saves unsaved Excel edits, validates a merged preview, atomically saves definitions, closes only the temporary workbook, deletes its file, and resets the buttons. Other workbooks stay open. Closing the workbook or Excel first triggers the requested two-line error and resets the session. Invalid imports retain the workbook and original catalog for correction. Hidden stable IDs allow original-name edits; new rows without IDs are added. Rows omitted from Excel stay in the catalog; deletion is explicit in the grid. Formulas and duplicate/conflicting original pairs are rejected. Catalog writes have a file lock, revision checks, atomic replacement, and a preceding-copy .json.bak.

qa/test_names.py tests storage, fallback, validation, deletion/re-encounter, Excel round trips, summary merging and original/printed history. qa/NamesUiSmoke.cs runs the actual form workflow against the packaged engine and desktop Excel, including manual workbook closure, full Excel exit, invalid input recovery, button states and other-workbook preservation. Compile the harness with the two naming UI source files and the same framework references as the main UI. It takes the installed engine path and a fresh QA directory as arguments.

Private historical test fixtures can be selected with PACKING_TEST_DATA. qa/prepare_local_fixtures.py reconstructs an isolated set from retained QA copies and a supplied shop-data root. Production folders and QA evidence are not tracked in Git. The complete 50-test regression suite and all 18 live Excel UI checks passed during the September 27 update. Installed generation and UI reports are recorded separately.


## Define Name update (version 1.6)

Schema 2 stores each base original pair plus `aliases`, a list of the other
original product/variation pairs linked to the definition. Matching uses uppercased
pairs with trailing spaces removed. All four catalog fields use uppercase; original
PDF Item objects, order snapshots and fingerprints remain unchanged. Single-line
validation rejects control characters and Unicode line/paragraph separators.

`names_prepare` validates and reports matching nonblank alternative pairs without
writing. The form confirms each candidate group on Save and Import. Approved ID
lists are passed to `names_save`; the first selected ID is the base (first displayed
or worksheet order if none selected). Declined groups remain separate. Only matching
nonblank alternatives may merge. Source-pair lookup spans bases and all aliases.
The View column exposes linked originals; search includes them. Removing a definition
removes all its pairs; future PDF encounters re-add them separately with blank aliases.

Excel exports every source pair, with hidden definition and original-pair IDs.
All original fields remain editable after a merge. Omitted rows are retained.
Changes to alternatives in one linked row propagate to its definition; contradictory
new values fail the entire import. Names imports normalize before comparisons.
Schema 1 upgrades under the catalog lock with a preceding backup; no source scan or
re-seeding occurs during migration. User-deleted entries stay deleted until encountered.

`qa/test_name_merging.py` covers merged lookup, yes/no outcomes, preferred base,
case-insensitive matching, trimming, paragraphs, migration, deletion and Excel links.
`qa/NamesUiSmoke.cs` also exercises these behaviors in the real form and desktop Excel.
Use `PYTHONUTF8=1` when running the historical Windows test suite.
The build script accepts `-ReleasePath` for testing a staged build before installation.


## Employee header update (version 1.6.1)

`documents.set_order_header` writes the order identifier, Employee:__________________,
Number and Time in the existing merged header cell. Arial 9.5 pt is retained, with
left tab stops at 185, 340 and 440 pt so the four fields fit the A4 table width.
Only the identifier is bold; the source orange/blue table designs remain intact.
The builder verifies each full header when reopening the created document.
`qa/employee_headers_smoke.py` checks the packaged engine on copied PDFs from all
four platforms, verifies quantities and history parsing, runs real Word pagination,
and checks in exported PDFs that every header field stays on one line without overlap.
Existing completed files are preserved; regenerating produces the updated layout.

## Version 1.7 verification

`qa/test_couriers_continuations.py` covers continuation numbering, folder priority,
courier persistence/reset, source changes, hidden unsupported PDFs and generation.
The full suite has 80 tests. `qa/couriers_smoke.py <release-directory> <source-pdf>`
checks the real Windows dropdown, read-only platform cells, packaged generation,
Word pagination and the reported 127-order/300-piece continuation fixture using
isolated copies. That private PDF is not distributed in Git. Run with PYTHONUTF8=1.

## Export and employee controls (version 1.8)

`handle(generate)` and the Windows export dialog default to PDF. The internal Python
`generate` helper retains its Word default for development callers. The export dialog
resets to PDF every time. Word builds and validates both formats; for PDF, its fixed
format export is checked for the expected page count and the temporary DOCX is removed
before publishing. Only the chosen format is recorded. PDF and DOCX share numeric suffix
allocation, completed-document listing, regeneration and missing-output coverage.

Before building anything, generation checks selected order IDs against all readable
PDFs in the dated folder and earlier recorded orders. Repeats return confirmation_required
and a token tied to selection, repeated IDs and output format. The UI shows the exact
requested warning; Enter/Escape select the larger blue No button. Yes is smaller.
A second generate request carries the token and repeats validation. Duplicates within
one generation remain printed once. This check is scoped to the selected dated folder.

Scan supplies per-order key/piece counts for local UI totals. Checked source PDFs are
counted by platform; order IDs/pieces are deduplicated in the same stable revision-first
order as generation. Only checked platforms appear below the main grid. Main section
and action labels no longer have step numbers. Both document tables use [     ] boxes.

NameDefinitionsForm opens with editLocked=true. Two GDI-drawn padlock icons and tooltips
control editing. Search/linked-original viewing stay available while locked. Save/add/
delete/Excel buttons and editable cells are disabled. Save/Export/Import methods also
reject locked calls. Pending Excel sessions cannot be relocked; unsaved edits must be
saved first. A new form always starts locked. This is an accidental-edit guard without
a password. The older NamesUiSmoke harness now explicitly unlocks before editing and
checks locked entry points, relocking and the Excel-session constraint.

QA: test_exports.py checks format history, shared numbering, missing PDFs, repeat-token
gates and wider checkbox text. ExportDialogsSmoke.cs checks real radio/button defaults.
release18_smoke.py tests packaged PDF and Word generation, warning cancellation before
publication, actual PDF bracket spacing, history, and dynamic main-form totals on copies.

Processed status reports PDF, Word or PDF + Word from surviving nonsuperseded outputs
matching current order fingerprints. Older Word records need no format field because
the extension identifies them. Deletions remove formats on rescan. Partial coverage is
labelled Some orders. Ignored changed-file status reports the retained earlier format.

## PDF-only update (version 1.9)

The Windows Generate action directly requests PDF. ExportOptions was removed.
The packaged engine's generate request rejects any output_format other than pdf,
including docx. The internal Python helper still supports Word to construct historical
fixtures and verify backward-compatible output history; it is not an employee export
option. Earlier Word-specific smoke scripts target version 1.8 or older. Use
qa/release19_smoke.py for current packaged checks.

RepeatWarning centers a 330 x 54 No button and keeps it as AcceptButton/CancelButton.
A smaller gray LinkLabel below it says Continue to Generate the Repeated Orders,
with no border or underline and no greater width than No. Clicking that link explicitly
approves; the warning token protocol and repeat detection are unchanged.

Fully processed rows say Process again, regardless of surviving file formats. Review,
missing-output and partially covered statuses remain meaningful. Order and summary
product cells now use uniform F2F2F2 shading. Platform header and summary-total colors
are retained. The new layout affects future exports; existing files are unchanged.

## Focused workflow update (version 1.10)

The repeat decision now sends repeat_action=skip for No, include for the continuation
link, and cancels on closing X. The existing selection-bound repeat token and fresh
source validation still apply to both decisions. Filtering uses surviving, current
output fingerprints; deleting every output copy makes the order eligible again.
An unprocessed order shared between PDFs remains printed once. Explicitly reviewed
changed orders are eligible when their fingerprint differs. Revision removals compare
the full original batch, so filtering an unchanged order does not falsely remove it.
If filtering leaves no orders, no document or new completion record is written.

Rows alternate FFFFFF/F2F2F2, starting with white within each order and each summary.
Platform headers/totals retain their colors. Lock controls share the same top margin
and size, with reference-style GDI outline padlocks and no circles/keyholes. Lock and
Excel state handling are unchanged. The warning omits the requested extra sentence.

verify_word.ps1 caches range/page values and disables painting in its owned Word
instance. A summary proven wholly on one page needs no separate row-boundary COM
calls; multi-page summaries still check every row. Source hashing, document content,
quantities, order tables, A4 and page-break validation remain in place.

Current checks: 105 regressions; qa/release110_smoke.py for actual packaged No/mixed,
No/all-processed and intentional reprint behavior; NameLocksSmoke.cs for aligned icons
and lock/unlock behavior; ExportDialogsSmoke.cs for default No, continuation and X.
qa/release110_performance.py compares old/new Word validation on identical documents;
its baseline files and copied private fixtures are local QA inputs, excluded from Git.

## Hidden internal storage (version 1.11)

INTERNAL is .PackingData. HISTORY, LOG, JOURNAL, COURIERS and LOCK are paths relative
to 000 Output containing that directory. internal_directory creates it and sets the
Windows Hidden attribute without changing Explorer preferences. These are persistent
records, never eligible for temporary cleanup. Completed files retain their root paths.

folder_lock acquires the hidden lock; if a legacy root lock exists it also checks that
lock before migration. Root records move via resumable individual renames. All target
collisions are checked before any move; identical duplicates may be removed but differing
copies stop with a review error. Only the obsolete unlocked root lock is removed.
Lock initialization happens after acquisition, avoiding reads of another process's locked
byte. New transaction staging folders live beneath .PackingData; journals store their
paths relative to 000 Output. Legacy journal paths still recover without rewriting them.
The UI processing-log button uses the hidden path.

qa/test_internal_storage.py covers hidden attributes, byte-preserving migration,
couriers, partly migrated dates, duplicate/conflicting records, both locks and new
hidden journal/stage recovery. test_workflow.py covers pre-app logs and old root-stage
recovery. qa/release111_smoke.py takes the candidate and a copied version-1.10 dated
fixture; it exercises packaged migration, No, missing-output regeneration and main UI.
Historical smoke tools may contain old raw metadata paths; use current smoke tests.

## Version 1.11.1

The main PDF grid and selected totals omit pieces. Documents show Total orders
before the first batch/platform heading, without per-batch counts. On scan or
generation, deleted outputs move to archived_outputs; their active order/source
coverage is cleared unless another current copy survives. Audit snapshots are
retained but excluded from repeat detection. Zero-output resets are saved too.
Current packaged checks: qa/release1111_smoke.py; release notes: Release Notes.txt.
