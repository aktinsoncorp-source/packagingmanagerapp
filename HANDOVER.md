# Packing List Manager 1.6 handover

## Current location and release

Workspace: C:\Users\Admin\Desktop\Packing List Manager Program
Installed application: Packing List Manager\Packing List Manager.exe
Live source folders are separate: C:\Users\Admin\Desktop\ONLINE SHOPS
The executable and engine are version 1.6.0. The preceding installed 1.5 release,
including its latest private catalog, is preserved in qa/previous-release-1.5.0.
The earlier 1.4 backup also remains. Installed files match the staged build.

## Version 1.5 foundation

Define Name appears near the bottom of the main screen. It remembers original
product/variation pairs across dates and folders. Version 1.6 matches without
case differences or trailing spaces. The original fields
and alternative fields are editable. Blank alternatives fall back independently.
Alternatives must start with a letter or number; common internal symbols are
allowed. Deleting a row and saving removes its alternatives. A later PDF scan
adds that pair as new with blank alternatives. Matching final name/variation
pairs combine in the per-document picking summary.

The small Excel Export/Import buttons alternate enabled states. Export saves
grid changes, opens an internal temporary workbook in desktop Excel, and enables
Import. Import saves unsaved Excel edits, validates all rows, persists the catalog,
closes only that temporary workbook, deletes it, and re-enables Export. The workbook
must stay open. Closing it or Excel first produces exactly:
Error: No file has been detected
Please do not exit excel file before clicking import
The controls then reset. Invalid data leaves the workbook open for correction.
Other Excel workbooks remain open. Original fields can be edited through hidden
stable Excel row IDs. Missing workbook rows remain in the catalog; delete through
the program. Plain text only; formulas and duplicate/conflicting rows are rejected.

Saved catalog: ProductNames.json beside the installed executable, plus a preceding
ProductNames.json.bak after changes. This private data is excluded from Git and
should travel with the installed application. Existing Word files are unchanged.
Original fingerprints and original snapshots stay in processing history; generated
outputs also retain printed snapshots and the catalog revision.

## Initial discovery

193 unique original product/variation pairs were loaded from 69 compatible PDFs
in the existing dated ONLINE SHOPS folders. Alternatives are initially blank.
124 other PDFs were skipped (mostly waybills; some older unsupported/scanned
packing files). Do not claim those unreadable files were imported. Detailed local
report: PackingListApp/qa/names-catalog-seed.json.

## Verification completed September 27, 2026

- All 50 regression tests passed using isolated fixtures: qa/names-full-tests.log.
- All 18 live Excel/form workflow checks passed against the packaged engine:
  qa/names-excel-verification.json. These include both button states, unsaved edits,
  all four columns and added rows, closed-workbook and full-Excel-exit errors,
  cleanup, preservation of other workbooks, and invalid-input correction.
- Installed engine generated a 7-order / 8-piece test document with alternatives
  in every order and a combined 8-piece summary. Word validated its 2 A4 pages:
  qa/names-packaged-verification.json. This version's installed generation check
  completed successfully; the old 1.4 smoke check had been interrupted.
- Both generated pages were visually checked through Word-exported PDF and
  Poppler. The document skill renderer could not run because LibreOffice is absent.
- Installed main and Define Name screens were visually checked:
  qa/names-main.png, qa/names-editor.png, qa/names-ui-verification.json.
- No production packing documents were regenerated during verification.

## Version 1.6 merged definitions

All four catalog fields now normalize to uppercase and remove trailing spaces.
Control characters, line breaks and Unicode paragraph separators are rejected.
Original PDF objects, quantities, order fingerprints and historical printed snapshots
are preserved. The four editable UI fields retain soft wrapping for readability.

Save names and Excel Import use names_prepare to identify equal, nonblank alternative
product AND variation pairs without writing. The form asks Yes/No per group. Yes saves
one base row plus all other original pairs in its aliases list. No keeps separate
entries. Matching is case insensitive. Selection order chooses the first base; if
none is selected, first visible-list order is used. Excel uses worksheet order.
Every pair in a merged definition is recognized on later scans. View (N) exposes all
linked originals and search finds them. Deleting a merged definition deletes all its
links; later encounters return independent rows with blank alternatives.

Excel exports one row per original pair with two hidden IDs: definition ID and the
original pair before editing. This allows every linked original to be changed while
retaining one merged definition. Omitted rows are retained. An alternative changed on
one linked row updates the shared definition; conflicting changes abort the import.
The original temporary Excel lifecycle and closed-file errors still work.

Schema 1 upgrades to schema 2 under the existing file lock with an atomic write and
backup. Installation preserved all 189 current definitions, IDs and user deletions,
normalizing revision 7 to 8. Do not re-seed the catalog from all historical PDFs;
actual future scans alone may legitimately rediscover deleted pairs.

Version 1.6 verification (September 27):
- 66 regression tests: qa/names-1.6-tests.log (run with PYTHONUTF8=1).
- 35 actual Windows form + desktop Excel checks on the installed engine:
  qa/names-1.6-excel-verification.json. Yes/No, first-selected base, uppercase entry,
  multiline rejection, linked originals, Excel merges, cleanup and search all passed.
- A staged-engine generation created 7 orders / 8 pieces on 2 A4 pages, with one
  merged picking-summary row, preserved original fingerprints and no split orders.
  qa/names-packaged-verification.json; both rendered pages were visually reviewed.
- Existing production documents were not regenerated. The installed UI and engine
  version is 1.6.0; qa/names-1.6-migration-verification.json records the catalog upgrade.
- Source and full employee help are updated. Local Git checkpoint is in Git/v1.6.0.

## Maintenance

Read PackingListApp/README.md and Employee Guide.txt. Current dependencies are in
requirements.txt and requirements-dev.txt. Build with PackingListApp/build.ps1;
-UiOnly compiles the interface without rebuilding the Python engine. Desktop Word
is needed for packing pagination; desktop Excel is needed for temporary workbook
editing only. Direct name-grid editing does not need Excel.

PACKING_TEST_DATA selects the root of the private dated regression fixtures.
qa/prepare_local_fixtures.py can reconstruct that set from retained QA copies and
a supplied shop-data root. The live shop folders now have different filenames and
include waybills, so do not point the historical suite directly at them. The latest
successful fixture root is PackingListApp/qa/names-regression-fixtures.

Git stays local. Keep version bundles in Git/v1.4.0, Git/v1.5.0, Git/v1.6.0, etc. Never overwrite
an earlier version checkpoint. Use git log for the current commit ID. Bundles hold
source/history and sanitized assets, not installed executables or business data.
Do not upload to GitHub until requested. Copy backups off this computer separately.
