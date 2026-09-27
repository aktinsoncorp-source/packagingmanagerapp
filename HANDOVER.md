# Packing List Manager 1.7.0 handover

## Current location and release

Workspace: C:\Users\Admin\Desktop\Packing List Manager Program
Installed application: Packing List Manager\Packing List Manager.exe
Live source folders are separate: C:\Users\Admin\Desktop\ONLINE SHOPS
The executable and engine are version 1.7.0. The preceding installed 1.6.1 release,
including its private catalog, is preserved in PackingListApp/qa/previous-release-1.6.1.
The 1.6.0 backup remains in PackingListApp/qa/previous-release-1.6.0.
The version 1.5 backup remains in qa/previous-release-1.5.0.
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

## Version 1.6.1 employee field

All generated order-table headers now read Order ID/Order Number, followed by
Employee:__________________, Number and Time. The employee line has 18 underscores.
The shared builder applies it to SPX, FLASH, Lazada and TikTok, in both orange and
blue templates. It retains Arial 9.5 pt and uses three tab stops (185/340/440 pt).
The introduction now tells staff to write the employee name as well as number/time.
Document re-open validation checks the full header text. No name-catalog schema or
business-data changes are involved. Existing completed documents keep their layout;
regenerate selected PDFs to obtain the employee field in a new numbered document.

Verification report: qa/employee-1.6.1-verification.json. The smoke script uses copied
PDFs and an isolated catalog, actual packaged generation, Word pagination, original
quantity checks, a second scan, and PDF checks for field order and single-line fit.
All six final preview pages were rendered with Poppler and visually reviewed.
17 order headers / 30 pieces passed across all four platforms. Installed files match
the tested staged build, and ProductNames.json was unchanged by installation. Source release
and employee help are updated; local version bundle is under Git/v1.6.1.

## Version 1.7 continued orders and courier selection

Shopee fragments with one footer Order ID are supported after a first label with
its heading and item 1. Every later fragment must continue the printed numbering
for that Order ID. They become one order; gaps, overlaps, duplicated full labels,
or orphan continuation fragments remain errors. Multiple-package labels remain
unsupported. Original product names/variations and quantities are retained.

The main PDF grid now has per-file Auto, SPX and FLASH choices for Shopee. Auto
prioritizes the nearest exact SPX/FLASH folder (case insensitive), then one filename
label, then unambiguous remaining path hints. Explicit choices override Auto and
persist by relative path and SHA256 in 000 Output/Courier_Selections.json. Changed
PDF contents reset the choice to Auto; stale UI choices are rejected at generation.
Lazada/TikTok have disabled dash cells. Compatible PDFs anywhere under the selected
dated folder are included, excluding 000 Output. Unrecognized, corrupt, scanned and
textless PDFs are omitted. Recognized packing data errors remain visible. Generic
Shopee PDFs lacking a courier remain visible with counts and require a choice.
Add PDFs accepts those generic files into the dated folder root.

Version 1.7 verification (September 27):
- All 80 regression tests passed: qa/couriers-1.7-tests.log.
- The final packaged Windows form verified Auto -> FLASH -> SPX -> Auto changes,
  persistence across scans, disabled non-Shopee cells, and select-all eligibility.
  Five supported PDFs appeared; blank/corrupt files were hidden. The main screenshot
  confirms FLASH/SPX_old.pdf resolves to FLASH in Auto because its folder wins.
- The reported 32-page SPX PDF generated 127 orders / 300 pieces. Continued order
  2609263ADBUPHF contains 11 product lines / 14 pieces in one table. The summary
  contains 300 pieces. Word verified all 127 orders stay together on A4, with the
  picking summary after a page break; the output is 16 pages including summary.
- The continued-order page and summary total page were rendered via Word/Poppler
  and visually checked. Report: qa/couriers-1.7-verification.json. QA used isolated
  copies and a separate name catalog; production documents were not regenerated.
- Installation checked all 297 packaged engine files plus the executable, icon and
  guide against the tested candidate. ProductNames.json, its backup and Settings.json
  were unchanged. Report: qa/couriers-1.7-installation.json. Employee help is updated.
- Local source checkpoint: Git/v1.7.0. Previous version folders are preserved.

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
