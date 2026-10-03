# Packing List Manager 1.11.1 handover

## Current location and release

Workspace: C:\Users\Admin\Desktop\Packing List Manager Program
Installed application: Packing List Manager\Packing List Manager.exe
Live source folders are separate: C:\Users\Admin\Desktop\ONLINE SHOPS
The executable and engine are version 1.11.1. The preceding installed 1.11.0 release, including its private catalog, is
preserved in PackingListApp/qa/previous-release-1.11.0. The 1.10.0 backup remains.
The 1.9.0 backup remains in PackingListApp/qa/previous-release-1.9.0.
The 1.8.0 backup remains in PackingListApp/qa/previous-release-1.8.0.
The 1.7.0 backup remains in PackingListApp/qa/previous-release-1.7.0.
The 1.6.1 backup remains in PackingListApp/qa/previous-release-1.6.1.
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

## Version 1.8 export and employee workflow

- Main labels no longer have step numbers. Clicking Generate opens a format choice
  that defaults to PDF each time; Word (.docx) is available. Word validates A4 layout
  for both formats and exports PDF. Temporary DOCX files are removed for PDF exports;
  only the chosen format is published and recorded. Existing DOCX history is retained.
- PDF and Word share platform/date numeric sequences. Completed Documents lists both
  formats, opens either, and retains source/order/item snapshots. Deleting outputs
  removes their coverage on scan; trailing suffix reuse and lower-gap rules still apply.
- Repeat-order checking covers earlier history and overlaps with other readable PDFs
  in the same dated folder (including unchecked PDFs). The engine returns a selection/
  format-specific confirmation token before building files. The exact requested warning
  uses a larger blue No button as Enter/Escape default; Yes is smaller. Declining creates
  no output. Identical selected orders continue to print once. Checks are per date.
- Define Name opens locked on every new form. Closed/open padlock icons have tooltips
  and accessible names. Search and linked-original viewing work while locked; edits,
  add/delete, Save and Excel Export/Import are disabled. Mutation methods reject locked
  calls too. Save edits before relocking. Pending Excel sessions cannot be relocked.
  The lock has no password and is intended to prevent accidental editing.
- Done boxes in packing and picking-summary tables contain five spaces: [     ].
- Totals below the main grid track only checked PDFs, unique orders and pieces, with
  selected platform breakdowns. Deduplication follows generation's revision-first order.
- Status shows Processed - PDF, Processed - Word or Processed - PDF + Word. Formats
  are inferred from surviving nonsuperseded outputs matching current order fingerprints;
  old Word records need no new metadata. Partial coverage is marked Some orders.
  Rescanning after deletion drops the missing format or shows Output missing.

Verification September 28:
- All 96 regression tests passed: qa/release-1.8-tests.log.
- All 44 actual name-form/desktop Excel checks passed; latest result directory is
  recorded in qa/release18-controls/excel-latest.txt. Existing merge, uppercase,
  single-line, closed-Excel errors and temporary-workbook cleanup remain verified.
- Real export dialogs confirmed PDF default on every open, selectable Word, a larger
  No button, Enter/Escape defaulting to No, and cancellation by the default action.
  Evidence: qa/release18-controls/dialogs-ok.txt and its PNGs.
- Packaged generation used isolated copies across all four platforms: 17 orders /
  30 pieces. PDF outputs used _1 and later Word outputs _2. Word checked all order
  tables and summary page breaks. All six PDF pages were rendered with Poppler and
  visually reviewed. All 38 tick boxes had the increased measured spacing; Word box
  text was verified too. Report: qa/release-1.8-verification.json.
- Main UI checks covered per-file courier changes, check-all and dynamic individual
  selection totals. Final rebuilt engine scan showed PDF + Word for all four processed
  platform rows. Final UI evidence is under qa/release18-0vi2ax_9/main-final.png and
  status-response.json. Locked Define Name and mixed-format Completed Documents were
  visually checked. Production packing documents were not regenerated during QA.
- Installed files match the final staged release: 297 engine files plus executable,
  icon and guide. ProductNames.json, its backup and Settings.json were preserved.
  Report: qa/release-1.8-installation.json. Local checkpoint: Git/v1.8.0.

## Version 1.9 PDF-only generation and neutral product rows

- Generate directly exports PDF with no format-choice dialog. The public engine also
  rejects Word requests. Desktop Word still performs internal A4 pagination and PDF
  export; temporary Word files are not published. Historical Word files remain valid
  for processing history, missing-output coverage and shared suffix numbering.
- Fully processed source rows show Process again, regardless of historical format.
  Partial coverage and review/missing-output warnings retain useful status messages.
- Repeated-orders warning has a large centered blue No button. No is focused and is
  the Enter/Escape default. Below it, smaller light-gray borderless text reads exactly
  Continue to Generate the Repeated Orders. Its width matches the No button.
- Every product row in order tables and picking summaries has F2F2F2 light-gray fill.
  Peach SPX/FLASH and blue Lazada/TikTok headers and summary totals are preserved.
- Employee Guide.txt explains the current workflow. Versions 1.8 and older above are
  historical; their format-choice and format-specific Status behavior is superseded.

Verification September 28:
- All 97 regression tests passed: qa/release-1.9-tests.log.
- Real warning-dialog checks passed: centered and enlarged No, default cancellation,
  smaller borderless continuation link within its width, and explicit continuation.
  Evidence: qa/release19-controls/dialogs-ok.txt and repeat-warning.png.
- Packaged engine generated three PDF-only outputs from all four platforms (17 orders,
  30 pieces), then confirmed a repeated SPX generation as _2.pdf. Word export requests
  were rejected. All four processed sources displayed Process again on rescan.
- All six initial A4 PDF pages were rendered and visually reviewed. Confirmed the
  gray product rows and colored headers, 38 wider boxes, and separate summary pages.
  The PDF inspection checks DeviceGray as well as RGB fills; 650 gray rectangles found.
  Report: qa/release-1.9-verification.json. The script resumed the same isolated run
  after correcting its color-inspection assumption and resumed-output path handling.
- Packaged main UI checks passed for per-file couriers, manual/check-all selection,
  and dynamic platform totals. Main screenshot: qa/release19-dqtgf_9i/main.png.
- Installation verified all 297 engine files plus executable, icon and employee guide
  against the tested candidate. ProductNames.json, its backup and Settings.json were
  unchanged. Report: qa/release-1.9-installation.json. Production documents were not
  regenerated. Source/history checkpoint: Git/v1.9.0; earlier versions are preserved.

## Version 1.10 focused repeat, row and lock changes

- No in the warning now generates orders without a surviving matching output and
  skips already-covered orders. The gray continuation link includes earlier orders.
  All-covered selections report no new orders and create no empty document/history.
  Deleted outputs and explicitly reviewed changed order versions remain eligible.
  The extra Matching orders sentence was removed only from the warning.
- Product rows alternate white/light gray; peach/blue headers and totals are retained.
- Lock icons are aligned and use simple outline bodies/open or closed shackles from
  the user's reference style, with no enclosing circles or keyholes. Lock logic stays.
- Performance work is restricted to Word automation: cached range/page properties,
  painting disabled in the owned Word instance, and an equivalent whole-summary check
  when it occupies one page. Multi-page summaries still check each row. All other
  source, order-content, quantity, hashing and pagination verification remains.
- Version 1.9 descriptions above are historical; No no longer cancels generation.
  Closing X still cancels; Enter/Escape choose No. Employee help is updated.

Verification September 28:
- All 105 regression tests passed in 68.7 seconds: qa/release-1.10-tests.log. New cases
  cover mixed old/new, all-covered, deleted outputs, new shared orders, repeat tokens,
  revision removals, public No forwarding and alternating shading.
- Actual dialog and aligned-padlock tests passed under qa/release110-controls. Default
  locked state, unlocking and relocking remain correct. X was tested in modal mode.
- Packaged engine generated three PDFs covering 17 orders/30 pieces, skipped the
  previously generated SPX selection while producing the other two groups, avoided an
  empty document for all-covered No, and produced an intentional SPX _2.pdf reprint.
  Main UI courier controls, check-all, manual selection and dynamic totals passed.
  Report: qa/release-1.10-verification.json; run: qa/release110-4rkvocaf.
- Six initial sample PDF pages and the 15-page 127-order/300-piece continuation sample
  were rendered and visually checked. All 38 widened boxes remained in the small set.
- Old/new Word check/export on the same three documents took 11.757/9.475 seconds in
  one local comparison (~19% faster for that phase). The large sample took 45.955 vs
  44.386 seconds for Word (a smaller gain). These are sample timings, not a promised
  overall speedup. Every page's PDF text and pagination matched in both comparisons.
  Evidence: qa/performance110/comparison.json, baseline.json and updated.json.
- Installation verified 297 engine files plus executable, icon and employee guide.
  ProductNames.json, its backup and Settings.json were preserved. Report:
  qa/release-1.10-installation.json. No production documents were regenerated.
- Local Git source/history checkpoint: Git/v1.10.0. Earlier version folders remain.

## Version 1.11 hidden processing records

- Processing_History.json, Processed_Files.txt and .packing.lock now live in the
  Windows-hidden 000 Output/.PackingData folder, with courier choices, transaction
  journals and new staging folders. The folder is persistent, not disposable temp data.
- Generated PDFs and historical Word documents remain directly in 000 Output.
  Open processing log opens the hidden log. No Explorer settings were changed.
- Scanning an existing date migrates root records under the new and legacy locks.
  Renames can resume after interruption. Identical leftovers are removed; conflicting
  copies stop before movement. Existing root-stage journals remain recoverable.
- All 113 regression tests passed in 72.6 seconds: qa/release-1.11-tests.log. New tests
  cover exact record preservation, hidden attributes, partial migration, conflict and
  duplicate handling, legacy/current lock exclusion and hidden-stage recovery.
- Packaged checks passed on copied 1.10 output: all four internal files migrated with
  exact bytes, PDFs remained unchanged during migration, No skipped existing orders,
  deleted SPX outputs regenerated as _1.pdf, and the final three outputs retained
  17 orders/30 pieces. Main UI checks for couriers, check-all and totals passed.
  Report: qa/release-1.11-verification.json; fixture: qa/release111-0o7p5asv.
- Installation verified 297 engine files plus executable, icon and employee guide;
  private catalog, catalog backup and settings were unchanged. Report:
  qa/release-1.11-installation.json.
- Live ONLINE SHOPS dates 20260924, 20260925, 20260926 and 20260928 were migrated.
  All 12 internal records retained exact hashes; all 13 completed documents retained
  exact hashes. No live packing documents were regenerated. Metadata backups are in
  qa/metadata-backup-1.10.0, one folder per date. Report:
  qa/release-1.11-live-migration.json.
- Copy the entire dated folder including .PackingData for business backups/transfers.
  Older executables do not understand the new location; do not run them against a
  migrated date without a deliberate storage rollback plan preserving current history.
- Local source/history checkpoint: Git/v1.11.0; previous checkpoints are preserved.

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

## Version 1.11.1

The main PDF list and selection totals show orders without pieces. Output documents
show a total order count above batch/platform headings; per-batch counts are removed.
Deleted outputs are removed from active history, with immutable output manifests
archived for audit. Uncovered orders become new, with no historical repeat warning.
Surviving copies still count; source overlap warnings and revision reviews for
surviving outputs remain. Local bundle: Git/v1.11.1.

Verification: 114 regressions passed; 17 regeneration checks reran after the
final history guard. Final packaged smoke: qa/release-1.11.1-verification.json.
17 orders / 30 pieces across all four platforms passed actual Word A4 and
whole-order/summary pagination checks; six PDF pages and main UI were visually
reviewed. 301 installed files match the tested candidate; ProductNames.json,
its backup and Settings.json retain their hashes. No production PDFs regenerated.
