# Current development checkpoint

## State

- Packing List Manager engine version 1.4.0.
- Full build completed September 26, 2026, at approximately 11:13 PM.
- A later UI-only build corrected the offscreen screenshot helper.
- Installed scan and UI checks passed for all nine September 26 PDFs.
- The Completed documents screen was visually checked.
- All 35 regression tests passed in the preceding development session.
- The final installed Word-generation smoke test was interrupted by a computer
  crash and has **not** completed. Do not report that test as passed.
- No production packing documents were regenerated during those latest checks.
- Repeated computer crashes have an unknown cause.

## Implemented behavior

- Processed PDFs can be selected for regeneration; the Use header selects all
  compatible files. New or missing-output files are automatically selected.
- Removed the courier dropdown and lower order-details list from the main UI.
- Completed documents contains existing Word files, grouped SPX, FLASH, then
  Lazada/TikTok, and ordered by numeric suffix within the platform.
- Expand a document to view its source PDFs, orders and original product details.
- Source links and immutable output snapshots prevent later source changes or
  reused output filenames from misrepresenting a generated document's contents.
- Output suffix is highest currently existing suffix plus one, or 1 if none
  remain. Deleted trailing numbers are reused; middle gaps are not filled.
- Every document ends with its own picking summary after a page break, using
  original product names and variations. Detailed layout rules are in AGENTS.md.

## Important files

- PackingListApp/src/PackingListManager.cs: Windows Forms UI.
- PackingListApp/src/engine.py: processing, history and numbering.
- PackingListApp/src/extractors.py: PDF parsing.
- PackingListApp/src/documents.py: Word generation.
- PackingListApp/src/picking.py: per-document summary.
- PackingListApp/src/assets/verify_word.ps1: Word pagination checks.
- PackingListApp/build.ps1: build and packaging.
- PackingListApp/Employee Guide.txt: employee instructions.

## Evidence available only in the existing local workspace

PackingListApp/qa/completed-documents-tests.log records the 35 passed tests.
The completed-documents-rebuild.log, completed-documents-ui-verification.json,
completed-documents-tree-verification.json, app-completed-main.png and
app-completed-documents.png contain build/UI evidence. These local QA outputs
are excluded from the repository because they can contain real order details.

## Next step

If continuing program development, inspect the current files and saved logs
first. Complete the installed Word-generation smoke check with disposable data
when the computer is stable. Do not repeat builds just to establish whether
the already-installed version exists. Read README.md for restoring the repository
and test-fixture limitations. Use git log for the latest checkpoint identifier.
