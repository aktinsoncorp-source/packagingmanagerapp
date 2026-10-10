# Packing List Manager

Offline Windows application for creating A4 packing lists and per-document
product picking summaries from Shopee SPX/FLASH, Lazada and TikTok PDFs.
Current release: **2.0.0**, in [V2.0.0](V2.0.0/README.md), approved for `main`.
V2 uses SKU-based definitions and the Station Board Python desktop interface.
Its local `V2.0.0/Release` folder contains a Windows installer and portable ZIP
with a private Python runtime. Earlier V1 sources and the installed **1.11.2**
application remain preserved in their original folders.

Version 1.5 adds **Define Name**: an editable catalog of original product and
variation pairs with independent alternative fields. Matching final names and
variations combine in each document's picking summary. Desktop Excel is needed
only for the temporary Export/Edit/Import workflow; it is not needed to edit
the catalog directly in the program. Keep the exported Excel workbook open
until clicking Import. The program then saves, imports, closes and removes it.

`ProductNames.json` and its preceding `.bak` file live beside the installed
program and are private data excluded from Git. Copy them with the application
when moving computers. Existing Word files retain their printed names.

V2 is maintained in [V2.0.0/PackingListApp](V2.0.0/PackingListApp/README.md).
Employee instructions are in [V2 Employee Guide](V2.0.0/PackingListApp/Employee%20Guide.txt).
The notes below and [HANDOVER.md](HANDOVER.md) retain the V1 maintenance history.

## Repository contents

This repository contains program source, sanitized Word templates, application
icon assets, the build script, employee help, QA scripts and project instructions.
The current reusable program and its ongoing maintenance are authorized.

The root .gitignore deliberately allows only program-related paths. Dated shop
folders, customer PDFs, completed packing lists, processing history, installed
executables, QA results, dependency caches and local settings are not tracked.
The unused historical product-alias catalog is also excluded.

Adding an unrelated file or folder at the repository root will not automatically
include it in Git. Review .gitignore deliberately when adding new source paths.

## Build on another Windows computer

Requirements: Windows x64, Python 3.12, .NET Framework's C# compiler, and desktop
Microsoft Word for generated-document pagination validation.

From this repository's root in PowerShell:

```powershell
py -3.12 -m venv .venv
& .\.venv\Scripts\python.exe -m pip install -r .\PackingListApp\requirements-dev.txt
& .\PackingListApp\build.ps1 -Python (Resolve-Path .\.venv\Scripts\python.exe).Path
```

The existing build uses Codex's bundled Python by default and a local
PackingListApp/.builddeps directory. The explicit -Python argument selects the
environment above. Its installation commands are documented for a new computer;
they were not rerun while creating this checkpoint. The direct dependency
versions are recorded; transitive dependencies are not fully locked.

The build creates `Packing List Manager/Packing List Manager.exe` and its
`App Files` folder. Keep the complete release folder together. Employees do not
need Python, ChatGPT or internet access to run that built application.

## Checks

The 16 synthetic regeneration tests require no production PDFs or Microsoft Word:

```powershell
& .\.venv\Scripts\python.exe -m unittest discover -s .\PackingListApp\qa -p test_regeneration.py -v
```

The full suite contains 158 tests. Several tests and development
helpers require the private September 24-26 source PDFs and approved Word files
at their original dated paths. Those fixtures and generated QA reports are not
distributed in Git. Restore them separately before running the full suite.
Installed-generation smoke tests additionally require the built app and Word.

## Transfer and backups

A Git bundle is a portable copy of committed repository history. To restore a
provided bundle into a new folder:

```powershell
git clone "Git/v1.11.1/PackingListManager-1.11.1.bundle" "PackingListManager-Restored"
```

This restores the tracked program files. It does not restore the installed
release, business PDFs, completed documents, processing history, or other ignored
files. Copy those separately when moving the working application.

After reviewing and committing a later release, create a new `Git/vVERSION`
folder, create its named bundle with `git bundle create <new-bundle-path> main`,
then run `git bundle verify <new-bundle-path>`. Preserve earlier version folders
and never overwrite a prior checkpoint. These backups are snapshots and do not
automatically update after new commits.

Keep a copy on another drive or an authorized private remote to protect against
loss of this computer. A bundle stored only beside the project is still local.


## Version 1.6 - merged standardized names

Define Name can merge matching alternative product/variation pairs after confirmation.
The first selected definition stays as the base and retains every linked original pair.
View reveals the linked pairs. All four fields are uppercase, trailing spaces are trimmed,
and multiline input is rejected. Excel edits preserve links and can edit every original.
Local source checkpoint: `Git/v1.6.0`; older version folders remain available.


## Version 1.6.1 - employee identification

Every generated order header now contains Order ID/Order Number, followed by
Employee:__________________, Number and Time. This applies to SPX, FLASH, Lazada
and TikTok, retaining the orange/blue themes and A4 layout. Regenerate an older
batch to include the field in a new numbered document. Source data is unchanged.
Local source checkpoint: `Git/v1.6.1`.

## Version 1.7 - continued orders and per-PDF couriers

Shopee order fragments with matching IDs and consecutive item numbers are joined.
Each Shopee PDF has Auto/SPX/FLASH choices. Auto uses the nearest exact SPX or FLASH
folder first, before filename hints. Manual choices persist with the PDF fingerprint.
Compatible PDFs are recognized throughout the dated folder. Unknown/unreadable PDFs
are hidden; genuine errors in recognized packing lists remain visible for review.
Lazada/TikTok courier cells show a disabled dash. Local checkpoint: `Git/v1.7.0`.

## Version 1.8 - export and employee controls

Generate offers PDF by default, with Word as an alternative. Both use Word's A4
validation and share output numbering/history. Repeated orders require confirmation;
No is the larger default button. Define Name opens locked with icon-only padlock
controls. Checked-PDF totals update below the main list, without double counting
matching orders. Main step numbers are removed and printed Done boxes use five spaces.
Local checkpoint: `Git/v1.8.0`.

Version 1.8 identified output formats in Status; version 1.9 replaces that label with Process again.

## Version 1.9 - PDF-only generation and simpler reprints

Generate now exports PDF directly, with no format selector. Completed rows show
Process again. The repeat warning has a large centered No button and small gray
borderless continuation text beneath it. Product rows use light gray shading;
SPX/FLASH retain peach headers and Lazada/TikTok retain blue headers. Existing Word
outputs remain available in history and count toward numbering. Checkpoint: `Git/v1.9.0`.

## Version 1.10 - skip earlier generated orders and alternating rows

No in the repeat warning generates orders without an existing matching output;
clicking the gray continuation link includes earlier generated orders. All-covered
selections display a no-new-orders message and create no empty document. Deleted
outputs remain eligible. Product rows alternate white and light gray, with the
existing platform headers. The two outline padlocks have matching alignment.
Word automation makes fewer redundant calls while preserving all correctness and
pagination requirements. Three-document Word check/export comparison: 11.76 to
9.47 seconds in one local test; larger batches may see a smaller gain. PDF text and
pagination matched. All 105 regression tests passed. Checkpoint: `Git/v1.10.0`.

## Version 1.11 - hidden processing records

Internal history, processing log, lock, courier choices and transaction/staging data
now live in the Windows-hidden `000 Output/.PackingData` folder. This is permanent
program data and must be included in dated-folder backups. Generated PDFs remain
in `000 Output`. Older records migrate automatically under the date lock; existing
copies are checked before moving. Open processing log uses the hidden location.
All 113 regression tests and packaged migration/regeneration checks passed. Four live
dates were migrated with record hashes preserved and all 13 completed documents
unchanged. Local source/history checkpoint: `Git/v1.11.0`.

## Version 1.11.1

See [Release Notes](PackingListApp/Release%20Notes.txt) for the new functions.
114 regressions passed, followed by 17 final regeneration checks. The final
packaged app passed PDF/Word layout verification on all four platforms and
Windows UI checks. Local checkpoint: `Git/v1.11.1`.

## Version 1.11.2

SPX/FLASH packing lists include buyer comments only when present. Define Name
matches spacing and wrapping variants against existing definitions; Lazada
variations show values without attribute labels. Local checkpoint: `Git/v1.11.2`.
