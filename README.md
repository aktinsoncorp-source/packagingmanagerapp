# Packing List Manager

Offline Windows application for creating A4 packing lists and per-document
product picking summaries from Shopee SPX/FLASH, Lazada and TikTok PDFs.
Current engine version: **1.6.0**.

Version 1.5 adds **Define Name**: an editable catalog of original product and
variation pairs with independent alternative fields. Matching final names and
variations combine in each document's picking summary. Desktop Excel is needed
only for the temporary Export/Edit/Import workflow; it is not needed to edit
the catalog directly in the program. Keep the exported Excel workbook open
until clicking Import. The program then saves, imports, closes and removes it.

`ProductNames.json` and its preceding `.bak` file live beside the installed
program and are private data excluded from Git. Copy them with the application
when moving computers. Existing Word files retain their printed names.

The program is maintained in [PackingListApp](PackingListApp/README.md).
Employee instructions are in [Employee Guide](PackingListApp/Employee%20Guide.txt).
See [HANDOVER.md](HANDOVER.md) for the current verification status.

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

The full suite contains 66 tests. Several tests and development
helpers require the private September 24-26 source PDFs and approved Word files
at their original dated paths. Those fixtures and generated QA reports are not
distributed in Git. Restore them separately before running the full suite.
Installed-generation smoke tests additionally require the built app and Word.

## Transfer and backups

A Git bundle is a portable copy of committed repository history. To restore a
provided bundle into a new folder:

```powershell
git clone "Git/v1.6.0/PackingListManager-1.6.0.bundle" "PackingListManager-Restored"
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
