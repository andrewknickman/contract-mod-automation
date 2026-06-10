# Changelog

## v00.04.028 — GitHub-ready repository package

### Added

- Prepared a clean repository layout for replacing the existing GitHub repository.
- Added `.gitignore` to exclude generated workbooks, local run folders, caches, virtual environments, and local config.
- Added `.gitattributes` for safer text normalization and binary workbook handling.
- Moved historical audit notes, diffs, and version-specific README/CHANGELOG files under `docs/`.

### Fixed

- Removed accidental leading `\` characters from dependency launcher files and `install_dependencies.py` so they run cleanly:
  - `install_dependencies.py`
  - `INSTALL_DEPENDENCIES.bat`
  - `INSTALL_DEPENDENCIES.command`
  - `VERIFY_DEPENDENCIES.bat`

### Preserved

- No workflow business logic changed from v00.04.027.
- Step 5 J.1 path-scope fix remains intact.
- Dependency installer from v00.04.026 remains included, with launcher cleanup.
- Build, Overview, F&R, Catalog, J.17, MFR, and controller UI logic are otherwise unchanged.

## v00.04.027 — J.1 current path scope fix

- Fixed Step 5 J.1 generation failure caused by `j1_current_file` being referenced before assignment.
- Preserved `.xlsb` / `.xls` conversion behavior from v00.04.025.

## v00.04.026 — Dependency installer

- Added `requirements.txt`.
- Added dependency installer and verifier launchers.
- Included Windows `pywin32` support for Excel-based `.xlsb` / `.xls` conversion.

## v00.04.025 — J.1 XLSB conversion fix

- Added real `.xlsb` / `.xls` to `.xlsx` conversion for previous J.1 workbooks before openpyxl writes to the file.

## v00.04.024 — Build Catalog PR concatenation fix

- Superseded the incorrect v00.04.023 Catalog identity change.
- Kept Price Request Number out of the Catalog duplicate lookup key.
- Concatenated PR numbers on collapsed duplicate Catalog items, for example `59984/59993`.

## v00.04.022 — Build Catalog conversion/deduplication fix

- Corrected Catalog deduplication order.
- Prevented blank/unmapped pricing methods from collapsing by EIS CLIN alone.

## v00.04.021 — Build pricing-factor method correction

- Corrected Build pricing-factor handling for `ORIG JUR`, `TERM JUR`, `ORIG JUR-TERM JUR`, and `ORIG NSC-TERM NSC`.
- Preserved use of PR/J.1 source fields `Orig CJID` and `Term CJID` for JUR matching.

## v00.04.020 — Build OpDiv approval fallback fix

- Blank OpDiv approval status is no longer treated as rejected.
- If no explicit Yes/No decisions exist, the Build includes all valid CLIN rows as pending-review candidates.

## Earlier v00.04.x changes

- Added controller UI.
- Removed hard-coded paths and filenames.
- Added file/folder picker support.
- Added Overview F&R handling corrections for blank, unknown, Pending, and Approved statuses.
- Added F&R script header-variant support for Verizon/HHS comments, Source/Networx/Network fields, case descriptions, and J.1 price matching.
