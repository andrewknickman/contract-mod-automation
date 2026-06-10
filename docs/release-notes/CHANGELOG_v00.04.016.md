# Changelog

All notable changes for this project will be documented here.

Versioning note: this project is currently continuing the `v00.04.xxx` iteration line unless/until the base version is intentionally bumped.

## v00.04.016 - F&R Comment Column and J.1 Rate Matching Fix

### Fixed

- Corrected `03_F_and_R_script.py` so the generated F&R workbook can copy Verizon/HHS comment text from accepted COMP-sheet header variants instead of requiring only the exact header `Verizon's Response/HHS Comment`.
- Accepted comment variants now include forms such as `Verizon Response`, `Verizon's Response`, `Verizon's Reponse`, `Verizon Response/HHS Comment`, and `HHS Comment`.
- When separate Verizon-response and HHS-comment columns are present, their nonblank values are combined into the generated F&R workbook's canonical `Verizon's Response/HHS Comment` column.
- Improved J.1 price matching while preserving the business rule that J.1 headers are read from Row 1 only.
- J.1 matching now tolerates Excel cell-type differences such as numeric case numbers versus text case numbers, SRE pricing element `1` versus `01`, and TO Period `OPT PD 5` versus `OPT PD 05`.
- Updated `.xlsx` / `.xlsm` J.1 reads to use `data_only=True` so Excel-cached formula results in `HHS Price` can be read when available.
- Added clearer diagnostic output when no J.1 match is found for a case/pricing/option-period combination.

### Preserved

- J.1 column headers are still expected on Row 1 of the PR file.
- Required J.1 header names remain `Case Number`, `SRE Pricing Element`, `TO Period`, and `HHS Price`.
- The F&R workbook still uses the canonical output header `Verizon's Response/HHS Comment`.
- The Overview F&R logic from `v00.04.015` was not changed.

### Validation

- Confirmed `03_F_and_R_script.py` compiles with `python -m py_compile`.
- Ran targeted tests confirming accepted Verizon/HHS comment header variants are copied into the canonical output field.
- Ran targeted workbook tests confirming J.1 Row 1 headers are used and numeric/text value differences still match the correct `HHS Price`.

## v00.04.015 - Overview Unknown F&R Status Handling

### Fixed

- Corrected `02_overview_file_script.py` so unexpected nonblank `F&R Status` values, such as `?`, are no longer treated as a confirmed `No` in the generated Overview.
- `F&R Needed` now stays blank when no in-scope `Pending` or `Approved` rows are found and at least one in-scope F&R status is blank or unknown.
- The adjacent Overview F&R `# of CLINs` value is now also left blank whenever `F&R Needed` is blank.
- Preserved the business rule that only `Pending` and `Approved` count as F&R items.
- Preserved the explicit `N/A` / `NA` behavior: if all in-scope statuses are recognized non-F&R values, the Overview may still show `F&R Needed = No` and F&R count `0`.

### Resulting Overview behavior

| In-scope F&R Status values | F&R Needed | F&R count column |
| --- | --- | --- |
| Pending | Yes | Counted |
| Approved | Yes | Counted |
| N/A / NA only | No | 0 |
| All blank | Blank | Blank |
| Unknown values only, such as `?` | Blank | Blank |
| Blank/unknown mixed with N/A, but no Pending/Approved | Blank | Blank |

### Validation

- Confirmed the updated Overview script compiles with `python -m py_compile`.
- Ran targeted in-memory workbook tests confirming `?`, blank, `Pending`, `Approved`, and `N/A` scenarios.

## v00.04.014 - Overview Blank F&R Status Handling

### Fixed

- Corrected `02_overview_file_script.py` so the generated Overview no longer defaults `F&R Needed` to `No` when every in-scope `F&R Status` cell is blank.
- `F&R Needed` is now left blank when the scoped CLIN rows have no F&R status values at all.
- Preserved the existing business rule that only `Pending` and `Approved` statuses produce `F&R Needed = Yes` and increment the numeric F&R count.
- Preserved the OpDiv row-scope behavior: if OpDiv-approved rows exist, only those rows are evaluated; otherwise, all listed CLIN rows are evaluated.

### Resulting Overview behavior

| In-scope F&R Status values | F&R Needed | F&R count |
| --- | --- | ---: |
| All blank | Blank | 0 |
| Pending | Yes | 1 per Pending row |
| Approved | Yes | 1 per Approved row |
| N/A / NA only | No | 0 |
| Blank plus N/A / NA only | No | 0 |

### Validation

- Confirmed all Python files compile with `python -m py_compile`.
- Ran targeted in-memory workbook tests confirming blank-only status rows leave `F&R Needed` blank while `Pending` and `Approved` still produce `Yes` and increment the count.

## v00.04.013 - Overview F&R Needed Rule Alignment

### Fixed

- Corrected `02_overview_file_script.py` so `F&R Needed = Yes` is triggered only when an in-scope CLIN row has `F&R Status = Pending` or `F&R Status = Approved`.
- Corrected the prior `v00.04.012` behavior where `N/A` / `NA` could still trigger `F&R Needed = Yes`.
- Kept the numeric F&R count aligned to the same rule: only `Pending` and `Approved` are counted.
- Preserved the CLIN-row scope behavior: if OpDiv-approved rows exist, evaluate only those rows; if none exist, evaluate all listed CLIN rows.

### Clarified

- `N/A`, `NA`, `N.A.`, and `Not Applicable` do not trigger `F&R Needed = Yes` and do not increment the F&R count.
- `Pending` and `Approved` both trigger `F&R Needed = Yes` and both increment the F&R count.

### Validation

- Confirmed the updated overview script compiles with `python -m compileall`.
- Ran targeted in-memory workbook tests confirming:
  - `N/A` / `NA` alone produces `F&R Needed = No` and F&R count `0`.
  - `Pending` and `Approved` produce `F&R Needed = Yes` and increment the F&R count.
  - When OpDiv-approved rows exist, unapproved rows are excluded from both the flag and the count.

## v00.04.012 - Overview F&R Count Business Rule Correction

> Superseded by `v00.04.013` for the `F&R Needed` yes/no trigger. The current rule is that only `Pending` or `Approved` in-scope rows trigger `F&R Needed = Yes`.

### Fixed

- Corrected `02_overview_file_script.py` so the numeric Overview F&R item count includes only `Pending` and `Approved` F&R statuses.
- Preserved the then-current `F&R Needed = Yes` trigger behavior for `Pending` and `N/A` / `NA` statuses. This was later corrected in `v00.04.013`.
- Preserved the CLIN-row scope behavior from `v00.04.010` / `v00.04.011`: when OpDiv-approved rows exist, the Overview uses those rows; when none exist, it falls back to all listed CLIN rows.

### Clarified

- `v00.04.012` still allowed `N/A` / `NA` to make the Overview say `F&R Needed = Yes`; this was later corrected in `v00.04.013`.
- `v00.04.012` counted `Approved` as an F&R item but did not use it by itself to trigger `F&R Needed = Yes`; this was later corrected in `v00.04.013`.

### Validation

- Confirmed the updated overview script compiles with `python -m py_compile`.
- Ran targeted in-memory workbook tests confirming:
  - `N/A` produced `F&R Needed = Yes` and F&R count `0` under the superseded `v00.04.012` rule.
  - `Pending` produced `F&R Needed = Yes` and incremented the F&R count.
  - `Approved` incremented the F&R count but did not by itself trigger `F&R Needed = Yes` under the superseded `v00.04.012` rule.

## v00.04.011 - Overview F&R Count Alignment Fix

### Fixed

- Fixed `02_overview_file_script.py` so `F&R Needed = Yes` and the Overview F&R item count are calculated from the same CLIN row scope.
- Included `N/A` / `NA` F&R statuses in the numeric F&R item count. Previously, `N/A` could trigger `F&R Needed = Yes` while still producing `# of CLINs = 0` in the F&R count column.
- Preserved the approval-row behavior when one or more rows are marked `OpDiv Approval = Yes`.
- Preserved the fallback-row behavior when no rows are marked with OpDiv approval: the script uses all listed CLIN rows instead of returning zero.
- Added normalization for common N/A variants, including `N/A`, `NA`, `N.A.`, and `Not Applicable`.

### Validation

- Confirmed the updated overview script compiles with `python -m py_compile`.
- Ran targeted in-memory workbook tests confirming that `F&R Needed = Yes` from an `N/A` status now produces a nonzero F&R count when that row is in scope.

## v00.04.010 - Overview CLIN Count Fallback and F&R Determination Audit

### Fixed

- Fixed `02_overview_file_script.py` so `# of CLINs (Line items in current OP)` no longer reports `0` solely because no CLIN rows are marked `OpDiv Approval = Yes`.
- Preserved the prior approval-based count when one or more rows are marked with OpDiv approval.
- Added a fallback count that counts actual listed CLIN line items when no rows are marked with OpDiv approval.
- Added the same fallback concept to the Overview `F&R Needed` result: when no OpDiv approval rows are marked, the script now looks at the listed CLIN rows' F&R statuses instead of automatically treating the PR as not needing F&R.
- Limited the fallback count to rows that appear to be actual CLIN line items so blank formatted table rows are not treated as CLINs.

### Clarified

- Added `F_R_DETERMINATION_AUDIT.md` explaining how `03_F_and_R_script.py` decides which PRs get F&R workbook sheets.
- Documented that the F&R workbook is driven by the generated Overview workbook's `F&R Needed` column, then by matching PR files and locating a `comp`/`comps` worksheet.
- Documented common skip causes: Overview row not marked `F&R Needed = Yes`, PR file name not matching the PR number prefix, missing `comp` sheet, empty comp data, or PR grouping by PR number.

### Validation

- Confirmed the updated overview script compiles with `python -m py_compile`.
- Ran targeted in-memory workbook tests confirming the CLIN fallback count works when no OpDiv approval rows are marked.

## v00.04.009 - Controller Regression Audit and Preflight Fix

### Fixed

- Fixed controller preflight handling for fields that are outputs in an earlier step and inputs in a later step.
- Full-workflow preflight now allows planned prior outputs, such as generated coversheets, overview workbook, build workbook, current J.1 workbook, and updated J.17 workbook, to be created during the run instead of incorrectly requiring them to exist before the workflow starts.
- Individual step preflight now correctly requires those same prior-output paths to already exist when running that step by itself.
- Restored the original J.17 expired-subscription default month behavior by defaulting the billing/current month to `December` unless the user overrides it in the UI or CLI.

### Audited

- Compared the original seven-script project against the controller UI package.
- Confirmed all original processing scripts remain present.
- Confirmed no original top-level processing functions were removed from the seven scripts.
- Confirmed the controller delegates to the scripts rather than replacing their business logic.

### Validation

- Confirmed all Python files compile with `python -m py_compile`.
- Confirmed the J.17 script exposes the restored `--current-month` default behavior through `--help`.

## v00.04.008 - Controller UI Update

### Added

- Added `mod_automation_ui.py`, a Tkinter desktop controller for the full seven-step workflow.
- Added file and folder selection fields for all core inputs:
  - J.1 previous workbook
  - CLIN lookup workbook
  - PR files folder
  - source build workbook/template
  - source J.17 workbook
  - billing detail workbook
  - generated/approved coversheets folders
  - generated output workbooks/documents
- Added run metadata fields for:
  - mod number
  - current option period
  - optional billing month
  - Python executable path
- Added default output path generation for a clean run folder structure.
- Added a preflight check before individual-step and full-workflow runs.
- Added live script output logging in the UI.
- Added per-step status tracking: Pending, Running, Complete, Failed, and Skipped.
- Added buttons to run a single step, run the full workflow, stop the active process, show the generated command, open the output folder, clear the log, and save the log.
- Added save/load support for JSON workflow configuration files.
- Added `RUN_UI.bat` as a Windows launcher.
- Added `RUN_UI.command` as a macOS/Linux launcher.

### Changed

- Updated README documentation to make the controller UI the recommended way to run the workflow.
- The workflow can now be facilitated from one screen instead of requiring users to run seven separate scripts manually.
- Output locations are still flexible, but the UI now offers standard default output names and folders for consistency.

### Validation

- Confirmed `mod_automation_ui.py` compiles successfully.
- Confirmed the UI delegates execution to the existing script CLI arguments added in `v00.04.007`.
- Full end-to-end execution still requires the actual Excel/Word inputs used by the mod process.

### Known limits

- The UI validates required file/folder existence and output path readiness, but it does not yet validate workbook schemas, required sheets, or required columns before launching scripts.
- The UI depends on Tkinter, which is included with most standard Python installations but may be missing from some restricted Python builds.
- `xlwings` and Microsoft Excel are still required for workflow steps that rely on the original Excel COM behavior.

## v00.04.007 - Path and Filename Flexibility Update

### Added

- Added `file_selection.py` as a shared helper for runtime file and folder selection.
- Added file picker fallback support for required input files.
- Added folder picker support for PR folders, coversheet folders, and output folders.
- Added save-as picker support for generated output files.
- Added command-line arguments to all seven workflow scripts.
- Added broader Excel extension recognition for folder scanning where supported:
  - `.xlsx`
  - `.xlsm`
  - `.xlsb`
  - `.xls`
- Added `.xlsb` read support for pandas-based reads through `pyxlsb`.
- Added dynamic PR matching through shared helper logic.
- Added `PATH_FLEXIBILITY_REPORT.md` documenting the hard-coded paths and filenames removed.
- Added `path_flexibility.diff` showing the patch-level code changes.

### Changed

- Replaced the shared hard-coded local data root:

  ```text
  D:\1_emerjence_work\04_HHS\09_new_mod_automation\Data
  ```

  with script-level selectable file and folder paths.

- Updated `01_final_coversheet_generation_script.py` to accept selected paths for:
  - J.1 previous workbook
  - CLIN lookup workbook
  - PR folder
  - coversheet output folder
  - current option period

- Updated `02_overview_file_script.py` to accept selected paths for:
  - coversheets folder
  - optional CLIN lookup workbook
  - overview output workbook

- Updated `03_F_and_R_script.py` to accept selected paths for:
  - overview workbook
  - PR folder
  - F&R output workbook
  - current option period

- Updated `04_build_file_script.py` to accept selected paths for:
  - source build workbook/template
  - approved coversheets folder
  - PR folder
  - build output workbook

- Updated `05_J1_script.py` to accept selected paths for:
  - build workbook
  - J.1 previous workbook
  - J.1 current output workbook
  - current option period override

- Updated `06_J17_file_script.py` to accept selected paths for:
  - source J.17 workbook
  - J.1 previous workbook
  - J.1 current workbook
  - billing detail workbook
  - J.17 output workbook
  - option period
  - current month

- Updated `07_MFR_walkthrough_script.py` to accept selected values for:
  - J.1 current workbook
  - Word document output path
  - option period
  - mod number

### Fixed

- Removed dependency on fixed file names such as:
  - `j1_previous_file.xlsx`
  - `clin_table_file.xls`
  - `build_file.xlsx`
  - `j17_file.xlsx`
  - `overview_file.xlsx`
  - `j1_current_file.xlsx`
  - `j17_updated_file.xlsx`
- Removed dependency on a specific EIS Billing Detail filename.
- Removed fixed `P00078` from the MFR walkthrough script.
- Replaced `.xlsx`-only folder discovery in patched paths with broader Excel file discovery.
- Replaced the hard-coded build script CLIN sheet name `CLIN Table 2024.10.04` with dynamic CLIN Table sheet detection.
- Updated J.1 categorization logic so it uses the selected/current option period instead of a fixed OP 5 / OP 6-11 split.
- Added fallback console prompts for environments where Tkinter file dialogs cannot launch.

### Validation

- Confirmed all patched Python files compile.
- Confirmed `--help` works for all seven workflow scripts.
- Full end-to-end run validation remains pending because the code package did not include the actual source Excel/Word workbooks needed to execute the workflow.

### Known limits

- `xlwings` is still required where the original scripts rely on Excel COM behavior.
- `openpyxl` cannot directly modify `.xlsb` workbooks; generated/modified outputs should generally be `.xlsx` or `.xlsm`.
- The scripts still run independently; a consolidated controller UI is a recommended next improvement.
- Preflight validation and schema enforcement are not yet fully implemented.
- User-facing error handling is improved by path selection, but workbook-specific validation messages should still be expanded.

## Earlier versions

Earlier package history was not included in this ZIP. This changelog starts with the path and filename flexibility refactor.
