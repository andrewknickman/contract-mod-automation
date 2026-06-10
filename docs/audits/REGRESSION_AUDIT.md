# Regression Audit - v00.04.009
This audit compares the original uploaded seven-script project against the controller UI package and the v00.04.009 fixes.
## Summary
- All seven original processing scripts are still present.
- No original top-level processing functions were removed from the seven scripts.
- The controller UI runs the existing scripts through their CLI arguments; it does not replace the workflow logic.
- Two issues were found and fixed in this package: full-workflow preflight validation for generated intermediate outputs, and the J.17 default current-month behavior.

## File presence
- `01_final_coversheet_generation_script.py`: present
- `02_overview_file_script.py`: present
- `03_F_and_R_script.py`: present
- `04_build_file_script.py`: present
- `05_J1_script.py`: present
- `06_J17_file_script.py`: present
- `07_MFR_walkthrough_script.py`: present

## Function preservation check
### `01_final_coversheet_generation_script.py`
- Original functions: 19
- Current functions: 21
- Missing original functions: None
- Added wrapper/helper functions: ['parse_args', 'configure_paths']
### `02_overview_file_script.py`
- Original functions: 17
- Current functions: 19
- Missing original functions: None
- Added wrapper/helper functions: ['parse_args', 'configure_paths']
### `03_F_and_R_script.py`
- Original functions: 9
- Current functions: 12
- Missing original functions: None
- Added wrapper/helper functions: ['parse_args', 'configure_paths', 'main']
### `04_build_file_script.py`
- Original functions: 26
- Current functions: 28
- Missing original functions: None
- Added wrapper/helper functions: ['parse_args', 'configure_paths']
### `05_J1_script.py`
- Original functions: 11
- Current functions: 13
- Missing original functions: None
- Added wrapper/helper functions: ['parse_args', 'configure_paths']
### `06_J17_file_script.py`
- Original functions: 17
- Current functions: 20
- Missing original functions: None
- Added wrapper/helper functions: ['extract_month_name_from_filename', 'parse_args', 'configure_paths']
### `07_MFR_walkthrough_script.py`
- Original functions: 14
- Current functions: 16
- Missing original functions: None
- Added wrapper/helper functions: ['parse_args', 'configure_paths']

## Issues fixed after comparison
### 1. Controller preflight incorrectly handled intermediate outputs
The UI used the same path field for outputs and later inputs. Before this fix, a full workflow could incorrectly require generated coversheets or generated workbooks to already exist before Step 1 ran. Conversely, running a later step by itself could pass preflight even when a required prior output did not exist.

Fixed behavior:
- During a full workflow, preflight recognizes outputs planned by earlier selected steps.
- During a single-step run, preflight requires needed prior-output files/folders to already exist.
- Output folders/files for the current step are still treated as creatable save locations.

### 2. J.17 current-month default changed from original behavior
The original `06_J17_file_script.py` defaulted `current_month` to `December`. The controller package previously inferred a month from the billing filename when the UI month field was blank, which was a business-logic change rather than only a path-selection change.

Fixed behavior:
- `06_J17_file_script.py` now defaults `--current-month` to `December`.
- The UI also defaults the billing/current month field to `December`.
- Users can still override the month through the UI or CLI.

## Remaining intentional changes
The following are intentional flexibility changes from v00.04.007/v00.04.008, not removed functionality:
- Hard-coded local root path replaced with selected file/folder paths.
- Fixed input/output filenames replaced with selected files and save locations.
- `.xlsx`-only discovery expanded where pandas/openpyxl/xlwings support it.
- MFR mod number is now selectable instead of fixed to `P00078`.
- Current option period can be selected instead of requiring code edits.

## Validation performed
- Ran Python compilation across all project `.py` files.
- Checked CLI help for the J.17 script after restoring the default month behavior.
- Full end-to-end workbook validation still requires the real Excel and Word inputs used by the mod workflow.
