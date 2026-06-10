# Changelog - v00.04.027

## Fixed

- Fixed Step 5 J.1 generation failure caused by `j1_current_file` being referenced before assignment.
- Corrected the variable scope introduced by the `.xlsb` / `.xls` conversion change so the prepared current J.1 output path is available before the script creates the output folder and prepares the workbook.

## Preserved

- `.xlsx` / `.xlsm` previous J.1 files still copy normally.
- `.xlsb` / `.xls` previous J.1 files still attempt real Excel conversion to `.xlsx` before `openpyxl` appends rows.
- Dependency installer files from `v00.04.026` are unchanged.
- Build, Overview, F&R, Catalog, J.17, MFR, and controller UI logic are unchanged.

## Validation

- Python compilation passed for `05_J1_script.py`.
- CLI help check passed for `05_J1_script.py`.
