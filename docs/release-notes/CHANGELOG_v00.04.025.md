# Changelog — v00.04.025

## v00.04.025 — J.1 previous workbook `.xlsb` conversion fix

### Fixed

- Fixed Step 5 failing with `OSError: File contains no valid workbook part` when the selected previous J.1 workbook was `.xlsb` and the output path was `.xlsx`.
- Replaced direct byte-copy behavior for `.xlsb` / `.xls` previous J.1 files with a real conversion step.
- Added validation so the script confirms the prepared J.1 output workbook is a valid `.xlsx` before trying to append rows.
- Added a clearer error message explaining that `.xlsb` / `.xls` conversion requires Microsoft Excel + `pywin32`, or a manually saved `.xlsx` copy.

### Changed

- Step 5 output is now forced to `.xlsx` because `openpyxl` performs the row appends.
- `.xlsx` and `.xlsm` previous J.1 files continue to be copied directly before modification.

### Unchanged

- Build/Catalog deduplication and PR-number concatenation from `v00.04.024` are unchanged.
- Catalog-to-J.1 row mapping is unchanged.
- Current/future option-period categorization is unchanged.

### Validation

- `05_J1_script.py` compiles successfully.
- Conversion helper paths were checked for `.xlsx`, `.xlsm`, `.xlsb`, and `.xls` branch behavior.
- Full `.xlsb` conversion could not be executed in this Linux sandbox because it requires desktop Microsoft Excel on Windows.
