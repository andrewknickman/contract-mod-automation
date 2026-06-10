# Changelog — v00.04.019

## v00.04.019 — Build coversheet schema handling fix

### Fixed

- Updated `04_build_file_script.py` so the Build step no longer treats every Excel file in the selected coversheet folder as a coversheet.
- Added workbook sheet-name detection so PR/J.1 workbooks are skipped when the selected coversheet folder is a mixed directory containing both coversheets and PR files.
- Added accepted coversheet CLIN sheet-name variants for `CLIN Table (Current OP)` and `CLIN Table (OY)`.
- Replaced the hard-coded OpDiv approval column lookup with normalized header matching.
- Accepted OpDiv approval header variants such as `OpDiv Approval (Yes/No)`, `OpDiv Approval (Yes / No)`, `OpDiv Approval (Y/N)`, `OpDiv Approval`, and `OpDiv Approved`.
- Added safe handling when an otherwise coversheet-like workbook is missing an OpDiv approval column: the script now logs a warning and skips that workbook instead of crashing with `KeyError`.

### Preserved

- Build logic still only uses rows where OpDiv approval is `Yes`.
- Build logic still matches approved coversheet rows against the corresponding PR/J.1 workbook.
- J.1 matching logic was not changed in this package.
- F&R source-header and J.1 fixes from `v00.04.018` were not changed.
- Overview F&R handling from `v00.04.015` was not changed.

### Validation

- Confirmed `04_build_file_script.py` compiles with `python -m py_compile`.
- Ran targeted tests for OpDiv approval header variants and non-coversheet workbook skipping.
