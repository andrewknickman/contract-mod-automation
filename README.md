# Automated Contract Modifications

Current package: **v00.04.028**

This repository contains the controller UI and workflow scripts for generating contract modification support files:

1. Final coversheets
2. Overview workbook
3. F&R workbook
4. Build workbook
5. Current J.1 workbook
6. J.17 workbook
7. MFR walkthrough document

The package is prepared for replacing the existing GitHub repository. Generated workbooks, local run folders, caches, and local configuration files are intentionally excluded from version control.

## Start here

On Windows, install dependencies first:

```text
INSTALL_DEPENDENCIES.bat
```

Then launch the workflow UI:

```text
RUN_UI.bat
```

From a terminal:

```bash
python install_dependencies.py
python mod_automation_ui.py
```

## Dependency notes

The project uses:

- pandas
- numpy
- openpyxl
- pyxlsb
- xlrd
- python-docx
- pywin32 on Windows

Step 5 can accept a previous J.1 workbook in `.xlsx`, `.xlsm`, `.xlsb`, or `.xls` form. If the previous J.1 file is `.xlsb` or `.xls`, conversion requires Microsoft Excel installed on Windows plus `pywin32` in the same Python environment used by the UI/script.

To verify installed dependencies on Windows:

```text
VERIFY_DEPENDENCIES.bat
```

## Main UI

```bash
python mod_automation_ui.py
```

The UI lets the user select input files/folders, set the mod number/current option period/billing month, run individual steps, run the full workflow, and view live logs.

## Direct script usage

Each script can still be run directly.

```bash
python 02_overview_file_script.py --help
python 03_F_and_R_script.py --help
python 04_build_file_script.py --help
python 05_J1_script.py --help
```

Example Step 5 command:

```bash
python 05_J1_script.py \
  --build-file "runs/P00078/04_build/build_file.xlsx" \
  --j1-previous-file "J.1_Pricing_Spreadsheet_P00079.xlsb" \
  --output-file "runs/P00078/05_j1/j1_current_file.xlsx" \
  --current-opt-pd 5
```

## Current business rules captured in this package

### Overview F&R handling

- `F&R Needed = Yes` only when in-scope F&R Status is `Pending` or `Approved`.
- The adjacent F&R count only includes `Pending` or `Approved`.
- Blank or unknown F&R statuses are left blank for manual review rather than defaulting to `No`.
- If `F&R Needed` is blank, the adjacent F&R CLIN count is also blank.

### Build OpDiv approval fallback

- If a coversheet has explicit `Yes` / `No` approval decisions, the Build includes only `Yes` rows.
- If a coversheet has no explicit approval decisions, the Build includes all valid CLIN rows as pending-review build candidates.
- If OpDiv approval is missing or only unknown values are present, the Build includes all valid CLIN rows and logs a warning.
- Overview `# of CLINs` represents the current option period only. Build output can be larger because it can include all option periods.

### Build pricing-factor matching

Pricing method handling uses the intended business-facing methods:

- `ICB`
- `ORIG NSC`
- `TERM NSC`
- `ORIG JUR`
- `TERM JUR`
- `ORIG JUR-TERM JUR`
- `ORIG NSC-TERM NSC`

`ORIG JUR` / `TERM JUR` match against the PR/J.1 source columns `Orig CJID` / `Term CJID`.

### Catalog deduplication

- Catalog deduplication does **not** include Price Request Number in the duplicate lookup key.
- If duplicated Catalog items appear across multiple PRs, the item is collapsed and the unique PR numbers are concatenated, for example `59984/59993`.
- Blank or unmapped pricing-method rows use a conservative full-row identity so potentially valid rows are not silently collapsed by CLIN alone.

### J.1 generation

- `.xlsx` and `.xlsm` previous J.1 files are copied normally.
- `.xlsb` and `.xls` previous J.1 files are converted to `.xlsx` before `openpyxl` appends rows.
- The Step 5 path-scope error from v00.04.026 is fixed in v00.04.027 and carried forward here.

## Repository structure

```text
.
├── 01_final_coversheet_generation_script.py
├── 02_overview_file_script.py
├── 03_F_and_R_script.py
├── 04_build_file_script.py
├── 05_J1_script.py
├── 06_J17_file_script.py
├── 07_MFR_walkthrough_script.py
├── mod_automation_ui.py
├── file_selection.py
├── install_dependencies.py
├── requirements.txt
├── RUN_UI.bat
├── RUN_UI.command
├── INSTALL_DEPENDENCIES.bat
├── INSTALL_DEPENDENCIES.command
├── VERIFY_DEPENDENCIES.bat
├── docs/
│   ├── audits/
│   ├── diffs/
│   └── release-notes/
```

## What should not be committed

Do not commit generated run outputs, real PR files, coversheets, J.1 workbooks, F&R outputs, Build outputs, local config files, or any sensitive contract data. The `.gitignore` included here excludes common generated workbook and local-output patterns.

If a workbook template truly needs to be versioned, review it first for sensitive content and then force-add it intentionally.
