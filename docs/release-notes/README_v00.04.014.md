# Mod Automation Project

This project automates the contract modification package workflow across coversheets, overview workbook, F&R workbook, build workbook, J.1 workbook, J.17 workbook, and MFR walkthrough output.

This version removes the prior dependency on one local `D:\...\Data` folder and fixed input/output filenames. Scripts can now accept paths by command line or prompt the user to select files and folders through the file system.

## What the project does

The workflow is organized into seven scripts:

1. **Generate coversheets** from the prior J.1 workbook, CLIN lookup workbook, and PR files.
2. **Generate the overview workbook** from completed coversheets.
3. **Generate the F&R workbook** from the overview workbook and PR files.
4. **Generate the build workbook** from the source build workbook/template, approved coversheets, and PR files.
5. **Generate the updated J.1 workbook** from the build workbook and prior J.1 workbook.
6. **Generate the updated J.17 workbook** from J.17, prior/current J.1, and billing detail files.
7. **Generate the MFR walkthrough document** from the updated J.1 workbook.

The scripts are still independent, but they are intended to be run in sequence when producing a full modification package.

## Project files

| File | Purpose |
| --- | --- |
| `01_final_coversheet_generation_script.py` | Creates individual coversheet workbooks. |
| `02_overview_file_script.py` | Builds the overview workbook from coversheets. |
| `03_F_and_R_script.py` | Builds the F&R workbook and PR tabs. |
| `04_build_file_script.py` | Builds the main build workbook from approved package materials. |
| `05_J1_script.py` | Creates the current J.1 workbook from the build workbook and previous J.1. |
| `06_J17_file_script.py` | Updates the J.17 workbook using J.1 and billing detail files. |
| `07_MFR_walkthrough_script.py` | Generates the MFR walkthrough Word document. |
| `file_selection.py` | Shared file/folder picker and Excel helper utilities. |
| `mod_automation_ui.py` | Desktop controller UI for selecting inputs, running preflight checks, and launching the full workflow. |
| `RUN_UI.bat` | Windows launcher for the controller UI. |
| `RUN_UI.command` | macOS/Linux launcher for the controller UI. |
| `PATH_FLEXIBILITY_REPORT.md` | Detailed report of removed hard-coded paths and filename assumptions. |
| `path_flexibility.diff` | Patch diff showing the path/file-selection refactor. |
| `REGRESSION_AUDIT.md` | Comparison notes confirming original script functions were preserved and documenting the controller preflight fix. |
| `OVERVIEW_FR_COUNT_AUDIT.md` | Documents the Overview F&R yes/no and numeric count rules. |
| `OVERVIEW_FR_BLANK_STATUS_AUDIT_v00.04.014.md` | Documents blank F&R Status handling for the Overview. |

## Requirements

Use Python 3.10+ if possible.

Install the project dependencies:

```bash
pip install pandas numpy openpyxl xlwings python-docx pyxlsb
```

Notes:

- `xlwings` requires Microsoft Excel on the machine where the relevant scripts run.
- `.xlsb` files can be read through `pyxlsb` in the patched pandas paths.
- Workbooks modified directly through `openpyxl` should still be saved as `.xlsx` or `.xlsm`; `openpyxl` cannot directly modify `.xlsb` files.
- File picker dialogs use Tkinter. If a dialog cannot open, the helper falls back to console prompts.


## Current Overview F&R Status rule

For the generated Overview workbook, `02_overview_file_script.py` evaluates the in-scope rows from `CLIN Table (Current OP)`. If one or more rows are marked with OpDiv approval, only those approved rows are evaluated. If no rows are marked with OpDiv approval, all listed CLIN rows are evaluated.

The current F&R behavior is:

| In-scope F&R Status values | Overview `F&R Needed` | Overview F&R count |
| --- | --- | ---: |
| All blank | Blank | 0 |
| `Pending` | `Yes` | Counted |
| `Approved` | `Yes` | Counted |
| `N/A`, `NA`, `N.A.`, or `Not Applicable` only | `No` | 0 |

This prevents blank coversheet F&R statuses from being treated as confirmed `No` determinations.

## Running the scripts

Each script supports two modes:

1. **Interactive file selection:** run the script without path arguments and select files/folders when prompted.
2. **CLI-driven execution:** pass file and folder paths as arguments for repeatable runs.

You can view available options for any script with:

```bash
python script_name.py --help
```

## Running the controller UI

The recommended way to run the workflow is now through the controller UI:

```bash
python mod_automation_ui.py
```

On Windows, users can also double-click:

```text
RUN_UI.bat
```

The UI lets a user select the files and folders needed for the full package process, set the mod number and option period, generate default output names, run preflight checks, run one step at a time, or run the full seven-step workflow. It streams each script's output into a live log and tracks step status as Pending, Running, Complete, Failed, or Skipped.

### UI workflow

1. Open `mod_automation_ui.py` or run `RUN_UI.bat`.
2. On the **Setup** tab, select the required source files and folders.
3. Enter the mod number, current option period, and billing month if needed. The billing month defaults to `December` to match the original J.17 script behavior; change it when processing a different cycle.
4. Select or create the run output folder.
5. Click **Set default output names** to create standard output paths for the run.
6. Click **Preflight all steps** to check for missing files and folders. The controller understands that some files are created by earlier steps and then consumed by later steps during a full workflow run.
7. Use **Run full workflow** or run individual steps from the **Run Workflow** tab.
8. Review the **Log** tab for script output and errors.

### UI-managed outputs

By default, the controller creates a clean run structure under the selected output folder:

```text
run-output/
  01_generated_coversheets/
  02_overview/overview_file.xlsx
  03_f_and_r/f_r_output.xlsx
  04_build/build_file.xlsx
  05_j1/j1_current_file.xlsx
  06_j17/j17_updated_file.xlsx
  07_mfr/<MOD> MFR Walkthrough Output.docx
```

The defaults are only suggestions. Every output path remains editable in the UI.

### UI configuration

The controller can save and reload a JSON configuration file so repeated runs do not require users to re-select every input. When a full workflow completes, the UI also saves `mod_automation_ui_config.json` in the project folder as the most recent configuration.

## Typical full workflow

### 1. Generate coversheets

```bash
python 01_final_coversheet_generation_script.py \
  --j1-previous-file "path/to/J1 previous.xlsx" \
  --clin-table-file "path/to/CLIN lookup.xls" \
  --pr-dir "path/to/PR files" \
  --output-dir "path/to/output/coversheets" \
  --current-op 5
```

### 2. Generate overview workbook

```bash
python 02_overview_file_script.py \
  --coversheets-dir "path/to/output/coversheets" \
  --clin-table-file "path/to/CLIN lookup.xls" \
  --output-file "path/to/overview_file.xlsx"
```

The CLIN table file is optional for this step. If skipped, the script can still run and leave Services blank.

### 3. Generate F&R workbook

```bash
python 03_F_and_R_script.py \
  --overview-file "path/to/overview_file.xlsx" \
  --pr-dir "path/to/PR files" \
  --output-file "path/to/f_r_output.xlsx" \
  --current-option-period 5
```

### 4. Generate build workbook

```bash
python 04_build_file_script.py \
  --build-file "path/to/source build.xlsx" \
  --coversheets-dir "path/to/approved coversheets" \
  --pr-dir "path/to/PR files" \
  --output-file "path/to/build_file.xlsx"
```

The script now looks for a CLIN Table sheet dynamically instead of requiring a date-stamped sheet name.

### 5. Generate current J.1 workbook

```bash
python 05_J1_script.py \
  --build-file "path/to/build_file.xlsx" \
  --j1-previous-file "path/to/J1 previous.xlsx" \
  --output-file "path/to/j1_current_file.xlsx" \
  --current-opt-pd 5
```

If `--current-opt-pd` is omitted, the script calculates the current option period by date.

### 6. Generate updated J.17 workbook

```bash
python 06_J17_file_script.py \
  --j17-file "path/to/J17 source.xlsx" \
  --j1-previous-file "path/to/J1 previous.xlsx" \
  --j1-current-file "path/to/j1_current_file.xlsx" \
  --billing-file "path/to/EIS Billing Detail.xlsx" \
  --output-file "path/to/j17_updated_file.xlsx" \
  --option-period 5 \
  --current-month March
```

If `--current-month` is omitted, the script attempts to infer the month from the billing filename.

### 7. Generate MFR walkthrough document

```bash
python 07_MFR_walkthrough_script.py \
  --j1-current-file "path/to/j1_current_file.xlsx" \
  --output-file "path/to/P00078 MFR Walkthrough Output.docx" \
  --option-period 5 \
  --mod-number P00078
```

## Accepted Excel file types

The shared helper recognizes these Excel extensions where the script logic supports them:

- `.xlsx`
- `.xlsm`
- `.xlsb`
- `.xls`

Important distinction: some inputs can be read from broader Excel formats, but generated or modified workbooks should generally be saved as `.xlsx` or `.xlsm`.

## Path and filename flexibility added in this version

The following assumptions were removed or reduced:

- One fixed local data folder.
- Fixed `input/`, `output/`, `coversheets/`, and `pr_files/` locations.
- Fixed input names such as `j1_previous_file.xlsx`, `clin_table_file.xls`, `build_file.xlsx`, and `j17_file.xlsx`.
- Fixed billing detail filename.
- Fixed output names such as `overview_file.xlsx`, `build_file.xlsx`, `j1_current_file.xlsx`, and `j17_updated_file.xlsx`.
- Fixed MFR mod number `P00078`.
- `.xlsx`-only folder discovery in places that can support additional Excel formats.
- Date-stamped CLIN Table sheet name assumption in the build script.

## Recommended folder organization

The scripts no longer require this exact layout, but this structure is still useful for clean runs:

```text
mod-package-run/
  inputs/
    J1 previous.xlsx
    CLIN lookup.xls
    source build.xlsx
    J17 source.xlsx
    EIS Billing Detail.xlsx
  pr_files/
    *.xlsx / *.xlsm / *.xlsb / *.xls
  coversheets/
  approved_coversheets/
  outputs/
```

## Validation performed for this package

The patched scripts were checked with Python compilation and `--help` execution. Full workflow validation still requires the actual Excel and Word inputs used by the mod process.

## Known limits / next improvements

Recommended next improvements:

- Expand workbook-level preflight validation before each script runs.
- Add a schema layer for required sheets/columns in each workbook.
- Add more workbook-specific user-facing error messages that explain what file/sheet/column failed.
## F&R workbook inclusion logic

The F&R workbook is not currently generated by scanning every coversheet directly. It is driven by the generated Overview workbook.

A PR is selected for F&R workbook processing when all of these are true:

1. The Overview workbook contains a row where `F&R Needed` is `Yes`.
2. The PR files folder contains an Excel file whose filename starts with the PR identifier, such as `PR59483...`.
3. That PR workbook contains a worksheet whose name includes `comp`, such as `Comps` or `COMP`.
4. The comp worksheet contains rows that the script can extract into F&R records.

The Overview script determines `F&R Needed` from the `CLIN Table (Current OP)` sheet. In `v00.04.010`, if no CLIN rows are marked with OpDiv approval, the Overview script now falls back to the listed CLIN rows and their F&R statuses instead of automatically producing zero CLINs / no F&R need. In `v00.04.011`, the F&R yes/no flag and F&R count were aligned to use the same row scope. In `v00.04.012`, the numeric F&R count was corrected to include only `Pending` and `Approved` statuses. In `v00.04.013`, the `F&R Needed` flag was aligned to that same business rule: only `Pending` or `Approved` in-scope rows trigger `F&R Needed = Yes`; `N/A` / `NA` does not trigger the flag and does not count as an F&R item.

Common reasons a PR may not receive an F&R tab:

- The Overview row is not marked `F&R Needed = Yes`.
- The PR workbook filename does not start with the PR number format expected by the script.
- The PR workbook does not have a sheet with `comp` in the sheet name.
- The comp sheet is empty or missing expected columns.
- Multiple files/versions exist for the same PR; the current script groups tabs primarily by PR number.

