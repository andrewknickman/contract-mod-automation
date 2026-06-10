# Path and Filename Flexibility Report

## What was hard-coded before

The project had a common hard-coded local root across all seven scripts:

```text
D:\1_emerjence_work\04_HHS\09_new_mod_automation\Data
```

From that root, scripts assumed fixed folders and filenames such as:

- `input/j1_previous_file.xlsx`
- `input/clin_table_file.xls`
- `input/build_file.xlsx`
- `input/j17_file.xlsx`
- `input/EIS Billing Detail - MAR 2026 - HHS EIS PMO 75P00120F80177.xlsx`
- `output/overview_file.xlsx`
- `output/build_file.xlsx`
- `output/j1_current_file.xlsx`
- `output/j17_updated_file.xlsx`
- `output/coversheets/`
- `pr_files/`
- `P00078` for the MFR walkthrough script

Several scripts also hard-coded `.xlsx` discovery only, which meant `.xlsm`, `.xlsb`, or `.xls` files were skipped even when pandas could read them. The build script also used a date-stamped sheet name, `CLIN Table 2024.10.04`, instead of finding the CLIN Table sheet dynamically.

## What changed

A shared helper module was added:

- `file_selection.py`

It provides:

- File picker fallback through Tkinter dialogs when paths are not passed by command line.
- Console prompt fallback when dialogs are unavailable.
- Folder picker support.
- Save-as picker support.
- Shared Excel extension handling for `.xlsx`, `.xlsm`, `.xlsb`, and `.xls` where pandas-based reading is used.
- Shared `read_excel_auto()` support for `.xlsb` via `pyxlsb`.
- Shared Excel folder scanning through `iter_excel_files()`.
- Shared PR matching through `find_excel_by_prefix()`.

## Script-by-script updates

### 01_final_coversheet_generation_script.py

Removed hard-coded:

- Base data directory
- J.1 previous workbook name
- CLIN lookup workbook name
- PR folder path
- coversheet output folder path
- `.xlsx`-only PR discovery

Added selectable/CLI inputs:

```bash
python 01_final_coversheet_generation_script.py \
  --j1-previous-file "path/to/J1 previous.xlsx" \
  --clin-table-file "path/to/CLIN lookup.xls" \
  --pr-dir "path/to/PR files" \
  --output-dir "path/to/output/coversheets" \
  --current-op 5
```

### 02_overview_file_script.py

Removed hard-coded:

- Base data directory
- coversheet folder path
- CLIN lookup workbook name
- overview output workbook name
- `.xlsx`-only coversheet discovery

Added selectable/CLI inputs:

```bash
python 02_overview_file_script.py \
  --coversheets-dir "path/to/coversheets" \
  --clin-table-file "path/to/CLIN lookup.xls" \
  --output-file "path/to/overview_file.xlsx"
```

The CLIN table is now optional. If skipped, the script still runs and leaves Services blank.

### 03_F_and_R_script.py

Removed hard-coded:

- Base data directory
- overview workbook path
- PR folder path
- F&R output workbook path
- `.xlsx`-only PR discovery

Added selectable/CLI inputs:

```bash
python 03_F_and_R_script.py \
  --overview-file "path/to/overview_file.xlsx" \
  --pr-dir "path/to/PR files" \
  --output-file "path/to/f_r_output.xlsx" \
  --current-option-period 5
```

Also added `.xlsb` read support for J.1 rate lookup through pandas because openpyxl cannot read binary Excel files.

### 04_build_file_script.py

Removed hard-coded:

- Base data directory
- build template filename
- build output filename
- coversheet folder path
- PR folder path
- `.xlsx`-only PR matching
- `.xlsx`-only coversheet discovery
- date-stamped CLIN Table sheet name `CLIN Table 2024.10.04`

Added selectable/CLI inputs:

```bash
python 04_build_file_script.py \
  --build-file "path/to/source build.xlsx" \
  --coversheets-dir "path/to/approved coversheets" \
  --pr-dir "path/to/PR files" \
  --output-file "path/to/build_file.xlsx"
```

The CLIN Table sheet is now located by finding a sheet whose name contains `CLIN Table`.

### 05_J1_script.py

Removed hard-coded:

- Base data directory
- build workbook path
- J.1 previous workbook path
- J.1 output workbook path
- hard-coded TO Period 5 / 6-11 split in categorization

Added selectable/CLI inputs:

```bash
python 05_J1_script.py \
  --build-file "path/to/build_file.xlsx" \
  --j1-previous-file "path/to/J1 previous.xlsx" \
  --output-file "path/to/j1_current_file.xlsx" \
  --current-opt-pd 5
```

The script still calculates the current option period by date by default, but it can now be overridden. The categorization now uses the selected/current option period instead of a fixed OP 5 split.

### 06_J17_file_script.py

Removed hard-coded:

- Base data directory
- J.17 source workbook name
- J.1 previous workbook name
- J.1 current workbook name
- specific EIS Billing Detail filename
- J.17 output workbook name
- hard-coded current month of `December`

Added selectable/CLI inputs:

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

If `--current-month` is omitted, the script tries to infer it from the billing filename.

### 07_MFR_walkthrough_script.py

Removed hard-coded:

- Base data directory
- J.1 current workbook path
- output Word document path
- fixed `P00078` mod number

Added selectable/CLI inputs:

```bash
python 07_MFR_walkthrough_script.py \
  --j1-current-file "path/to/j1_current_file.xlsx" \
  --output-file "path/to/P00078 MFR Walkthrough Output.docx" \
  --option-period 5 \
  --mod-number P00078
```

## Notes and limits

- File selection now happens at runtime. You can provide command-line paths for repeatable runs, or omit paths and select files/folders through the file picker.
- pandas-readable Excel inputs support `.xlsx`, `.xlsm`, `.xlsb`, and `.xls` where the required engine is installed.
- Workbooks that are modified with openpyxl should still be saved as `.xlsx` or `.xlsm`. Binary `.xlsb` files can be read by pandas in the patched paths, but openpyxl cannot directly modify `.xlsb` files.
- `xlwings` is still required by the overview/F&R scripts where the original code depended on Excel COM behavior.
