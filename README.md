# Contract Modification Builder

## What This Project Is
This project is a Python-based Excel automation workflow for Contract Modifications. It currently runs as a six-script pipeline. Each script performs one stage of the workflow and produces an output workbook used by a later stage.

## What This Current Build Represents
This repo state is intended to capture the full set of updates made so far from the original source package to the current working build.

The original source package contained only the six Python scripts and relied on machine-specific hardcoded paths and workbook names.

The current build adds:
- runtime file, folder, and save-location selection through `workflow_io.py`
- project documentation (`README.md`, `CHANGELOG.md`, `WORKFLOW.md`)
- setup and run guidance for the current workflow

## Project Files
- `01_final_coversheet_generation_script.py`
- `02_overview_file_script.py`
- `03_F_and_R_script.py`
- `04_build_file_script.py`
- `05_J1_script.py`
- `06_J17_file_script.py`
- `workflow_io.py`
- `README.md`
- `CHANGELOG.md`
- `WORKFLOW.md`

## Requirements
- Windows
- Microsoft Excel installed
- Python 3.10+

## Required Python Packages
Install the required packages before running the scripts:

```bash
pip install pandas openpyxl numpy xlwings pyxlsb
```

## Setup
1. Place the project files in a working folder on your machine.
2. Open a terminal in the project folder.
3. Install the required Python packages.
4. Make sure the source Excel workbooks needed for your run are available on your machine.
5. Run the scripts from the project folder so `workflow_io.py` and `.workflow_last_paths.json` can be found.

## Runtime File Selection
The scripts no longer rely on a fixed machine-specific base directory. They now use `workflow_io.py` to prompt the user for required files, folders, and output save locations at runtime.

Current helper methods in `workflow_io.py`:
- `choose_file(...)`
- `choose_files(...)`
- `choose_directory(...)`
- `choose_save_file(...)`
- `ask_integer(...)`
- `ask_text(...)`

`workflow_io.py` also stores last-used paths in `.workflow_last_paths.json` in the project folder.

## Script Order
Run the scripts in this order unless you are intentionally starting from an intermediate output:
1. `01_final_coversheet_generation_script.py`
2. `02_overview_file_script.py`
3. `03_F_and_R_script.py`
4. `04_build_file_script.py`
5. `05_J1_script.py`
6. `06_J17_file_script.py`

## Run Guide
Run each script from the project folder.

### 1. Coversheet Generation
```bash
python 01_final_coversheet_generation_script.py
```
Prompts for:
- source J.1 workbook
- CLIN lookup workbook
- PR workbook folder
- coversheet output folder
- current option period

Produces:
- coversheet workbooks

### 2. Overview File
```bash
python 02_overview_file_script.py
```
Prompts for:
- coversheet workbook folder
- CLIN lookup workbook
- overview output file location

Produces:
- overview workbook

### 3. F&R File
```bash
python 03_F_and_R_script.py
```
Prompts for:
- overview workbook
- PR workbook folder
- F&R output file location

Produces:
- F&R workbook

### 4. Build File
```bash
python 04_build_file_script.py
```
Prompts for:
- build template workbook
- coversheet workbook folder
- PR workbook folder
- build output file location

Produces:
- build workbook

### 5. Updated J.1 File
```bash
python 05_J1_script.py
```
Prompts for:
- build workbook
- source J.1 workbook
- updated J.1 output file location

Produces:
- updated J.1 workbook

**Run note:** The source J.1 workbook used by Script 05 should be a true `.xlsx` or `.xlsm` workbook if it needs to be opened and updated with `openpyxl`. A `.xlsb` source file can be read in some contexts, but the current Script 05 workflow does not safely support updating a `.xlsb` workbook.

### 6. Updated J.17 File
```bash
python 06_J17_file_script.py
```
Prompts for:
- source J.17 workbook
- updated J.1 workbook
- billing workbook
- updated J.17 output file location
- current option period
- billing month

Produces:
- updated J.17 workbook

## Summary of Updates Captured in This Build
- added `workflow_io.py`
- removed the original hardcoded machine-specific path dependency from Scripts 01-06
- replaced hardcoded input/output workbook locations with runtime selection
- added runtime prompting for option period and billing month where applicable
- added lean documentation for setup, workflow, and cumulative update tracking
