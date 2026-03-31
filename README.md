# Contract Modification Builder

## What This Project Is
This project is an Excel-based contract modification automation workflow. The current build wraps the existing six processing scripts in a PySide6 desktop UI so the workflow can be run from one guided interface instead of a terminal.

## What This Build Captures
This build reflects the cumulative updates made from the original source package to the current working state.

Key updates included in this build:
- removed the original machine-specific hardcoded path dependency from Scripts 01-06
- replaced fixed input/output locations with runtime selection
- added shared workflow input helpers in `workflow_io.py`
- added project documentation (`README.md`, `CHANGELOG.md`, `WORKFLOW.md`)
- added a PySide6 workflow UI (`ui_app.py`, `ui_runner.py`, `ui_workflow.py`)
- preserved the existing six-step business logic rather than redesigning the processing rules

## Project Files
Core workflow scripts:
- `01_final_coversheet_generation_script.py`
- `02_overview_file_script.py`
- `03_F_and_R_script.py`
- `04_build_file_script.py`
- `05_J1_script.py`
- `06_J17_file_script.py`

Shared helpers and UI files:
- `workflow_io.py`
- `ui_app.py`
- `ui_runner.py`
- `ui_workflow.py`

Documentation:
- `README.md`
- `CHANGELOG.md`
- `WORKFLOW.md`

## Requirements
- Windows
- Microsoft Excel installed
- Python 3.10+

## Required Python Packages
Install the required packages before running the project:

```bash
pip install pandas openpyxl numpy xlwings pyxlsb PySide6
```

## How to Run the UI
From the project folder:

```bash
python ui_app.py
```

The UI provides:
- workflow overview
- per-step input and output forms
- file and folder browsing
- reuse of previous step outputs where applicable
- step-by-step execution
- run logs and run history
- saved UI state in `ui_state.json`

## How the UI Works
The UI is the front end for the existing six workflow steps. Each step page collects the required inputs for that step and runs the corresponding underlying script.

The UI does not redesign the business logic. It drives the existing workflow in this order:
1. Generate Coversheets
2. Generate Overview Workbook
3. Generate F&R Workbook
4. Generate Build Workbook
5. Generate Updated J.1 Workbook
6. Generate Updated J.17 Workbook

## Script-by-Script Inputs and Outputs
### 1. Coversheets
Inputs:
- source J.1 workbook
- CLIN lookup workbook
- PR workbook folder
- coversheet output folder
- current option period

Output:
- coversheet workbooks

### 2. Overview Workbook
Inputs:
- coversheet workbook folder
- CLIN lookup workbook
- overview output workbook path

Output:
- overview workbook

### 3. F&R Workbook
Inputs:
- overview workbook
- PR workbook folder
- F&R output workbook path

Output:
- F&R workbook

### 4. Build Workbook
Inputs:
- build template workbook
- coversheet workbook folder
- PR workbook folder
- build output workbook path

Output:
- build workbook

### 5. Updated J.1 Workbook
Inputs:
- build workbook
- source J.1 previous workbook
- updated J.1 output workbook path

Output:
- updated J.1 workbook

Run note:
- The source J.1 workbook for Step 5 should be a true `.xlsx` or `.xlsm` workbook. The current Step 5 workflow does not safely update a `.xlsb` source workbook.

### 6. Updated J.17 Workbook
Inputs:
- source J.17 workbook
- updated J.1 workbook
- billing workbook
- updated J.17 output workbook path
- current option period
- billing month

Output:
- updated J.17 workbook

## Runtime State Files
- `.workflow_last_paths.json` stores last-used paths for the shared file selection helpers
- `ui_state.json` stores saved UI values, statuses, and run history

## Notes
- The UI uses the existing workflow logic. It is a guided front end for the current process.
- Each step can be rerun individually.
- Downstream steps can reuse outputs from earlier steps through the UI.
