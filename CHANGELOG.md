# CHANGELOG

## v00.04.003

### Update Scope
Expand build with an input/output selection helper, remove hard-coded file and folder dependencies, and .

### Overall Updates
- Added `workflow_io.py` as a shared runtime input/output selection helper
- Removed the original machine-specific hardcoded base path dependency from Scripts 01-06
- Replaced hardcoded file paths, folder paths, and save targets with runtime prompts
- Added runtime prompting for option period and billing month where applicable
- Added lean documentation for project setup, run order, workflow, and cumulative updates
- Added a run note for Script 05 clarifying that the source J.1 workbook should be `.xlsx` or `.xlsm` for the current update flow

### Per-File Updates
- `workflow_io.py`
  - Added shared helper methods for file selection, folder selection, save selection, integer prompts, and text prompts
  - Added remembered last-used paths through `.workflow_last_paths.json`

- `01_final_coversheet_generation_script.py`
  - Replaced hardcoded J.1 workbook path with runtime file selection
  - Replaced hardcoded CLIN lookup workbook path with runtime file selection
  - Replaced hardcoded PR folder path with runtime directory selection
  - Replaced hardcoded coversheet output folder with runtime directory selection
  - Added runtime prompt for current option period

- `02_overview_file_script.py`
  - Replaced hardcoded coversheets folder path with runtime directory selection
  - Replaced hardcoded CLIN lookup workbook path with runtime file selection
  - Replaced hardcoded overview output path with runtime save selection

- `03_F_and_R_script.py`
  - Replaced hardcoded overview workbook path with runtime file selection
  - Replaced hardcoded PR folder path with runtime directory selection
  - Replaced hardcoded F&R output path with runtime save selection

- `04_build_file_script.py`
  - Replaced hardcoded build template path with runtime file selection
  - Replaced hardcoded coversheets folder path with runtime directory selection
  - Replaced hardcoded PR folder path with runtime directory selection
  - Replaced hardcoded build output path with runtime save selection

- `05_J1_script.py`
  - Replaced hardcoded build workbook path with runtime file selection
  - Replaced hardcoded source J.1 workbook path with runtime file selection
  - Replaced hardcoded updated J.1 output path with runtime save selection
  - Documentation now notes that the current update flow should use `.xlsx` or `.xlsm` for the source J.1 workbook

- `06_J17_file_script.py`
  - Replaced hardcoded source J.17 workbook path with runtime file selection
  - Replaced hardcoded updated J.1 workbook path with runtime file selection
  - Replaced hardcoded billing workbook path with runtime file selection
  - Replaced hardcoded updated J.17 output path with runtime save selection
  - Added runtime prompt for current option period
  - Added runtime prompt for billing month

- `README.md`
  - Added project description
  - Added setup guide
  - Added package installation instructions
  - Added script order and run commands
  - Added script-by-script prompt and output summary
  - Added cumulative update summary for the current build

- `WORKFLOW.md`
  - Added workflow layout for the six-script pipeline
  - Added input/output mapping for each script
  - Added Script 05 source-workbook format note
