
## v00.04.007

### Intention
Fix the Step 2 UI-run regression caused by the child process environment and harden UI-launched workflow steps against missing home-directory variables.

### Overall Updates
- Fixed the UI process launch so workflow steps inherit the full system environment instead of a stripped environment.
- Added a runtime environment guard in `ui_runner.py` to ensure `HOME`, `USERPROFILE`, and `MPLCONFIGDIR` are available for libraries such as `matplotlib` when a step is launched from the UI.
- Preserved the existing step business logic; this fix is in the UI execution wrapper only.

### Per-File Updates
- `ui_app.py`
  - Changed the step-launch environment setup to use the full system environment before adding `PYTHONUNBUFFERED`.
- `ui_runner.py`
  - Added environment normalization for home-directory and Matplotlib config variables before loading and running a workflow step.

## v00.04.006

### Intention
Improve workflow clarity and running feedback in the PySide6 UI.

### Overall Updates
- replaced vague prerequisites text with actual upstream dependency and required-input summaries on each step page
- added live running feedback with an animated running indicator and an indeterminate progress bar while a step is executing
- improved step log behavior so output updates in real time during a run

### Per-File Updates
- `ui_app.py`
  - rewrote the prerequisites panel to show upstream step status, linked outputs, and required inputs for the current step
  - added a running indicator, animated state text, and an indeterminate progress bar
  - updated process launching to use unbuffered Python output for better real-time log updates
  - updated stdout/stderr handling to stream log output into the UI while the process is still running

## v00.04.005

### Intention
Improve the PySide6 UI formatting and add practical workflow quality-of-life features.

### Overall Updates
- Fixed major UI contrast issues, including white text on white backgrounds
- Added clearer visual hierarchy and button styling for primary vs secondary actions
- Added open/copy path actions directly in step fields
- Added clearer validation presentation, status bar updates, and improved table readability
- Added small usability refinements such as clearable text fields and richer tooltips

### Per-File Updates
- `ui_app.py`
  - improved global styling and contrast
  - added secondary/primary button roles
  - added open/copy path actions for file and folder fields
  - improved history table readability
  - added status bar run/error/result messages
  - improved field and log placeholders

# Changelog

## v00.04.004

### Intention
Add a usable PySide6 workflow UI on top of the current six-step automation so the project can be run from a guided desktop interface instead of a terminal.

### Overall Updates
- added a PySide6 desktop UI shell for the workflow
- added step pages for all six workflow stages
- added workflow overview, run history, settings, and help pages
- added UI-driven input collection for files, folders, save locations, integers, and text values
- added step status tracking, run logs, and run history persistence
- added previous-output reuse for downstream step inputs
- kept the current six-step business logic intact and ran it through the new UI layer
- updated README to explain how to launch and use the UI

### Per-File Updates
- `ui_app.py`
  - added the main PySide6 application window
  - added workflow navigation, step pages, overview page, run history page, settings page, and help page
  - added validation, logging, and step execution from the UI
- `ui_runner.py`
  - added the runner used by the UI to execute existing workflow steps with injected UI responses
- `ui_workflow.py`
  - added the shared metadata for step layout, fields, dependencies, and output reuse
- `README.md`
  - updated setup and run guidance to include the PySide6 UI

## v00.04.003

### Intention
Capture the cumulative updates made from the original source package to the current working build in the project documentation.

### Overall Updates
- rewrote README, CHANGELOG, and WORKFLOW to reflect the full change set from the original source to the current working build
- clarified that the project had moved away from hardcoded machine-specific paths toward runtime file and folder selection
- documented the current Script 05 `.xlsb` limitation for J.1 source files

## v00.04.002

### Intention
Make the documentation clearer for setup, run flow, and current project state.

### Overall Updates
- improved README setup and run instructions
- tightened WORKFLOW documentation
- adjusted CHANGELOG wording to read more cleanly as a cumulative update record

## v00.04.001

### Intention
Add a setup and run guide to the project documentation.

### Overall Updates
- added setup and run instructions to README
- documented runtime prompts and outputs for each step

## v00.04

### Intention
Decouple the workflow from hardcoded machine-specific directories and workbook names.

### Overall Updates
- added `workflow_io.py`
- removed the original hardcoded `D:\...` path dependency from Scripts 01-06
- replaced fixed input/output workbook locations with runtime file and folder selection
- added runtime prompting for option period and billing month where applicable
- added lean project documentation

### Per-File Updates
- `01_final_coversheet_generation_script.py`
  - replaced hardcoded J.1, CLIN, PR, and output paths with runtime selection
- `02_overview_file_script.py`
  - replaced hardcoded coversheet, CLIN, and output paths with runtime selection
- `03_F_and_R_script.py`
  - replaced hardcoded overview, PR, and output paths with runtime selection
- `04_build_file_script.py`
  - replaced hardcoded build template, coversheet, PR, and output paths with runtime selection
- `05_J1_script.py`
  - replaced hardcoded build, J.1, and output paths with runtime selection
- `06_J17_file_script.py`
  - replaced hardcoded J.17, J.1, billing, and output paths with runtime selection
- `workflow_io.py`
  - added shared file, folder, save-file, integer, and text prompts
