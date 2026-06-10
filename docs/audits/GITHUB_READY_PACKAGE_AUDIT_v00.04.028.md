# GitHub-ready package audit — v00.04.028

## Purpose

Prepare the v00.04.027 package for replacing an existing GitHub repository.

## Changes made for repository readiness

- Reorganized historical release notes, audit files, and diffs under `docs/`.
- Added repository-level `README.md` focused on setup, workflow use, and current business rules.
- Added repository-level `CHANGELOG.md` summarizing the current package lineage.
- Added `.gitignore` to prevent generated workbooks, run folders, Python caches, virtual environments, and local config from being committed.
- Added `.gitattributes` for line-ending normalization and binary workbook handling.
- Added `GITHUB_UPLOAD_GUIDE.md` with GitHub Desktop and command-line replacement instructions.

## Installer cleanup

The following files had an accidental leading backslash before their first line in the prior package. This package removes that character:

- `install_dependencies.py`
- `INSTALL_DEPENDENCIES.bat`
- `INSTALL_DEPENDENCIES.command`
- `VERIFY_DEPENDENCIES.bat`

## Preserved

No workflow business logic changed from v00.04.027.

Preserved scripts:

- `01_final_coversheet_generation_script.py`
- `02_overview_file_script.py`
- `03_F_and_R_script.py`
- `04_build_file_script.py`
- `05_J1_script.py`
- `06_J17_file_script.py`
- `07_MFR_walkthrough_script.py`
- `mod_automation_ui.py`
- `file_selection.py`
- `install_dependencies.py`

## Validation

- Python compilation passed for all project `.py` files.
- `__pycache__` files were removed from the final package.
- No generated Excel/Word/PDF output files are included.
