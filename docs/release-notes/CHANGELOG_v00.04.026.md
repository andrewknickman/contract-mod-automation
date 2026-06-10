# Changelog — v00.04.026

## Added

- Added `requirements.txt` listing the runtime dependencies for the automation package.
- Added `install_dependencies.py`, a Python installer/verification script that installs requirements into the active Python environment and checks the imports afterward.
- Added `INSTALL_DEPENDENCIES.bat` for Windows users.
- Added `INSTALL_DEPENDENCIES.command` for macOS/Linux users.
- Added `VERIFY_DEPENDENCIES.bat` for checking the environment without reinstalling packages.

## Notes

- `pywin32` is included with a Windows-only requirement marker because it is only needed when Step 5 converts `.xlsb` or `.xls` previous J.1 workbooks to `.xlsx` through Microsoft Excel COM automation.
- On locked-down work machines, installation may still require use of the approved Python environment or IT approval.

## Preserved

- No workflow logic was changed in this iteration.
- The `v00.04.025` J.1 `.xlsb` conversion behavior is preserved.
