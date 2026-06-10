# Dependency Installer Audit — v00.04.026

## Reason for change

Step 5 can now convert previous J.1 `.xlsb` / `.xls` workbooks to `.xlsx`, but that conversion requires additional runtime support on Windows. The package also depends on several Python libraries for Excel, Word, and data handling.

## Added files

| File | Purpose |
| --- | --- |
| `requirements.txt` | Canonical dependency list for the package. |
| `install_dependencies.py` | Installs dependencies and verifies imports. |
| `INSTALL_DEPENDENCIES.bat` | Windows double-click launcher. |
| `INSTALL_DEPENDENCIES.command` | macOS/Linux shell launcher. |
| `VERIFY_DEPENDENCIES.bat` | Windows verification-only launcher. |

## Dependency list

| Package | Used for |
| --- | --- |
| `pandas` | Reading/transformation of workbook tables. |
| `numpy` | Data handling support. |
| `openpyxl` | Reading/writing `.xlsx` workbooks. |
| `pyxlsb` | Reading `.xlsb` files through pandas where supported. |
| `xlrd` | Reading legacy `.xls` files where supported. |
| `python-docx` | MFR walkthrough Word document generation. |
| `pywin32` | Windows-only Microsoft Excel automation for `.xlsb`/`.xls` to `.xlsx` conversion in Step 5. |

## Validation performed

- Confirmed all Python files still compile.
- Confirmed installer script syntax compiles.
- Confirmed `requirements.txt` is present in the package root.
- No workflow logic was changed.
