# Automated Contract Modifications - v00.04.027

This package supersedes `v00.04.026` with a targeted Step 5 J.1 generation fix.

## Main fix

`05_J1_script.py` no longer fails with:

```text
UnboundLocalError: local variable 'j1_current_file' referenced before assignment
```

The error was caused by a Python variable-scope issue introduced while adding `.xlsb` / `.xls` conversion support. The script now correctly uses the selected output path before preparing the current J.1 workbook.

## Step 5 command example

```bash
python 05_J1_script.py \
  --build-file "runs/P00078/04_build/build_file.xlsx" \
  --j1-previous-file "J.1_Pricing_Spreadsheet_P00079.xlsb" \
  --output-file "runs/P00078/05_j1/j1_current_file.xlsx" \
  --current-opt-pd 5
```

## Dependency note

If the previous J.1 workbook is `.xlsb` or `.xls`, the conversion path requires Microsoft Excel on Windows and `pywin32`.

Run the dependency installer first:

```bash
python install_dependencies.py
```

or on Windows double-click:

```text
INSTALL_DEPENDENCIES.bat
```

## Validation

- `05_J1_script.py` compiles.
- `05_J1_script.py --help` runs successfully.
