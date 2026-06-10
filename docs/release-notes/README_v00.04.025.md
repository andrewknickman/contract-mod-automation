# Mod Automation — v00.04.025

This package continues the `v00.04.xxx` line and supersedes `v00.04.024` for Step 5 / J.1 workbook generation.

## Main fix in this package

`05_J1_script.py` now handles a previous J.1 workbook selected as `.xlsb` or `.xls` more safely.

The prior behavior copied the selected previous J.1 file directly to the requested output path. If the selected previous J.1 file was `.xlsb` and the output path was `.xlsx`, the script produced an `.xlsx`-named file that still contained binary `.xlsb` content. The next `openpyxl.load_workbook(...)` call then failed with:

```text
OSError: File contains no valid workbook part
```

## Corrected behavior

- `.xlsx` / `.xlsm` previous J.1 files are copied to the current J.1 output workbook.
- `.xlsb` / `.xls` previous J.1 files are converted to `.xlsx` before the script appends rows.
- The script validates that the prepared output is a real `.xlsx` workbook before continuing.
- The output path is forced to `.xlsx` if a different extension is supplied.

## Important note for `.xlsb` inputs

Python `openpyxl` cannot directly edit `.xlsb` workbooks. When the previous J.1 workbook is `.xlsb` or `.xls`, the script attempts conversion through desktop Microsoft Excel using `pywin32`.

On the Windows workstation running Step 5, this requires:

```bash
pip install pywin32
```

and Microsoft Excel must be installed.

If that is not available, manually open the previous J.1 workbook in Excel, use **Save As → .xlsx**, and select that saved `.xlsx` file as the Step 5 previous J.1 input.

## Running the UI

```bash
python mod_automation_ui.py
```

On Windows, you can also run:

```text
RUN_UI.bat
```

## Scope preserved

This package does not change the Build/Catalog logic from `v00.04.024`.

The Step 5 J.1 logic still reads the Build workbook `Catalog` sheet and appends rows into:

- `2B_Opt Pd <current OP> Catalog`
- `2C_Opt Pd <next OP>-11 Catalog`
