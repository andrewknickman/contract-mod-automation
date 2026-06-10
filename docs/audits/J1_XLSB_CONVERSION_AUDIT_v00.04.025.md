# J.1 `.xlsb` Conversion Audit — v00.04.025

## Reported failure

Step 5 failed while appending Catalog rows into the current J.1 workbook:

```text
OSError: File contains no valid workbook part
```

The run showed this sequence:

```text
--j1-previous-file "...J.1_Pricing_Spreadsheet_P00079.xlsb"
--output-file "...j1_current_file.xlsx"
Copied J1 previous file to: ...j1_current_file.xlsx
...
wb = load_workbook(j1_file_path)
OSError: File contains no valid workbook part
```

## Root cause

The previous J.1 workbook was an `.xlsb` binary workbook. The script copied that binary file directly to a path ending in `.xlsx`.

That changed the filename extension but did not change the actual workbook format. `openpyxl` then attempted to open the file as an Open XML `.xlsx` workbook and failed because the file did not contain a valid `.xlsx` workbook package.

## Corrected rule

The J.1 output workbook must be a real `.xlsx` workbook before `openpyxl` appends rows.

| Previous J.1 input type | New behavior |
|---|---|
| `.xlsx` | Copy to output and validate |
| `.xlsm` | Copy to output and validate |
| `.xlsb` | Convert to `.xlsx` using Excel COM, then validate |
| `.xls` | Convert to `.xlsx` using Excel COM, then validate |

## Implementation

`05_J1_script.py` now includes:

- `_force_xlsx_output_path(...)`
- `_convert_with_excel_com(...)`
- `prepare_j1_current_workbook(...)`

Step 5 now calls `prepare_j1_current_workbook(...)` instead of directly calling `shutil.copy2(...)`.

## Operational note

`.xlsb` / `.xls` conversion requires Microsoft Excel on Windows and `pywin32`:

```bash
pip install pywin32
```

If Excel COM is not available, the script now fails early with a clear action:

1. Open the previous J.1 workbook in Excel.
2. Save As `.xlsx`.
3. Rerun Step 5 using the saved `.xlsx` file.

## Scope

This fix only addresses Step 5 J.1 preparation and workbook format handling. It does not alter Build/Catalog generation, Catalog deduplication, PR concatenation, or Catalog-to-J.1 field mapping.
