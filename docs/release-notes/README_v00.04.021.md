# Mod Automation Project — v00.04.020

This package carries forward the controller UI, flexible file selection, Overview/F&R fixes, F&R script source/J.1 fixes, and the Build schema fixes from prior `v00.04.xxx` iterations.

## New in v00.04.020

`04_build_file_script.py` now handles coversheets where OpDiv approval decisions have not yet been entered.

The Build inclusion rule is now:

| Coversheet OpDiv Approval state | Build behavior |
| --- | --- |
| Explicit `Yes`/`No` decisions exist | Include only rows marked `Yes`. |
| All approval cells blank | Include all valid CLIN rows as pending-review build candidates. |
| Approval column missing | Include all valid CLIN rows and log a warning. |
| Unknown values only, such as `?`, with no explicit `Yes`/`No` | Include all valid CLIN rows and log a warning. |
| Mixed `Yes`/`No`/blank/unknown | Include only rows marked `Yes`. |

A valid Build candidate row must have at least a `CLIN` and `TO Period` so the script can attempt to match it to the PR/J.1 file.

## Important count note

The Overview workbook’s `# of CLINs` is the count for the current option period. Build output can be larger when all option periods are included through the coversheet/PR/J.1 matching process.

## Run the UI

```bash
python mod_automation_ui.py
```

Or on Windows, double-click:

```text
RUN_UI.bat
```

## Run the Build step directly

```bash
python 04_build_file_script.py ^
  --build-file "path\to\Build Template.xlsx" ^
  --coversheets-dir "path\to\coversheets" ^
  --pr-dir "path\to\PR files" ^
  --output-file "path\to\build_file.xlsx"
```

## Validation

Validated with Python compilation and targeted DataFrame tests for the new OpDiv approval fallback behavior. Full workflow validation still requires the complete local source folders and workbooks.

---

## v00.04.021 note — Build pricing-factor method correction

`04_build_file_script.py` now uses the correct business-facing pricing methods for Build pricing-factor matching:

- `ICB`
- `ORIG NSC`
- `TERM NSC`
- `ORIG JUR`
- `TERM JUR`
- `ORIG JUR-TERM JUR`
- `ORIG NSC-TERM NSC`

`ORIG JUR` and `TERM JUR` still read the PR/J.1 columns named `Orig CJID` and `Term CJID`; those are source data fields, not pricing-method names.

The matching logic now also normalizes hyphen spacing and numeric pricing-factor values, so values such as `120036`, `120036.0`, and `'120036'` compare correctly.
