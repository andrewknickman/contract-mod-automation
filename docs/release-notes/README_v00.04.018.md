# Mod Automation — v00.04.018

This package continues the `v00.04.xxx` controller/UI line and applies a targeted fix to the F&R workbook script.

## Primary fix

`03_F_and_R_script.py` now handles additional COMP-sheet header variants for the generated output column:

`Source or Networx Information`

Accepted input headers include:

- `Source or Networx Information`
- `Source or Network Information`
- `Source Information`
- `Networx Information`
- `Network Information`
- `Source or Networx Info`
- `Source or Network Info`
- `Source Info`
- `Networx Info`
- `Network Info`

The output workbook still uses the canonical header:

`Source or Networx Information`

## J.1 price matching

The J.1 price-matching approach from `v00.04.017` is preserved:

- J.1 headers are expected on Row 1 only.
- Matching uses normalized case number, pricing element, and option period.
- If the COMP pricing element is blank and multiple J.1 rows match the same case/period, the case description is used as a fallback tie-breaker.

## What this version avoids

This version does not hard-code PR-specific examples, case numbers, or source cell values. The source-field fix is based on header detection only.

## Run

Use the controller UI:

```bash
python mod_automation_ui.py
```

Or run the F&R script directly:

```bash
python 03_F_and_R_script.py --overview-file "path/to/overview.xlsx" --pr-dir "path/to/pr_files" --output-file "path/to/f_r_output.xlsx"
```

## Validation

- Python compilation completed for `03_F_and_R_script.py`.
- Targeted tests confirmed accepted source-header variants and rejected unrelated headers.
