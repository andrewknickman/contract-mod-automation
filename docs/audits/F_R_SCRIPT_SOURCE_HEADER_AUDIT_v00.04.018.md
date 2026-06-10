# v00.04.018 F&R Source Header Audit

## Issue addressed

The F&R workbook script needed broader handling for the COMP-sheet field that feeds the generated output column:

`Source or Networx Information`

The prior fix accepted the observed `Source or Network Information` header, but the implementation needed to be generalized so it accepts the family of source/network/networx information headers instead of relying on one narrow alias.

## Corrected behavior

`03_F_and_R_script.py` now detects source-information headers by normalized header meaning, not by cell value and not by PR-specific data.

Accepted header examples include:

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

The generated workbook still uses the canonical output header:

`Source or Networx Information`

## Multiple matching columns

If a COMP sheet contains more than one matching source/network/networx information column, the script combines the unique nonblank values into the generated output cell, separated by ` / `.

This mirrors the approach used for separate Verizon response and HHS comment fields.

## What was not changed

- No PR-specific case numbers were added.
- No source cell values were hard-coded.
- No logic was added for literal values such as vendor-quote text.
- J.1 headers are still expected on Row 1 only.
- J.1 price matching still uses normalized value matching across case number, pricing element, option period, and case description fallback when the COMP pricing element is blank.

## Validation performed

Targeted validation confirmed that the source-header detector accepts the intended header variants and rejects unrelated headers such as:

- `Verizon Response`
- `HHS Comment`
- `Case Number`
- `Information`

Python compilation was also run against `03_F_and_R_script.py`.
