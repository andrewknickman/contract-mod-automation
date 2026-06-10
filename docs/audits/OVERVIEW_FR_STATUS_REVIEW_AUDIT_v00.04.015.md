# Overview F&R Status Review Audit - v00.04.015

## Issue

The Overview generation step correctly stopped treating all-blank `F&R Status` cells as `No` in `v00.04.014`, but it still treated unexpected nonblank values such as `?` as `No`.

That created another false-negative path: a coversheet could contain suspicious or invalid text in the `F&R Status` column, and the Overview would still imply the item had been reviewed and did not need F&R.

## Corrected behavior

`02_overview_file_script.py` now separates F&R status values into three groups:

1. **Counted F&R statuses**: `Pending` and `Approved`.
2. **Recognized non-F&R statuses**: `N/A`, `NA`, `N.A.`, and `Not Applicable`.
3. **Manual-review statuses**: blanks and unexpected values such as `?`.

If any in-scope row has `Pending` or `Approved`, the Overview shows `F&R Needed = Yes` and counts only those `Pending` / `Approved` rows.

If no in-scope row has `Pending` or `Approved`, and at least one in-scope row is blank or unknown, the Overview leaves both `F&R Needed` and the adjacent F&R `# of CLINs` value blank.

If all in-scope rows are recognized non-F&R statuses such as `N/A`, the Overview still shows `F&R Needed = No` and F&R count `0`.

## Current rule

| In-scope F&R Status values | F&R Needed | F&R count column |
| --- | --- | --- |
| Pending | Yes | Counted |
| Approved | Yes | Counted |
| Pending/Approved mixed with unknown values | Yes | Count only Pending/Approved |
| N/A / NA only | No | 0 |
| All blank | Blank | Blank |
| Unknown values only, such as `?` | Blank | Blank |
| Blank/unknown mixed with N/A, but no Pending/Approved | Blank | Blank |

## Row scope

The script continues to use the same row scope from the prior fixes:

- If one or more CLIN rows are marked `OpDiv Approval = Yes`, only those approved rows are evaluated.
- If no rows are marked with OpDiv approval, all listed CLIN rows are evaluated.

## Validation

Targeted in-memory workbook tests confirmed:

- `Pending` produces `F&R Needed = Yes` and increments the F&R count.
- `Approved` produces `F&R Needed = Yes` and increments the F&R count.
- `N/A` alone produces `F&R Needed = No` and count `0`.
- All blank values produce blank `F&R Needed` and blank F&R count.
- Unknown values such as `?` produce blank `F&R Needed` and blank F&R count when no Pending/Approved rows are in scope.
- Pending/Approved rows still produce `Yes` and are counted even if another in-scope row has an unknown value.
