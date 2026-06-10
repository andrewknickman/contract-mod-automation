# Overview F&R Blank Status Audit - v00.04.014

## Issue

The Overview generation step defaulted `F&R Needed` to `No` whenever no in-scope CLIN row had `F&R Status = Pending` or `F&R Status = Approved`.

That produced false negatives when the coversheet had CLIN rows but the `F&R Status` cells were still empty. Blank status means the coversheet has not provided enough information to determine whether F&R is needed; it should not be treated as a confirmed `No`.

## Corrected behavior

`02_overview_file_script.py` now distinguishes between:

1. **Blank status set**: all in-scope F&R Status cells are empty.
2. **Nonblank status set with no F&R statuses**: values exist, but none are `Pending` or `Approved`, such as `N/A`.
3. **F&R status set**: at least one in-scope row is `Pending` or `Approved`.

## Current rule

| In-scope F&R Status values | F&R Needed | F&R count |
| --- | --- | ---: |
| All blank | Blank | 0 |
| Pending | Yes | Counted |
| Approved | Yes | Counted |
| N/A / NA only | No | 0 |
| Blank plus N/A / NA only | No | 0 |

## Row scope

The script continues to use the same row scope from the prior fixes:

- If one or more CLIN rows are marked `OpDiv Approval = Yes`, only those approved rows are evaluated.
- If no rows are marked with OpDiv approval, all listed CLIN rows are evaluated.

## Validation

Targeted in-memory workbook tests confirmed:

- All blank F&R Status rows produce blank `F&R Needed` and count `0`.
- `Pending` produces `F&R Needed = Yes` and increments the count.
- `Approved` produces `F&R Needed = Yes` and increments the count.
- `N/A` alone produces `F&R Needed = No` and count `0`.
- When OpDiv-approved rows exist, the blank-status rule applies only to the approved row scope.
