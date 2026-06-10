# Overview F&R Count and F&R Needed Audit

## Corrected rule in v00.04.013

`02_overview_file_script.py` now applies one consistent business rule for both the Overview `F&R Needed` flag and the numeric F&R count.

## Row scope

- If any CLIN rows are marked `OpDiv Approval = Yes`, the script evaluates those approved rows only.
- If no CLIN rows are marked `OpDiv Approval = Yes`, the script falls back to all listed CLIN rows in `CLIN Table (Current OP)`.

## F&R status rule

Only these in-scope F&R statuses trigger `F&R Needed = Yes`:

- `Pending`
- `Approved`

Only these in-scope F&R statuses increment the numeric F&R count:

- `Pending`
- `Approved`

The following statuses do not trigger `F&R Needed = Yes` and do not increment the F&R count:

- `N/A`
- `NA`
- `N.A.`
- `Not Applicable`
- Blank values

## Practical examples

| In-scope F&R statuses | F&R Needed | F&R count |
|---|---:|---:|
| `N/A`, `NA` | `No` | `0` |
| `Pending` | `Yes` | `1` |
| `Approved` | `Yes` | `1` |
| `Pending`, `Approved`, `N/A` | `Yes` | `2` |

## Validation

Targeted in-memory workbook tests confirmed:

1. A coversheet with no OpDiv approval rows and only `N/A` / `NA` F&R statuses produces `F&R Needed = No` and F&R count `0`.
2. A coversheet with no OpDiv approval rows and statuses `Pending`, `Approved`, and `N/A` produces `F&R Needed = Yes` and F&R count `2`.
3. When approved rows exist, non-approved rows are excluded from both the flag and the count.
