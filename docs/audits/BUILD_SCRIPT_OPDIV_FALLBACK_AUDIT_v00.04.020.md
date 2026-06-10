# Build Script OpDiv Fallback Audit — v00.04.020

## Issue

`04_build_file_script.py` previously treated Build eligibility as approved-only. Rows were included only when the OpDiv approval cell was explicitly `Yes`. This caused Build false omissions for coversheets where the CLIN rows were present but the OpDiv Approval column had not yet been filled out.

## Corrected business rule

The Build script now distinguishes between explicit approval decisions and pending-review coversheets:

| Coversheet OpDiv Approval state | Build inclusion behavior |
| --- | --- |
| At least one explicit `Yes`/`No` decision exists | Include only rows marked `Yes`. |
| All approval values are blank | Include all valid CLIN rows as pending-review build candidates. |
| Approval column is missing | Include all valid CLIN rows and log a warning. |
| Unknown/nonstandard approval values only, such as `?`, with no `Yes`/`No` anywhere | Include all valid CLIN rows and log a warning. |
| Mixed `Yes`/`No`/blank/unknown | Include only rows marked `Yes`. |

## Valid CLIN row screening

Before applying approval logic, the script screens out rows that do not have enough information to attempt a PR/J.1 match. A valid Build candidate row must have at least:

- `CLIN`
- `TO Period`

Rows missing either value are skipped and logged.

## Current OP versus all-OP counts

The Overview workbook’s `# of CLINs` represents the count for the current option period. The Build script reads the coversheet CLIN table sheets available to it and then matches against PR/J.1 records, so the final Build output may be larger when all option periods are represented.

The expected comparison is:

- current-OP Overview count should align with the current-OP Build candidates for that PR/coversheet, assuming the rows match the PR/J.1;
- all-OP Build output may legitimately be larger.

## Files changed

- `04_build_file_script.py`

## Validation performed

Validated with targeted in-memory DataFrame tests covering:

- all approval cells blank;
- mixed `Yes`/`No`/blank/unknown;
- unknown values only;
- missing OpDiv approval column;
- rows missing CLIN values.

Also validated Python compilation for the updated Build script.
