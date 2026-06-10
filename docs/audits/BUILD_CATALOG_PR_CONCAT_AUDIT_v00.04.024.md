# Build Catalog PR Concatenation Audit — v00.04.024

## Scope

This audit covers the Catalog-generation deduplication rule in `04_build_file_script.py`.

This package supersedes the incorrect `v00.04.023` approach that added `Price Request Number` to the Catalog lookup key.

## Correct business rule

Catalog deduplication should identify duplicate Catalog items by the Catalog item identity, not by the Price Request Number.

When the same Catalog item appears in more than one PR:

- keep one Catalog row;
- remove the duplicate row or rows;
- concatenate the unique Price Request Numbers into the kept row using `/`.

Example:

```text
59984 + 59993 -> 59984/59993
```

## What changed

The recognized pricing-method lookup keys do **not** include `Price Request Number`.

The fallback lookup key for blank or unmapped pricing methods also no longer includes `Price Request Number`.

This means duplicate Catalog items can collapse across PRs while preserving their PR history in the `Price Request Number` field.

## Lookup-key behavior

Recognized pricing methods continue to use their business-specific keys:

| Pricing method | Catalog duplicate key |
|---|---|
| ICB | EIS CLIN / Case Number / SRE Pricing Element / TO Period |
| ORIG NSC | EIS CLIN / Orig NSC / SRE Pricing Element / TO Period |
| TERM NSC | EIS CLIN / Term NSC / SRE Pricing Element / TO Period |
| ORIG NSC-TERM NSC | EIS CLIN / Orig NSC / Term NSC / SRE Pricing Element / TO Period |
| ORIG JUR | EIS CLIN / Orig CJID / SRE Pricing Element / TO Period |
| TERM JUR | EIS CLIN / Term CJID / SRE Pricing Element / TO Period |
| ORIG JUR-TERM JUR | EIS CLIN / Orig CJID / Term CJID / SRE Pricing Element / TO Period |

For blank or unmapped pricing methods, the fallback key uses a conservative full-row identity, excluding `Price Request Number` so duplicate items across PRs can still be collapsed and concatenated.

## Validation performed

- Python compilation passed for `04_build_file_script.py`.
- Targeted tests confirmed that duplicate ICB rows across PRs collapse into one Catalog row with `Price Request Number = 59984/59993`.
- Targeted tests confirmed that fallback blank/unmapped rows also collapse across PRs without using `Price Request Number` as part of the key.
- The uploaded sample Build workbook was regenerated. It produced 333 Catalog rows from 382 J.1 Automated rows, with 49 duplicate elements collapsed and 49 Price Request Number concatenations.
