# Build Script Pricing-Factor Audit — v00.04.021

## Issue

PR60244 was omitted from the generated Build output even though the coversheet rows were valid and OpDiv-approved.

The failure was caused by the Build script's pricing-factor helper using PR/J.1 field-style labels as pricing-method names:

- `ORIG CJID`
- `TERM CJID`
- `ORIG CJID-TERM CJID`

Those are not the intended CLIN Table pricing methods. The business-facing pricing methods that should control pricing-factor matching are:

- `ICB`
- `ORIG NSC`
- `TERM NSC`
- `ORIG JUR`
- `TERM JUR`
- `ORIG JUR-TERM JUR`
- `ORIG NSC-TERM NSC`

The PR/J.1 source columns for JUR matching are still named `Orig CJID` and `Term CJID`; the mistake was treating `ORIG CJID` / `TERM CJID` as pricing methods.

## Correction

Updated `04_build_file_script.py` so pricing-factor matching now maps pricing methods to PR/J.1 fields as follows:

| Pricing Method | PR/J.1 field used for pricing-factor match |
|---|---|
| `ICB` | `Case Number` |
| `ORIG NSC` | `Orig NSC` |
| `TERM NSC` | `Term NSC` |
| `ORIG JUR` | `Orig CJID` |
| `TERM JUR` | `Term CJID` |
| `ORIG JUR-TERM JUR` | `Orig CJID-Term CJID` |
| `ORIG NSC-TERM NSC` | `Orig NSC-Term NSC` |

Also added normalization so these labels match whether the CLIN Table uses `ORIG JUR-TERM JUR` or `ORIG JUR - TERM JUR`.

## Numeric normalization

Pricing-factor values are normalized before comparison so equivalent values match correctly:

- `120036`
- `120036.0`
- `'120036'`

all compare as the same pricing factor.

## PR60244 validation

Using the attached PR60244 coversheet and PR/J.1 file, the corrected pricing-factor logic matched PR60244 rows for CLINs:

- `VS11235`
- `VS11240`
- `VS12235`
- `VS12240`

The targeted validation matched 28 PR/J.1 rows across OP 5 through OP 11. This aligns with the current-OP Overview count of 4 while allowing the Build to include the larger all-OP row set.

## Preserved behavior

This change does not alter:

- Overview logic
- F&R workbook logic
- OpDiv approval fallback logic from `v00.04.020`
- PR/J.1 Row 1 header expectation
- Catalog output column names such as `Orig CJID` and `Term CJID`
