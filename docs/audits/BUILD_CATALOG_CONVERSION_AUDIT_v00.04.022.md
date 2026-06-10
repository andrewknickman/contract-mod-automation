# Build Catalog Conversion Audit - v00.04.022

## Issue

The generated `Catalog` sheet did not appear to convert everything present in `J.1 Automated`. In the provided sample workbook, `J.1 Automated` contained 382 data rows while the generated `Catalog` contained 307 data rows.

## Root cause

The Catalog-generation logic intentionally runs deduplication, but two parts of the implementation were too aggressive:

1. `apply_pricing_method_cleanup()` ran before `deduplicate_catalog()`. That meant fields used by the deduplication key could be blanked before the duplicate lookup was calculated.
2. Rows with blank or unmapped `YY_Price_Method` fell back to a lookup key containing only `EIS CLIN`. That caused unrelated rows across multiple option periods or PRs to collapse into one Catalog row.

In the provided workbook, CLINs `AA00240` and `AA00340` each had 14 rows in `J.1 Automated`, but each collapsed to 1 Catalog row because their pricing method did not map and the fallback lookup was only the CLIN value.

## Correction

`04_build_file_script.py` now:

- Runs `deduplicate_catalog()` before `apply_pricing_method_cleanup()`.
- Keeps recognized pricing-method deduplication rules intact.
- Uses a conservative full-row identity for blank or unmapped pricing methods, including fields such as EIS CLIN, CLIN name, NSC/CJID values, case number, pricing element, TO period, HHS price, and PR number.

## Expected behavior after fix

- Catalog still consolidates true duplicate Catalog elements for recognized pricing methods.
- Catalog no longer collapses blank/unmapped pricing-method rows by EIS CLIN alone.
- Manual-review rows remain visible instead of being silently removed.

## Validation against uploaded sample

- Before fix: `J.1 Automated` = 382 rows; `Catalog` = 307 rows.
- After fix: `J.1 Automated` = 382 rows; regenerated `Catalog` = 333 rows.
- `AA00240`: 14 J.1 rows -> 14 Catalog rows.
- `AA00340`: 14 J.1 rows -> 14 Catalog rows.
- Recognized duplicates between PR59984 and PR59993 are still consolidated using existing business-specific keys and Price Request Number concatenation.
