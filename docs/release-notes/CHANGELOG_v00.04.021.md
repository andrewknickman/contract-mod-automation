# Changelog — v00.04.020

## v00.04.020 — Build OpDiv approval fallback fix

### Fixed

- Updated `04_build_file_script.py` so blank OpDiv approval status is no longer treated as rejected for Build purposes.
- If a coversheet has explicit `Yes`/`No` approval decisions, the Build still includes only rows marked `Yes`.
- If a coversheet has no explicit `Yes`/`No` decisions, the Build now includes all valid CLIN rows as pending-review build candidates.
- If the OpDiv approval column is missing, the Build now includes all valid CLIN rows and logs a warning instead of omitting the whole coversheet.
- If the OpDiv approval column contains only unknown values such as `?`, the Build now includes all valid CLIN rows and logs a warning.
- Rows missing a usable `CLIN` or `TO Period` are skipped before matching because they cannot be reliably matched to PR/J.1 rows.

### Preserved

- Explicit `No` rows remain excluded when the coversheet contains actual approval decisions.
- Mixed approval sheets still use only rows marked `Yes`.
- Existing PR/J.1 matching logic was not changed.
- Existing Overview and F&R script fixes were not changed.

### Note

- Overview `# of CLINs` represents the current option period count. Build output can be larger when the Build process accounts for all option periods.

---

## v00.04.021 — Build pricing-factor method correction

### Fixed

- Corrected Build pricing-factor matching in `04_build_file_script.py`.
- Replaced the mistaken pricing-method handling for `ORIG CJID`, `TERM CJID`, and `ORIG CJID-TERM CJID` with the intended business-facing pricing methods:
  - `ORIG JUR`
  - `TERM JUR`
  - `ORIG JUR-TERM JUR`
- Preserved use of the PR/J.1 source columns `Orig CJID` and `Term CJID` for JUR-based matching.
- Added support for all intended pricing-factor methods:
  - `ICB`
  - `ORIG NSC`
  - `TERM NSC`
  - `ORIG JUR`
  - `TERM JUR`
  - `ORIG JUR-TERM JUR`
  - `ORIG NSC-TERM NSC`
- Added pricing-method normalization so both `ORIG JUR-TERM JUR` and `ORIG JUR - TERM JUR` are treated the same.
- Added pricing-factor value normalization so `120036`, `120036.0`, and `'120036'` compare as equivalent.

### Validation

- Python compilation passed.
- Targeted function tests passed for all intended pricing methods.
- Targeted PR60244 validation matched 28 rows across OP 5 through OP 11 using the corrected `ORIG JUR` mapping.
