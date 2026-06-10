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
