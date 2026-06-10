# Changelog — v00.04.024

## v00.04.024 — Build Catalog PR concatenation fix

### Changed

- Supersedes the incorrect `v00.04.023` Catalog identity change.
- Removed `Price Request Number` from the fallback blank/unmapped Catalog deduplication key.
- Confirmed that recognized pricing-method Catalog lookup keys do not include `Price Request Number`.
- Preserved the intended duplicate handling: duplicate Catalog items collapse into one row, and unique Price Request Numbers are concatenated with `/`.

### Corrected behavior

If the same Catalog item appears in PR59984 and PR59993, the Catalog should keep one item row and show:

```text
59984/59993
```

in the `Price Request Number` column.

### Validation

- `04_build_file_script.py` compiles successfully.
- Targeted duplicate tests passed for recognized pricing methods and blank/unmapped pricing methods.
- Uploaded sample Build workbook regenerated with PR concatenation behavior preserved.
