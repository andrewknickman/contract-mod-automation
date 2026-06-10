# Mod Automation — v00.04.024

This package continues the `v00.04.xxx` line and supersedes the incorrect `v00.04.023` Catalog deduplication change.

## Main fix in this package

`04_build_file_script.py` now preserves the intended Catalog duplicate behavior:

- `Price Request Number` is **not** part of the Catalog duplicate lookup key.
- Duplicate Catalog items across multiple PRs collapse into one row.
- The kept row concatenates the unique PR numbers with `/`.

Example:

```text
59984/59993
```

## Catalog deduplication rule

Recognized pricing methods use business-specific duplicate keys, such as:

- `ICB`: `EIS CLIN / Case Number / SRE Pricing Element / TO Period`
- `ORIG NSC`: `EIS CLIN / Orig NSC / SRE Pricing Element / TO Period`
- `TERM NSC`: `EIS CLIN / Term NSC / SRE Pricing Element / TO Period`
- `ORIG JUR`: `EIS CLIN / Orig CJID / SRE Pricing Element / TO Period`
- `TERM JUR`: `EIS CLIN / Term CJID / SRE Pricing Element / TO Period`

The fallback key for blank or unmapped pricing methods uses a conservative full-row identity, but still excludes `Price Request Number` so duplicate Catalog items from different PRs can collapse correctly.

## Running the UI

```bash
python mod_automation_ui.py
```

On Windows, you can also run:

```text
RUN_UI.bat
```

## Notes

- `J.1 Automated` remains the expanded working set.
- `Catalog` is allowed to have fewer rows than `J.1 Automated` when true duplicate Catalog items are collapsed.
- When duplicates are collapsed, the Catalog should preserve the PR history by concatenating the Price Request Numbers.
