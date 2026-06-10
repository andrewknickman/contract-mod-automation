# F&R Workbook Determination Audit

## How the current script decides what gets an F&R sheet

`03_F_and_R_script.py` does not decide F&R inclusion by re-reading every coversheet. It starts from the generated Overview workbook.

The chain is:

1. `02_overview_file_script.py` reads each coversheet.
2. It writes one row per coversheet/PR into the Overview workbook.
3. `03_F_and_R_script.py` reads that Overview workbook.
4. It selects only rows where `F&R Needed` equals `Yes`, case-insensitive after trimming whitespace.
5. For each selected PR, it looks in the selected PR folder for Excel files whose names start with that PR identifier.
6. For each matching PR file, it looks for a worksheet whose name contains `comp`.
7. It extracts F&R records from that comp worksheet.
8. It creates the main `Overview` sheet in the F&R workbook and then creates PR tabs through `xlwings`.

## Exact current filter

The current F&R selection starts here:

```python
fr_needed = df[df["F&R Needed"].astype(str).str.strip().str.lower() == "yes"]
```

That means the F&R workbook only sees PRs that the Overview workbook already marked as needing F&R.

## Why sheets may be missing

A PR can be skipped for any of these reasons:

- The generated Overview row says `F&R Needed = No` or is blank.
- The Overview workbook is stale and was not regenerated after coversheet corrections.
- The PR number in Overview does not match the beginning of the PR workbook filename.
- The PR workbook is not in the selected PR folder.
- The PR workbook does not contain a sheet with `comp` in the sheet name.
- The comp worksheet is empty or cannot be read by pandas.
- Multiple workbooks/versions share the same PR number; the tab creation logic groups primarily by PR number.

## Related v00.04.010 fix

Before `v00.04.010`, the Overview script set `# of CLINs (Line items in current OP)` from the number of rows marked `OpDiv Approval = Yes`. If no rows were marked Yes, the CLIN count could be `0` even when CLIN rows existed.

The same approval-only dependency could also prevent `F&R Needed` from being set when no approval rows were marked. This could cause the F&R workbook script to skip a PR because the PR never made it into the `F&R Needed = Yes` set.

`v00.04.010` keeps the prior approval-based behavior when approval rows are present, but adds a fallback: when no approval rows are marked, the Overview script counts listed CLIN rows and checks their F&R statuses.
