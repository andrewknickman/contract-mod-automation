# v00.04.019 Build Script Coversheet Schema Audit

## Issue found

The Build step was run with the same mixed folder selected for both `--coversheets-dir` and `--pr-dir`. That folder contains both TO Mod Coversheet workbooks and PR/J.1 workbooks.

The previous Build script treated every Excel workbook in the selected coversheet folder as a coversheet. As a result, it attempted to read PR/J.1 files as coversheets and produced repeated messages like:

```text
Worksheet named 'CLIN Table (Current OP)' not found
```

The run then failed on a real coversheet whose OpDiv approval header did not exactly match the hard-coded string:

```text
KeyError: 'OpDiv Approval (Yes/ No)'
```

## Root cause

`04_build_file_script.py` had two brittle assumptions:

1. Every Excel file in the selected coversheet folder was a coversheet.
2. The OpDiv approval column was always named exactly `OpDiv Approval (Yes/ No)`.

## Fix implemented

### 1. Skip non-coversheet Excel workbooks

The Build script now checks workbook sheet names before adding a file to the coversheet-processing list.

A workbook is treated as a coversheet only if it contains a recognized coversheet CLIN sheet, such as:

- `CLIN Table (Current OP)`
- `CLIN Table Current OP`
- `CLIN Table - Current OP`
- `CLIN Table (OY)`
- `CLIN Table OY`
- `CLIN Table - OY`

This prevents PR/J.1 files from being processed as coversheets when the coversheet folder and PR folder point to the same mixed directory.

### 2. Accept OpDiv approval header variants

The Build script now detects the OpDiv approval column by normalized header matching.

Accepted variants include:

- `OpDiv Approval (Yes/ No)`
- `OpDiv Approval (Yes/No)`
- `OpDiv Approval (Yes / No)`
- `OpDiv Approval Yes No`
- `OpDiv Approval (Y/N)`
- `OpDiv Approval Y/N`
- `OpDiv Approval`
- `OpDiv Approved`

### 3. Fail safely instead of crashing

If a coversheet-like workbook is missing the OpDiv approval column entirely, the Build script now logs a warning and skips that file instead of raising a `KeyError` and stopping the full Build step.

## Preserved behavior

- The Build step still uses OpDiv-approved rows only.
- The Build step still matches coversheet rows against the corresponding PR/J.1 file.
- The F&R script fixes from `v00.04.018` were not changed.
- The Overview script fixes from `v00.04.015` were not changed.

## Validation

- Confirmed `04_build_file_script.py` compiles with `python -m py_compile`.
- Ran targeted tests confirming:
  - `OpDiv Approval (Yes/No)` is accepted.
  - `OpDiv Approval (Y/N)` is accepted.
  - Missing OpDiv approval headers produce a warning instead of a crash.
  - Workbooks with coversheet CLIN sheets are classified as coversheets.
  - Workbooks with only a `J1` sheet are skipped as non-coversheet workbooks.
