# Mod Automation v00.04.019

This package continues the `v00.04.xxx` line and focuses on the Build step.

## Main fix

`04_build_file_script.py` now handles mixed input folders more safely.

Previously, when the same folder was selected for both PR files and coversheets, the Build script attempted to process every Excel file in that folder as a coversheet. PR/J.1 workbooks do not contain the coversheet sheets, so this produced repeated `CLIN Table (Current OP)` errors. The script could then crash when a real coversheet used an OpDiv approval header variant instead of the exact hard-coded header `OpDiv Approval (Yes/ No)`.

## What changed

- PR/J.1 workbooks are skipped when scanning the coversheet folder.
- Workbooks are treated as coversheets only when they contain recognized coversheet CLIN sheets.
- The OpDiv approval column now accepts reasonable header variants.
- Missing OpDiv approval columns now produce a warning instead of a hard crash.

## Accepted coversheet CLIN sheet names

- `CLIN Table (Current OP)`
- `CLIN Table Current OP`
- `CLIN Table - Current OP`
- `CLIN Table (OY)`
- `CLIN Table OY`
- `CLIN Table - OY`

## Accepted OpDiv approval headers

- `OpDiv Approval (Yes/ No)`
- `OpDiv Approval (Yes/No)`
- `OpDiv Approval (Yes / No)`
- `OpDiv Approval Yes No`
- `OpDiv Approval (Y/N)`
- `OpDiv Approval Y/N`
- `OpDiv Approval`
- `OpDiv Approved`

## Run the Build step directly

```bash
python 04_build_file_script.py \
  --build-file "path/to/Build Template.xlsx" \
  --coversheets-dir "path/to/coversheets-or-mixed-folder" \
  --pr-dir "path/to/pr-folder" \
  --output-file "path/to/output/build_file.xlsx"
```

The coversheets folder and PR folder can point to the same mixed folder, but using separate folders is still cleaner when available.

## Validation

- `04_build_file_script.py` compiles successfully.
- Targeted tests confirmed OpDiv approval header variants are recognized.
- Targeted tests confirmed J.1-only workbooks are skipped during coversheet scanning.
