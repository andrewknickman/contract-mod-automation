# Changelog — v00.04.018

## Fixed

### F&R workbook source-information field

- Generalized COMP-sheet source-header detection for the generated `Source or Networx Information` output column.
- Accepted broader header variants, including:
  - `Source Information`
  - `Networx Information`
  - `Network Information`
  - `Source or Networx Information`
  - `Source or Network Information`
  - shortened `Info` variants of the above.
- Combined unique nonblank values when more than one matching source/network/networx information column exists.

## Preserved

- J.1 headers are still read from Row 1 only.
- J.1 price matching still uses normalized case number, pricing element, option period, and case-description fallback when the COMP pricing element is blank.
- No PR-specific case numbers, source cell values, or vendor-quote text were added as hard-coded exceptions.
- Overview logic from `v00.04.015` remains unchanged.

## Validation

- Confirmed `03_F_and_R_script.py` compiles.
- Ran targeted header-detection tests for accepted and rejected source/header variants.
