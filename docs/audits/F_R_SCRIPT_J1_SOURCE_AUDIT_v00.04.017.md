# v00.04.017 F&R Script J.1 / Source Audit

## Issues addressed in v00.04.017

1. **COMP source/header variants were too narrow**
   - Some COMP sheets use `Source or Network Information` instead of the canonical generated output header `Source or Networx Information`.
   - The script needed to accept the observed header variant while still writing the canonical output field.

2. **COMP pricing-element headers were too narrow**
   - Some COMP sheets use `Element Number` instead of `SRE Pricing Element`.
   - The script needed to read the pricing element from accepted header variants before attempting J.1 price lookup.

3. **Blank COMP pricing element needed a safer tie-breaker**
   - Some COMP rows have a blank pricing element while multiple J.1 rows share the same case number and option period.
   - In those cases, selecting the first matching row can return the wrong J.1 rate.
   - The script now uses the COMP case description as a fallback tie-breaker against J.1 `ICB Case Description` when the COMP pricing element is blank.

## Scope preserved

- J.1 headers are still expected on Row 1 only.
- No Row 1 header search expansion was added.
- No PR-specific case numbers were intended to be coded as exceptions.
- No source cell values were intended to be coded as exceptions.

## Field behavior after v00.04.017

### Comment field

The generated output field remains:

`Verizon's Response/HHS Comment`

The script accepts multiple source headers such as Verizon response and HHS comment variants, including common punctuation and spelling differences.

### Source field

The generated output field remains:

`Source or Networx Information`

v00.04.017 accepted the observed `Source or Network Information` header variant. v00.04.018 supersedes this with broader source/network/networx header detection.

### Case description field

The generated output field remains:

`Verizon Case Description`

The script accepts source headers such as:

- `Verizon Case Description`
- `Case Description`
- `ICB Case Description`
- `Description`

### Pricing element field

The generated output field remains:

`Pricing Element`

The script accepts source headers such as:

- `SRE Pricing Element`
- `Pricing Element`
- `Element Number`
- `Element #`
- `Element No`
- `Element`

## J.1 matching rule

For each COMP row, the script attempts to locate the corresponding J.1 rate by matching:

1. Case number
2. Current option period
3. Pricing element, when present
4. Case description, only as a fallback when pricing element is blank and multiple same-case/same-period J.1 candidates exist

This is a general matching rule and should not depend on specific PR numbers, case numbers, or source cell text.
