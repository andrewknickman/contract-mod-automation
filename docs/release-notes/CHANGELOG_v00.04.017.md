# Changelog — v00.04.017

## v00.04.017 — F&R J.1/source extraction fix

### Fixed
- Expanded COMP-sheet source-field detection so `Source or Network Information` is accepted in addition to `Source or Networx Information`.
- Expanded COMP-sheet case-description detection so `Case Description` can populate the generated `Verizon Case Description` field.
- Expanded COMP-sheet pricing-element detection so `Element Number ` is accepted as the pricing element source.
- Improved J.1 price lookup for blank-element COMP rows by using COMP case description as a tie-breaker against Row 1 J.1 field `ICB Case Description`.
- Avoided unsafe J.1 first-match selection when multiple same-case/same-period candidates remain ambiguous and have different prices.

### Preserved
- J.1 headers are still expected on Row 1 only.
- Overview F&R business-rule logic from v00.04.015 remains unchanged.
- Generated output column names remain unchanged.
