# Mod Automation — v00.04.017

This version keeps the controller/UI and prior Overview fixes, and updates the F&R workbook generation script.

## Main fix in this version

`03_F_and_R_script.py` now handles additional real PR COMP-sheet header variants and safer J.1 price matching.

### Accepted COMP source/comment fields
The generated output still uses the canonical field name:

`Source or Networx Information`

The script now accepts source data from headers such as:

- `Source or Networx Information`
- `Source or Network Information`
- `Source Information`
- `Networx Information`
- `Network Information`

### Accepted COMP case-description fields
The generated output still uses:

`Verizon Case Description`

The script now accepts source data from headers such as:

- `Verizon Case Description`
- `Case Description`
- `ICB Case Description`
- `Description`

### Accepted COMP pricing-element fields
The script now accepts pricing elements from headers such as:

- `SRE Pricing Element`
- `Pricing Element`
- `Element Number`
- `Element #`
- `Element No`
- `Element`

### J.1 matching rule
J.1 headers are still expected on Row 1 of the PR file.

Required J.1 Row 1 headers remain:

- `Case Number`
- `SRE Pricing Element`
- `TO Period`
- `HHS Price`

If a COMP row includes a pricing element, the J.1 price is matched by case number, pricing element, and option period.

If a COMP row has a blank pricing element, the script matches by case number and option period, then uses the COMP case description to choose the corresponding J.1 `ICB Case Description` when multiple J.1 rows exist for the same case.

If the row is still ambiguous and candidate prices differ, the script leaves `J.1 Rate` blank rather than selecting the wrong row.

## Run

From the project folder:

```bash
python mod_automation_ui.py
```

Or run the F&R step directly:

```bash
python 03_F_and_R_script.py --overview-file path/to/overview.xlsx --pr-dir path/to/pr_files --output-file path/to/f_r_output.xlsx --current-option-period 5
```
