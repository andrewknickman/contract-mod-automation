# WORKFLOW

## Current Workflow Layout
This project currently runs as a six-script pipeline. Each script produces an output workbook used by a later stage.

## Script Order
1. `01_final_coversheet_generation_script.py`
2. `02_overview_file_script.py`
3. `03_F_and_R_script.py`
4. `04_build_file_script.py`
5. `05_J1_script.py`
6. `06_J17_file_script.py`

## Input and Output Flow

### 1. Coversheet Generation
**Script:** `01_final_coversheet_generation_script.py`

**Inputs**
- source J.1 workbook
- CLIN lookup workbook
- PR workbook folder
- current option period

**Output**
- coversheet workbooks

---

### 2. Overview File
**Script:** `02_overview_file_script.py`

**Inputs**
- coversheet workbook folder
- CLIN lookup workbook

**Output**
- overview workbook

---

### 3. F&R File
**Script:** `03_F_and_R_script.py`

**Inputs**
- overview workbook
- PR workbook folder

**Output**
- F&R workbook

---

### 4. Build File
**Script:** `04_build_file_script.py`

**Inputs**
- build template workbook
- coversheet workbook folder
- PR workbook folder

**Output**
- build workbook

---

### 5. Updated J.1 File
**Script:** `05_J1_script.py`

**Inputs**
- build workbook
- source J.1 workbook

**Output**
- updated J.1 workbook

**Run note:** Use a true `.xlsx` or `.xlsm` source J.1 workbook for the current update flow. A `.xlsb` source file is not safely supported for the workbook update step.

---

### 6. Updated J.17 File
**Script:** `06_J17_file_script.py`

**Inputs**
- source J.17 workbook
- updated J.1 workbook
- billing workbook
- current option period
- billing month

**Output**
- updated J.17 workbook
