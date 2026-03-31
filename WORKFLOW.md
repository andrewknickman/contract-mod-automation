# Workflow Layout

## Current Workflow Structure
The project currently runs as a six-step workflow. The new UI presents these steps in one guided desktop interface while using the existing processing logic underneath.

## Step Order
1. Coversheets
2. Overview Workbook
3. F&R Workbook
4. Build Workbook
5. Update J.1
6. Update J.17

## Step Inputs and Outputs
### Step 1 - Coversheets
Inputs:
- source J.1 previous workbook
- CLIN lookup workbook
- PR workbook folder
- coversheet output folder
- current option period

Output:
- coversheet workbooks

### Step 2 - Overview Workbook
Inputs:
- coversheet workbook folder
- CLIN lookup workbook
- overview output workbook path

Output:
- overview workbook

### Step 3 - F&R Workbook
Inputs:
- overview workbook
- PR workbook folder
- F&R output workbook path

Output:
- F&R workbook

### Step 4 - Build Workbook
Inputs:
- build template workbook
- coversheet workbook folder
- PR workbook folder
- build output workbook path

Output:
- build workbook

### Step 5 - Update J.1
Inputs:
- build workbook
- source J.1 previous workbook
- updated J.1 output workbook path

Output:
- updated J.1 workbook

Run note:
- The source J.1 workbook should be `.xlsx` or `.xlsm` for the current update workflow. `.xlsb` is not safely supported for updating in Step 5.

### Step 6 - Update J.17
Inputs:
- source J.17 workbook
- updated J.1 workbook
- billing workbook
- updated J.17 output workbook path
- current option period
- billing month

Output:
- updated J.17 workbook

## UI Layout
The current UI includes:
- workflow overview page
- one page per workflow step
- run history page
- settings page
- help / run guide page

## State and Reuse
The UI stores:
- saved field values by step
- step status
- run history
- latest configured outputs

The UI can reuse prior outputs for downstream inputs where the workflow naturally connects.
