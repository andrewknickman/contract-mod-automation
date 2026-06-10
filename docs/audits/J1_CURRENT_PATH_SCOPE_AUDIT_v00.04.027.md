# J.1 Current Path Scope Audit - v00.04.027

## Issue

Step 5 failed immediately after argument parsing with:

```text
UnboundLocalError: local variable 'j1_current_file' referenced before assignment
```

## Root cause

`configure_paths()` assigns the selected output path to the module-level `j1_current_file` variable. In `main()`, the script later reassigned `j1_current_file` after preparing/copying/converting the previous J.1 workbook:

```python
j1_current_file = prepare_j1_current_workbook(j1_previous_file_input, j1_current_file)
```

Because Python treats any variable assigned inside a function as local unless declared `global`, the earlier line:

```python
j1_current_file.parent.mkdir(parents=True, exist_ok=True)
```

was trying to read an uninitialized local variable instead of the module-level value set by `configure_paths()`.

## Correction

`main()` now explicitly declares that it is using the module-level output path variable:

```python
def main(argv=None):
    """Main function to execute the J1 automation."""
    global j1_current_file
```

This preserves the `.xlsb` / `.xls` conversion behavior introduced in `v00.04.025` and fixes the startup failure introduced during that change.

## Validation

- `05_J1_script.py` compiles successfully.
- `python 05_J1_script.py --help` runs successfully.
- No Build, Overview, F&R, Catalog, or dependency-installer workflow logic was changed.
