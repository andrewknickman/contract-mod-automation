from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional


@dataclass
class FieldDef:
    key: str
    label: str
    kind: str  # file, directory, save_file, integer, text
    required: bool = True
    file_filter: str = "All Files (*.*)"
    default_name: Optional[str] = None
    placeholder: str = ""
    help_text: str = ""
    is_output: bool = False
    source_step_id: Optional[str] = None
    source_field_key: Optional[str] = None
    validator: Optional[Callable[[str], Optional[str]]] = None


@dataclass
class StepDef:
    step_id: str
    title: str
    short_title: str
    script_name: str
    entry_mode: str  # main or configure_then_call
    entry_function: str = "main"
    description: str = ""
    fields: List[FieldDef] = field(default_factory=list)
    outputs: List[str] = field(default_factory=list)
    prerequisites: List[str] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)

    @property
    def script_path(self) -> Path:
        return Path(__file__).with_name(self.script_name)


EXCEL_FILES = "Excel Files (*.xlsx *.xlsm *.xlsb *.xls)"
OPENPYXL_J1_FILES = "Excel Files (*.xlsx *.xlsm)"


def _warn_if_not_openpyxl(path_str: str) -> Optional[str]:
    if not path_str:
        return None
    suffix = Path(path_str).suffix.lower()
    if suffix not in {".xlsx", ".xlsm"}:
        return "Script 05 updates the workbook with openpyxl. Use a true .xlsx or .xlsm source file."
    return None


STEPS: List[StepDef] = [
    StepDef(
        step_id="step1",
        title="Step 1 - Generate Coversheets",
        short_title="Coversheets",
        script_name="01_final_coversheet_generation_script.py",
        entry_mode="main",
        description="Generate coversheet workbooks from the prior J.1 workbook, CLIN lookup workbook, and PR folder.",
        fields=[
            FieldDef("j1_previous_workbook", "Source J.1 Previous Workbook", "file", file_filter=EXCEL_FILES),
            FieldDef("clin_lookup_workbook", "CLIN Lookup Workbook", "file", file_filter=EXCEL_FILES),
            FieldDef("pr_folder", "PR Workbook Folder", "directory"),
            FieldDef("coversheet_output_folder", "Coversheet Output Folder", "directory", is_output=True),
            FieldDef("current_option_period", "Current Option Period", "integer", placeholder="5"),
        ],
        outputs=["coversheet_output_folder"],
    ),
    StepDef(
        step_id="step2",
        title="Step 2 - Generate Overview Workbook",
        short_title="Overview Workbook",
        script_name="02_overview_file_script.py",
        entry_mode="main",
        description="Build the overview workbook from generated coversheets and the CLIN lookup workbook.",
        prerequisites=["step1"],
        fields=[
            FieldDef("coversheet_folder", "Coversheet Workbook Folder", "directory", source_step_id="step1", source_field_key="coversheet_output_folder"),
            FieldDef("clin_lookup_workbook", "CLIN Lookup Workbook", "file", file_filter=EXCEL_FILES),
            FieldDef("overview_output_file", "Overview Output Workbook", "save_file", file_filter="Excel Workbook (*.xlsx)", default_name="overview_file.xlsx", is_output=True),
        ],
        outputs=["overview_output_file"],
    ),
    StepDef(
        step_id="step3",
        title="Step 3 - Generate F&R Workbook",
        short_title="F&R Workbook",
        script_name="03_F_and_R_script.py",
        entry_mode="configure_then_call",
        entry_function="build_FR",
        description="Generate the F&R workbook from the overview workbook and PR folder.",
        prerequisites=["step2"],
        fields=[
            FieldDef("overview_workbook", "Overview Workbook", "file", file_filter=EXCEL_FILES, source_step_id="step2", source_field_key="overview_output_file"),
            FieldDef("pr_folder", "PR Workbook Folder", "directory", source_step_id="step1", source_field_key="pr_folder"),
            FieldDef("fr_output_file", "F&R Output Workbook", "save_file", file_filter="Excel Workbook (*.xlsx)", default_name="f_r_output.xlsx", is_output=True),
        ],
        outputs=["fr_output_file"],
    ),
    StepDef(
        step_id="step4",
        title="Step 4 - Generate Build Workbook",
        short_title="Build Workbook",
        script_name="04_build_file_script.py",
        entry_mode="main",
        description="Generate the build workbook using approved coversheet rows, the PR folder, and a build template workbook.",
        prerequisites=["step1"],
        fields=[
            FieldDef("build_template_workbook", "Build Template Workbook", "file", file_filter=EXCEL_FILES),
            FieldDef("coversheet_folder", "Coversheet Workbook Folder", "directory", source_step_id="step1", source_field_key="coversheet_output_folder"),
            FieldDef("pr_folder", "PR Workbook Folder", "directory", source_step_id="step1", source_field_key="pr_folder"),
            FieldDef("build_output_file", "Build Output Workbook", "save_file", file_filter="Excel Workbook (*.xlsx)", default_name="build_file.xlsx", is_output=True),
        ],
        outputs=["build_output_file"],
    ),
    StepDef(
        step_id="step5",
        title="Step 5 - Generate Updated J.1 Workbook",
        short_title="Update J.1",
        script_name="05_J1_script.py",
        entry_mode="main",
        description="Generate the updated J.1 workbook from the build workbook and a source J.1 previous workbook.",
        prerequisites=["step4"],
        notes=["Source J.1 workbook should be a true .xlsx or .xlsm file. .xlsb is not safely supported for updating in Script 05."],
        fields=[
            FieldDef("build_workbook", "Build Workbook", "file", file_filter=EXCEL_FILES, source_step_id="step4", source_field_key="build_output_file"),
            FieldDef("source_j1_workbook", "Source J.1 Previous Workbook", "file", file_filter=OPENPYXL_J1_FILES, validator=_warn_if_not_openpyxl),
            FieldDef("updated_j1_output_file", "Updated J.1 Output Workbook", "save_file", file_filter="Excel Workbook (*.xlsx)", default_name="j1_previous_file.xlsx", is_output=True),
        ],
        outputs=["updated_j1_output_file"],
    ),
    StepDef(
        step_id="step6",
        title="Step 6 - Generate Updated J.17 Workbook",
        short_title="Update J.17",
        script_name="06_J17_file_script.py",
        entry_mode="main",
        description="Generate the updated J.17 workbook from the source J.17 workbook, updated J.1 workbook, and billing workbook.",
        prerequisites=["step5"],
        fields=[
            FieldDef("source_j17_workbook", "Source J.17 Workbook", "file", file_filter=EXCEL_FILES),
            FieldDef("updated_j1_workbook", "Updated J.1 Workbook", "file", file_filter=EXCEL_FILES, source_step_id="step5", source_field_key="updated_j1_output_file"),
            FieldDef("billing_workbook", "Billing Workbook", "file", file_filter=EXCEL_FILES),
            FieldDef("updated_j17_output_file", "Updated J.17 Output Workbook", "save_file", file_filter="Excel Workbook (*.xlsx)", default_name="j17_updated_file.xlsx", is_output=True),
            FieldDef("current_option_period", "Current Option Period", "integer", placeholder="5"),
            FieldDef("billing_month", "Billing Month", "text", placeholder="December"),
        ],
        outputs=["updated_j17_output_file"],
    ),
]

STEP_MAP: Dict[str, StepDef] = {step.step_id: step for step in STEPS}
