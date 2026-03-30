import pandas as pd
import os
from pathlib import Path

from workflow_io import choose_directory, choose_file, choose_save_file
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl import Workbook, load_workbook
from collections import defaultdict
import math



# ─── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR = None
OUTPUT_DIR = None
PR_DIR = None

overview_file = None
f_r_output_file = None


def get_fr_pr_numbers(overview_path: Path):
    """
    Reads the overview file and returns a list of tuples for rows where 'F&R Needed' == 'Yes'.
    Each tuple contains:
        (PR#, Version, OpDiv, SF30 Description, 12M+ CLIN)
    """
    # Read Excel and strip column names
    df = pd.read_excel(overview_path, dtype=str)  # read all as string to preserve formatting
    df.columns = df.columns.str.strip()

    # Filter rows where F&R Needed == 'Yes'
    fr_needed = df[df["F&R Needed"].astype(str).str.strip().str.lower() == "yes"]

    # Strip whitespace and get required columns
    pr_version_list = list(
        zip(
            fr_needed["PR#"].astype(str).str.strip(),
            fr_needed["Version"].astype(str).str.strip(),
            fr_needed["OpDiv"].astype(str).str.strip() if "OpDiv" in fr_needed.columns else [""] * len(fr_needed),
            fr_needed["SF30 Description"].astype(str).str.strip() if "SF30 Description" in fr_needed.columns else [""] * len(fr_needed),
            fr_needed["12M+ CLIN"].astype(str).str.strip() if "12M+ CLIN" in fr_needed.columns else [""] * len(fr_needed),
        )
    )

    print(f"Found {len(pr_version_list)} PRs needing F&R.")
    return pr_version_list


def extract_comps_data(comps_df):
    """
    Extract all relevant F&R fields from the Comps worksheet in order:
    - Verizon Response
    - Source Info
    - Case Number
    - Verizon Case Description
    - Pricing Element
    - Comp Rate (as float)
    Returns a list of dicts, one per row.
    """
    records = []
    num_rows = len(comps_df)

    for i in range(num_rows):
        verizon_response = comps_df.get("Verizon's Response/HHS Comment", [""] * num_rows)[i] if "Verizon's Response/HHS Comment" in comps_df else ""
        source_info = comps_df.get("Source or Networx Information", [""] * num_rows)[i] if "Source or Networx Information" in comps_df else ""
        case_number = comps_df.get("Case Number", [""] * num_rows)[i] if "Case Number" in comps_df else ""
        verizon_case_desc = comps_df.get("Verizon Case Description", [""] * num_rows)[i] if "Verizon Case Description" in comps_df else ""
        pricing_element = comps_df.get("Pricing Element", [""] * num_rows)[i] if "Pricing Element" in comps_df else ""
        comp_rate_raw = comps_df.get("Comp Rate", [""] * num_rows)[i] if "Comp Rate" in comps_df else ""
        
        # Convert Comp Rate to float
        comp_rate = None
        if comp_rate_raw and comp_rate_raw != "":
            try:
                comp_rate = float(str(comp_rate_raw).replace('$', '').replace(',', ''))
            except (ValueError, TypeError):
                comp_rate = None

        record = {
            "Verizon's Response/HHS Comment": verizon_response,
            "Source or Networx Information": source_info,
            "Case Number": case_number,
            "Verizon Case Description": verizon_case_desc,
            "Pricing Element": pricing_element,
            "Comp Rate": comp_rate,  # Now a float or None
        }

        records.append(record)

    return records


def format_currency(value):
    """Format a numeric value as currency with 6 decimal places."""
    if value is None:
        return ""
    try:
        return f"${float(value):,.6f}"
    except (ValueError, TypeError):
        return ""


def get_comps_sheet(pr_file: Path):
    """Load the 'comps' worksheet (case-insensitive) and return as DataFrame."""
    try:
        xls = pd.ExcelFile(pr_file)
        comps_sheet = next((s for s in xls.sheet_names if "comp" in s.lower()), None)
        if comps_sheet:
            df = pd.read_excel(xls, sheet_name=comps_sheet)
            df.columns = df.columns.str.strip()
            return df
        else:
            print(f"  No 'comps' sheet found in {pr_file.name}")
            return pd.DataFrame()
    except Exception as e:
        print(f"  Error reading comps sheet from {pr_file.name}: {e}")
        return pd.DataFrame()


def determine_12m_clin(verizon_case_desc):
    """
    Determine 12M+ CLIN based on Verizon Case Description.
    Returns 'Yes' if '.ANN.' is found in the description, otherwise 'No'.
    """
    if verizon_case_desc is None:
        return "No"
    
    desc_str = str(verizon_case_desc).strip()
    return "Yes" if ".ANN." in desc_str.upper() else "No"


def get_j1_rate(pr_file: Path, case_number: str, pricing_element: str):
    """
    Extract J1 rate from the J1 worksheet for a given case number and pricing element.
    Returns the HHS Price as a float for the row where TO Period == "OPT PD 5".
    """
    try:
        wb = load_workbook(pr_file)
        j1_sheet = next((s for s in wb.sheetnames if "j1" in s.lower() or "j.1" in s.lower()), None)
        
        if not j1_sheet:
            print(f"  No 'J1' sheet found in {pr_file.name}")
            wb.close()
            return None
        
        ws = wb[j1_sheet]
        
        # Get headers from first row
        headers = {}
        for idx, cell in enumerate(ws[1], start=1):
            if cell.value:
                headers[str(cell.value).strip()] = idx
        
        # Check if required columns exist
        if "Case Number" not in headers or "SRE Pricing Element" not in headers or "TO Period" not in headers or "HHS Price" not in headers:
            print(f"  Missing required columns in J1 sheet of {pr_file.name}")
            wb.close()
            return None
        
        case_col = headers["Case Number"]
        pricing_col = headers["SRE Pricing Element"]
        period_col = headers["TO Period"]
        price_col = headers["HHS Price"]
        
        # Search for matching row
        for row in ws.iter_rows(min_row=2):
            row_case = str(row[case_col - 1].value).strip() if row[case_col - 1].value else ""
            row_pricing = str(row[pricing_col - 1].value).strip() if row[pricing_col - 1].value else ""
            row_period = str(row[period_col - 1].value).strip() if row[period_col - 1].value else ""
            
            # Match case number, pricing element, and TO Period
            if row_case == case_number and row_pricing == pricing_element and row_period == "OPT PD 5":
                hhs_price_cell = row[price_col - 1]
                
                if hhs_price_cell.value is None:
                    wb.close()
                    return None
                
                # Return as float
                if isinstance(hhs_price_cell.value, (int, float)):
                    wb.close()
                    return float(hhs_price_cell.value)
                else:
                    # Try to parse string value
                    try:
                        value_str = str(hhs_price_cell.value).replace('$', '').replace(',', '')
                        wb.close()
                        return float(value_str)
                    except (ValueError, TypeError):
                        wb.close()
                        return None
        
        wb.close()
        return None
        
    except Exception as e:
        print(f"  Error reading J1 sheet from {pr_file.name}: {e}")
        return None



def configure_runtime():
    global BASE_DIR, OUTPUT_DIR, PR_DIR, overview_file, f_r_output_file

    print("Select the files and folders needed for the F&R build...")
    overview_path = choose_file(
        title="Select the overview workbook",
        filetypes=[("Excel Files", "*.xlsx *.xlsm *.xlsb *.xls")],
        state_key="script03_overview_file",
    )
    pr_dir_path = choose_directory(
        title="Select the folder that contains the PR files",
        state_key="script03_pr_dir",
    )
    output_path = choose_save_file(
        title="Choose where to save the F&R workbook",
        default_name="f_r_output.xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        state_key="script03_fr_output_file",
    )

    overview_file = overview_path
    PR_DIR = pr_dir_path
    f_r_output_file = output_path
    OUTPUT_DIR = output_path.parent
    BASE_DIR = pr_dir_path.parent
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def build_FR():
    """
    Builds the F&R Overview Excel file replicating the official example formatting,
    converts Pricing Element to numeric when possible, and merges duplicate Case Numbers.
    Also creates separate PR tabs before merging.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Step 1: Get PRs needing F&R
    fr_pr_list = get_fr_pr_numbers(overview_file)
    fr_overview_records = []

    # Step 2: Gather data
    for pr, version, opdiv, sf30_desc, cl12m in fr_pr_list:
        pr_identifier = pr if "PR" in pr.upper() else f"PR{pr}"
        print(f"Processing {pr_identifier}...")

        pr_files = [
            f for f in PR_DIR.glob(f"{pr_identifier}*.xlsx")
            if not f.name.startswith("~$")
        ]
        if not pr_files:
            print(f"  No PR files found for {pr_identifier} in {PR_DIR}")
            continue

        for pr_file in pr_files:
            comps_df = get_comps_sheet(pr_file)
            if comps_df.empty:
                continue

            comps_records = extract_comps_data(comps_df)

            for rec in comps_records:
                price_elem = rec.get("Pricing Element", "")
                if price_elem is None or (isinstance(price_elem, float) and math.isnan(price_elem)):
                    price_elem_str = ""
                else:
                    try:
                        price_elem_num = int(float(price_elem))
                        price_elem_str = str(price_elem_num).zfill(2)
                    except (ValueError, TypeError):
                        price_elem_str = str(price_elem).zfill(2)

                case_num = rec.get("Case Number", "")
                j1_rate_float = get_j1_rate(pr_file, case_num, price_elem_str)
                comp_rate_float = rec.get("Comp Rate")

                delta_float = None
                if j1_rate_float is not None and comp_rate_float is not None:
                    delta_float = j1_rate_float - comp_rate_float

                j1_rate = format_currency(j1_rate_float)
                comp_rate = format_currency(comp_rate_float)
                delta = format_currency(delta_float)

                fr_overview_records.append({
                    "Type of F&R": "",
                    "Verizon's Response/HHS Comment": rec.get("Verizon's Response/HHS Comment", ""),
                    "J.1 Rate": j1_rate,
                    "Comp Rate": comp_rate,
                    "Delta": delta,
                    "Source or Networx Information": rec.get("Source or Networx Information", ""),
                    "Case Number": rec.get("Case Number", ""),
                    "Verizon Case Description": rec.get("Verizon Case Description", ""),
                    "Pricing Element": price_elem_str,
                    "PR#": pr,
                    "Version": str(version).zfill(2),
                    "OpDiv": opdiv,
                    "SF30 Description": sf30_desc,
                    "12M+ CLIN": determine_12m_clin(rec.get("Verizon Case Description", ""))
                })

    if not fr_overview_records:
        print("No F&R records found.")
        return None

    # Create workbook early (before merging) to build PR tabs
    wb = Workbook()

    # --- Shared Styles ---
    green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    black_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    light_green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    white_font = Font(bold=True, color="FFFFFF")
    black_font = Font(bold=True, color="000000")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000")
    )

    # --- Define columns ---
    columns = [
        "Type of F&R",
        "Verizon's Response/HHS Comment",
        "J.1 Rate",
        "Comp Rate",
        "Delta",
        "Source or Networx Information",
        "Case Number",
        "Verizon Case Description",
        "Pricing Element",
        "PR#",
        "Version",
        "OpDiv",
        "SF30 Description",
        "12M+ CLIN"
    ]

    # Create a tab for each PR before merging
    pr_groups = defaultdict(list)
    for rec in fr_overview_records:
        pr_groups[rec["PR#"]].append(rec)

    # Remove default "Sheet"
    default_ws = wb.active
    wb.remove(default_ws)

    # Create PR tabs
    for pr, records in pr_groups.items():
        pr_version = records[0]["Version"]
        ws_pr = wb.create_sheet(title=f"{pr} v{pr_version}")
        
        headers = columns[:9]  # first 9 columns only
        
        # Header row
        for col_idx, header in enumerate(headers, start=1):
            cell = ws_pr.cell(row=1, column=col_idx, value=header)
            if col_idx <= 6:
                cell.fill = green_fill
                cell.font = black_font
            elif 7 <= col_idx <= 9:
                cell.fill = black_fill
                cell.font = white_font
            cell.alignment = center_align
            cell.border = thin_border

        # Data rows
        for row_idx, record in enumerate(records, start=2):
            for col_idx, header in enumerate(headers, start=1):
                value = record.get(header, "")
                c = ws_pr.cell(row=row_idx, column=col_idx, value=value)
                c.alignment = Alignment(vertical="center", wrap_text=True)
                c.border = thin_border

        # Auto-width
        for col in ws_pr.columns:
            max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
            ws_pr.column_dimensions[col[0].column_letter].width = max(12, min(max_len + 3, 45))
        ws_pr.freeze_panes = "A2"

    # Now proceed to merge duplicates and create Overview tab
    merged_records = []
    merge_key_columns = ("Case Number", "Pricing Element")
    concat_columns = {
        "PR#", "Version", "OpDiv", "SF30 Description",
        "12M+ CLIN", "Verizon's Response/HHS Comment", "Source or Networx Information"
    }

    temp_dict = defaultdict(list)
    for record in fr_overview_records:
        key = (record["Case Number"], record["Pricing Element"])
        temp_dict[key].append(record)

    for (case_number, pricing_element), records in temp_dict.items():
        merged_record = {}
        for col in records[0].keys():
            if col in merge_key_columns:
                merged_record[col] = records[0][col]
            elif col in concat_columns:
                values = [str(r[col]).strip() for r in records if str(r[col]).strip()]
                merged_record[col] = "/".join(values) if values else ""
            else:
                values = [str(r[col]).strip() for r in records if str(r[col]).strip()]
                merged_record[col] = max(values, key=len) if values else ""
        merged_records.append(merged_record)

    fr_overview_records = merged_records

    # Create Overview tab
    ws = wb.create_sheet(title="Overview")
    for col_idx, header in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        if col_idx <= 6:
            cell.fill = green_fill
            cell.font = black_font
        elif 7 <= col_idx <= 9:
            cell.fill = black_fill
            cell.font = white_font
        elif 10 <= col_idx <= 11:
            cell.fill = gray_fill
            cell.font = black_font
        elif 12 <= col_idx <= 13:
            cell.fill = yellow_fill
            cell.font = black_font
        elif col_idx == 14:
            cell.fill = light_green_fill
            cell.font = black_font
        cell.alignment = center_align
        cell.border = thin_border

    for row_idx, record in enumerate(fr_overview_records, start=2):
        for col_idx, header in enumerate(columns, start=1):
            value = record.get(header, "")
            c = ws.cell(row=row_idx, column=col_idx, value=value)
            if 10 <= col_idx <= 11:
                c.fill = gray_fill
            elif 12 <= col_idx <= 13:
                c.fill = yellow_fill
            elif col_idx == 14:
                c.fill = light_green_fill
            c.alignment = Alignment(vertical="center", wrap_text=True)
            c.border = thin_border

    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max(12, min(max_len + 3, 45))
    ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=len(columns)).coordinate}"
    ws.freeze_panes = "A2"

    wb.save(f_r_output_file)
    print(f"\nF&R Overview document created with PR tabs and merged duplicates: {f_r_output_file}")
    return f_r_output_file


if __name__ == "__main__":
    configure_runtime()
    build_FR()
