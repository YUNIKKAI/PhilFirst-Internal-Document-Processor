import os
import tempfile
import zipfile
import pandas as pd
from datetime import datetime
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

# Constants
CASHCALL_COLUMNS = [ 
    'Branch', 'Line', 'Reinsurer', 'Assured', 'Policy Number',
    'Claim Number', 'FLA Number', 'FLA Date', 'Loss Date', 'Aging',
    'Total Amount Due'
]

COLUMN_WIDTHS = {
    'Branch': 13,
    'Line': 11,
    'Reinsurer': 17.875,
    'Assured': 21.5,
    'Policy Number': 15,
    'Claim Number': 12,
    'FLA Number': 12,
    'FLA Date': 11,
    'Loss Date': 11,
    'Aging': 13.25,
    'Total Amount Due': 13
}

def make_filename_safe(name: str) -> str:
    """Clean reinsurer name for use in filenames"""
    name = str(name).strip()
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    name = re.sub(r'\s+', ' ', name)
    return name[:100]

def calculate_aging_cashcall(fla_date):
    """Calculate aging based on FLA Date for cash call"""
    try:
        if pd.isna(fla_date):
            return 'CURRENT'
        
        if isinstance(fla_date, str):
            fla_date = pd.to_datetime(fla_date, format='mixed', dayfirst=False)
        
        today = datetime.now()
        days_diff = (today - fla_date).days
        
        if days_diff <= 30:
            return 'CURRENT'
        elif days_diff <= 60:
            return 'Over 30 days'
        elif days_diff <= 90:
            return 'Over 60 days'
        elif days_diff <= 120:
            return 'Over 90 days'
        elif days_diff <= 180:
            return 'Over 120 days'
        elif days_diff <= 360:
            return 'Over 180 days'
        else:
            return 'Over 360 days'
    except:
        return 'CURRENT'

def clean_cashcall_bulk_data(df):
    """Remove rows where only 'Assured' has data and other critical columns are empty"""
    print("Cleaning bulk data...")
    original_count = len(df)
    
    critical_columns = ['Reinsurer', 'Policy Number', 'Claim Number']
    
    mask = pd.Series([True] * len(df))
    
    for col in critical_columns:
        if col in df.columns:
            col_empty = df[col].isna() | (df[col].astype(str).str.strip() == '') | (df[col].astype(str).str.strip() == 'nan')
            mask = mask & col_empty
    
    df_cleaned = df[~mask].copy()
    
    removed_count = original_count - len(df_cleaned)
    print(f"  Original rows: {original_count}")
    print(f"  Removed rows: {removed_count}")
    print(f"  Remaining rows: {len(df_cleaned)}")
    
    return df_cleaned

def load_merge_config():
    """Load merge configuration from merge_cashcall.py"""
    try:
        from soa_reinsurer.merge_reinsurers import merge_cashcall, cashcall_rename
        return merge_cashcall, cashcall_rename
    except (ImportError, AttributeError):
        return [], {}

def normalize_reinsurer_name(name):
    """Normalize reinsurer name for comparison"""
    return str(name).strip().upper()

def find_merge_group(reinsurer_name, merge_list):
    """Find which merge group a reinsurer belongs to"""
    reinsurer_normalized = normalize_reinsurer_name(reinsurer_name)
    
    for group in merge_list:
        for member in group:
            if normalize_reinsurer_name(member) == reinsurer_normalized:
                return group
    return None

def process_cashcall(files):
    """Process cashcall files (requires 2 files: bulk and summary)"""
    print("Processing Cash Call files...")
    
    bulk_file = None
    summary_file = None
    
    for file in files:
        filename_lower = file.filename.lower()
        print(f"Checking file: {file.filename}")
        if 'bulk' in filename_lower:
            bulk_file = file
            print(f"  -> Identified as BULK")
        elif 'summary' in filename_lower:
            summary_file = file
            print(f"  -> Identified as SUMMARY")
    
    if not bulk_file or not summary_file:
        print("Auto-assigning files based on position...")
        bulk_file = files[0]
        summary_file = files[1] if len(files) > 1 else None
        print(f"  Bulk: {bulk_file.filename}")
        print(f"  Summary: {summary_file.filename if summary_file else 'None'}")
    
    if not summary_file:
        raise ValueError("Cash Call requires 2 files: bulk data and summary with loss date")
    
    print("Reading CSV files...")
    df_bulk = pd.read_csv(bulk_file)
    df_bulk.columns = df_bulk.columns.str.strip()
    df_details = pd.read_csv(summary_file, header=7)
    df_details.columns = df_details.columns.str.strip()
    print(f"Bulk file rows (before cleaning): {len(df_bulk)}")
    print(f"Summary file rows: {len(df_details)}")
    print(f"Bulk columns: {list(df_bulk.columns)}")
    
    df_bulk = clean_cashcall_bulk_data(df_bulk)
    
    if len(df_bulk) == 0:
        raise ValueError("No valid data found in bulk file after cleaning")
    
    df_bulk_copy = df_bulk.copy()
    df_details_copy = df_details.copy()
    
    print("Parsing FLA dates...")
    df_bulk_copy['match_fla_date'] = pd.to_datetime(
        df_bulk_copy['FLA Date'], 
        format='mixed',
        dayfirst=False,
        errors='coerce'
    )
    
    df_details_copy['match_fla_date'] = pd.to_datetime(
        df_details_copy['FLA DATE'], 
        format='mixed',
        dayfirst=False,
        errors='coerce'
    )
    
    print("Merging dataframes on FLA Date match only...")
    merged_df = df_bulk_copy.merge(
        df_details_copy[['match_fla_date', 'LOSS DATE']],
        on=['match_fla_date'],
        how='left'
    )
    print(f"Merged rows: {len(merged_df)}")
    
    print("Populating Loss Date with matched values or '-'...")
    df_bulk_copy['Loss Date'] = merged_df['LOSS DATE'].fillna('-')
    print(f"  Loss Date populated: {len(df_bulk_copy)} rows")
    print(f"  Rows with Loss Date value: {(df_bulk_copy['Loss Date'] != '-').sum()}")
    print(f"  Rows with '-': {(df_bulk_copy['Loss Date'] == '-').sum()}")
    
    print("Calculating aging...")
    df_bulk_copy['Aging'] = df_bulk_copy['FLA Date'].apply(calculate_aging_cashcall)
    
    print(f"Columns in df_bulk_copy: {list(df_bulk_copy.columns)}")
    
    available_columns = [col for col in CASHCALL_COLUMNS if col in df_bulk_copy.columns]
    print(f"CASHCALL_COLUMNS: {CASHCALL_COLUMNS}")
    print(f"Actual columns in df_bulk_copy: {list(df_bulk_copy.columns)}")
    print(f"Available columns after filter: {available_columns}")
    df_processed = df_bulk_copy[available_columns].copy()
    
    if 'Total Amount Due' in df_processed.columns:
        print("Processing Total Amount Due...")
        df_processed['Total Amount Due'] = pd.to_numeric(
            df_processed['Total Amount Due'].astype(str).str.replace(',', ''), 
            errors='coerce'
        )
        print(f"  Processed {len(df_processed)} rows for Total Amount Due")
    
    merge_list, rename_map = load_merge_config()
    
    output_dfs = []
    processed_reinsurers = set()
    
    if 'Reinsurer' in df_processed.columns:
        reinsurers = df_processed['Reinsurer'].dropna().unique()
        
        for reinsurer in reinsurers:
            reinsurer_normalized = normalize_reinsurer_name(reinsurer)
            
            if reinsurer_normalized in processed_reinsurers:
                continue
            
            merge_group = find_merge_group(reinsurer, merge_list)
            
            if merge_group:
                # Merge group: combine all members
                merged_sections = []
                master_name = None
                
                for group_member in merge_group:
                    group_member_data = df_processed[
                        df_processed['Reinsurer'].apply(
                            lambda x: normalize_reinsurer_name(x) == normalize_reinsurer_name(group_member)
                        )
                    ].copy()
                    
                    if not group_member_data.empty:
                        processed_reinsurers.add(normalize_reinsurer_name(group_member))
                        
                        if master_name is None:
                            master_name = rename_map.get(group_member, group_member)
                        
                        merged_sections.append((group_member, group_member_data))
                
                if merged_sections and master_name:
                    output_dfs.append((master_name, merged_sections, True))
            else:
                # Single reinsurer
                reinsurer_df = df_processed[df_processed['Reinsurer'] == reinsurer].copy()
                processed_reinsurers.add(reinsurer_normalized)
                
                output_dfs.append((reinsurer, reinsurer_df, False))
    
    print(f"Created {len(output_dfs)} output dataframes")
    return output_dfs

def apply_cashcall_formatting(file_path, reinsurer_name):
    """Apply formatting to single cashcall Excel file"""
    wb = load_workbook(file_path)
    ws = wb.active
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center')
    
    # Insert header rows - header starts at row 9
    ws.insert_rows(1, 8)
    header_row = 9
    
    # Add titles
    ws['A1'] = 'PHILIPPINE FIRST INSURANCE CO. INC'
    ws['A1'].font = Font(bold=True, size=12)
    
    ws['A2'] = 'STATEMENT OF ACCOUNT - CASH CALL'
    ws['A2'].font = Font(bold=True)
    
    today = datetime.now().strftime("AS OF %B %d, %Y").upper()
    ws['A3'] = today
    ws['A3'].font = Font(bold=True)
    
    ws['A5'] = 'NEW CASH CALL'
    ws['A5'].font = Font(bold=True)
    
    ws['A7'] = reinsurer_name
    ws['A7'].font = Font(bold=True)
    
    # Format header row
    for col_idx, cell in enumerate(ws[header_row], 1):
        if cell.value:
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border
    
    # Apply borders to data rows
    data_start_row = header_row + 1
    data_end_row = ws.max_row
    
    for row in range(data_start_row, data_end_row + 1):
        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).border = thin_border
    
    # Set column widths
    for col_idx, cell in enumerate(ws[header_row], 1):
        header_val = cell.value
        col_letter = chr(64 + col_idx)
        
        if header_val in COLUMN_WIDTHS:
            ws.column_dimensions[col_letter].width = COLUMN_WIDTHS[header_val]
        else:
            ws.column_dimensions[col_letter].width = 15
    
    # Find Total Amount Due column index
    amount_col = None
    for col_idx, cell in enumerate(ws[header_row], 1):
        if cell.value == 'Total Amount Due':
            amount_col = col_idx
            break
    
    # Calculate totals and aging summary
    total_amount = 0
    aging_summary = {
        'CURRENT': 0,
        'Over 30 days': 0,
        'Over 60 days': 0,
        'Over 90 days': 0,
        'Over 120 days': 0,
        'Over 180 days': 0,
        'Over 360 days': 0
    }
    
    aging_col = None
    for col_idx, cell in enumerate(ws[header_row], 1):
        if cell.value == 'Aging':
            aging_col = col_idx
            break
    
    for row in range(data_start_row, data_end_row + 1):
        if amount_col:
            amount_cell = ws.cell(row=row, column=amount_col)
            if amount_cell.value and isinstance(amount_cell.value, (int, float)):
                total_amount += amount_cell.value
                # Format negative values as (value) in red
                if amount_cell.value < 0:
                    amount_cell.value = abs(amount_cell.value)
                    amount_cell.number_format = '(#,##0.00)'
                    amount_cell.font = Font(color='FF0000')
                else:
                    amount_cell.number_format = '#,##0.00'
        
        if aging_col:
            aging_cell = ws.cell(row=row, column=aging_col)
            aging_val = aging_cell.value
            amount_cell = ws.cell(row=row, column=amount_col) if amount_col else None
            
            if aging_val and amount_cell and amount_cell.value:
                aging_str = str(aging_val).strip()
                if aging_str in aging_summary:
                    try:
                        amt = float(str(amount_cell.value).replace(',', ''))
                        aging_summary[aging_str] += amt
                    except (ValueError, TypeError):
                        pass
    
    # Add subtotal row
    subtotal_row = data_end_row + 1
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=subtotal_row, column=col)
        cell.border = thin_border
        if col == amount_col:
            cell.value = total_amount
            cell.number_format = '#,##0.00'
            cell.font = Font(bold=True)
    
    current_row = subtotal_row + 2
    
    # Add aging summary
    ws.cell(row=current_row, column=1).value = 'AGING SUMMARY'
    ws.cell(row=current_row, column=1).font = Font(bold=True)
    current_row += 1
    
    for aging_label in ['CURRENT', 'Over 30 days', 'Over 60 days', 'Over 90 days', 'Over 120 days', 'Over 180 days', 'Over 360 days']:
        ws.cell(row=current_row, column=1).value = aging_label
        aging_cell = ws.cell(row=current_row, column=2)
        aging_cell.value = aging_summary.get(aging_label, 0)
        aging_cell.number_format = '#,##0.00'
        aging_cell.font = Font(underline='single')
        current_row += 1
    
    # Total aging row
    total_aging = sum(aging_summary.values())
    ws.cell(row=current_row, column=1).value = 'Total'
    ws.cell(row=current_row, column=1).font = Font(bold=True)
    total_aging_cell = ws.cell(row=current_row, column=2)
    total_aging_cell.value = total_aging
    total_aging_cell.font = Font(bold=True, underline='single')
    total_aging_cell.number_format = '#,##0.00'
    current_row += 2
    
    # === Add footer at the end (only once) ===
    current_row += 1
    italic_fmt = Font(italic=True)
    bold_fmt = Font(bold=True)
    regular_fmt = Font()
    italic_lines = {
        "Thank you for trusting your insurance needs with Philippines First Insurance Co., Inc. (PFIC)",
        "Under the Insurance Code: NO INSURANCE POLICY is VALID & BINDING until it is fully paid.",
        "For your convenience, you may pay your insurance premium using the following payment channels:"
    }
    bold_lines = {
        "1. BDO Bills Payment", "a. BDO Mobile Application", "b. Over the Counter",
        "2. BPI Bills Payment", "a. BPI Mobile Application or BPI Online Banking",
        "b. Over the Counter using BPI Express Assist (BEA) Machine",
        "NOTE: Please make checks payable to PHILIPPINES FIRST INSURANCE CO., INC"
    }
    footer_lines = [
        "Thank you for trusting your insurance needs with Philippines First Insurance Co., Inc. (PFIC)",
        "Under the Insurance Code: NO INSURANCE POLICY is VALID & BINDING until it is fully paid.",
        "For your convenience, you may pay your insurance premium using the following payment channels:",
        "", "1. BDO Bills Payment", "a. BDO Mobile Application",
        "   i. Biller: Philippines First Insurance Co., Inc.", "   ii. Reference Number: Policy Invoice Number (Bill Number)",
        "b. Over the Counter", "   i. Company Name: Philippines First Insurance Co., Inc.",
        "   ii. Subscriber Name: Assured Name", "   iii. Subscriber Account Number: Billing Invoice Number",
        "", "2. BPI Bills Payment", "a. BPI Mobile Application or BPI Online Banking",
        "   i. Biller: Philippines First Insurance Co or PFSINC(for short name)", "   ii. Reference Number: Billing Invoice Number",
        "b. Over the Counter using BPI Express Assist (BEA) Machine",
        "   i. Transaction: Bills Payment", "   ii. Merchant: Other Merchant",
        "   iii. Reference Number: Billing Invoice Number", "",
        "NOTE: Please make checks payable to PHILIPPINES FIRST INSURANCE CO., INC",
    ]
    for line in footer_lines:
        cell = ws.cell(row=current_row, column=1)
        cell.value = line
        if line in italic_lines:
            cell.font = italic_fmt
        elif line in bold_lines:
            cell.font = bold_fmt
        else:
            cell.font = regular_fmt
        current_row += 1
    
    wb.save(file_path)

def apply_cashcall_formatting_merged(file_path, reinsurer_groups):
    """Apply formatting to merged cashcall Excel file with separate sections per reinsurer"""
    wb = load_workbook(file_path)
    ws = wb.active
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    header_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center')
    separator_font = Font(bold=True, size=11)
    
    # Clear existing content
    ws.delete_rows(1, ws.max_row)
    
    current_row = 1
    
    # Header information
    ws.cell(row=current_row, column=1).value = 'PHILIPPINE FIRST INSURANCE CO. INC'
    ws.cell(row=current_row, column=1).font = Font(bold=True, size=12)
    current_row += 1
    
    ws.cell(row=current_row, column=1).value = 'STATEMENT OF ACCOUNT - CASH CALL'
    ws.cell(row=current_row, column=1).font = Font(bold=True)
    current_row += 1
    
    today = datetime.now().strftime("AS OF %B %d, %Y").upper()
    ws.cell(row=current_row, column=1).value = today
    ws.cell(row=current_row, column=1).font = Font(bold=True)
    current_row += 1
    
    current_row += 1
    
    ws.cell(row=current_row, column=1).value = 'NEW CASH CALL'
    ws.cell(row=current_row, column=1).font = Font(bold=True)
    current_row += 1
    
    current_row += 1
    
    # Track grand totals for all sections
    grand_total_amount = 0
    grand_aging_summary = {
        'CURRENT': 0,
        'Over 30 days': 0,
        'Over 60 days': 0,
        'Over 90 days': 0,
        'Over 120 days': 0,
        'Over 180 days': 0,
        'Over 360 days': 0
    }
    
    # Process each reinsurer section
    for group_idx, (reinsurer_name, reinsurer_df) in enumerate(reinsurer_groups):
        if group_idx > 0:
            # Separator between sections
            current_row += 1
            separator_cell = ws.cell(row=current_row, column=1)
            separator_cell.value = '.' * 100
            separator_cell.font = separator_font
            current_row += 2
        
        # Reinsurer name
        ws.cell(row=current_row, column=1).value = reinsurer_name
        ws.cell(row=current_row, column=1).font = Font(bold=True)
        current_row += 1
        
        current_row += 1
        
        # Header row for this reinsurer's data
        header_row = current_row
        
        for col_idx, col_name in enumerate(reinsurer_df.columns, 1):
            cell = ws.cell(row=header_row, column=col_idx)
            cell.value = col_name
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border
            
            col_letter = chr(64 + col_idx)
            if col_name in COLUMN_WIDTHS:
                ws.column_dimensions[col_letter].width = COLUMN_WIDTHS[col_name]
            else:
                ws.column_dimensions[col_letter].width = 15
        
        current_row += 1
        data_start_row = current_row
        
        # Write data rows
        max_col = len(reinsurer_df.columns)
        total_amount = 0
        amount_col = None
        aging_col = None
        
        # Find column indices
        for idx, col_name in enumerate(reinsurer_df.columns, 1):
            if col_name == 'Total Amount Due':
                amount_col = idx
            elif col_name == 'Aging':
                aging_col = idx
        
        aging_summary = {
            'CURRENT': 0,
            'Over 30 days': 0,
            'Over 60 days': 0,
            'Over 90 days': 0,
            'Over 120 days': 0,
            'Over 180 days': 0,
            'Over 360 days': 0
        }
        
        for _, row_data in reinsurer_df.iterrows():
            for col_idx in range(1, max_col + 1):
                cell = ws.cell(row=current_row, column=col_idx)
                value = row_data.iloc[col_idx - 1]
                
                if pd.notna(value) and value != '':
                    cell.value = value
                
                cell.border = thin_border
                
                # Format numeric columns
                col_name = reinsurer_df.columns[col_idx - 1]
                if col_name == 'Total Amount Due' and pd.notna(value) and value != '':
                    try:
                        num_val = float(str(value).replace(',', ''))
                        total_amount += num_val
                        grand_total_amount += num_val
                        # Format negative values as (value) in red
                        if num_val < 0:
                            cell.value = abs(num_val)
                            cell.number_format = '(#,##0.00)'
                            cell.font = Font(color='FF0000')
                        else:
                            cell.value = num_val
                            cell.number_format = '#,##0.00'
                    except (ValueError, TypeError):
                        pass
                
                # Track aging
                if col_name == 'Aging' and aging_col and amount_col:
                    aging_val = str(value).strip() if pd.notna(value) else ''
                    amount_val = row_data.iloc[amount_col - 1]
                    
                    if aging_val in aging_summary and pd.notna(amount_val):
                        try:
                            amt = float(str(amount_val).replace(',', ''))
                            aging_summary[aging_val] += amt
                            grand_aging_summary[aging_val] += amt
                        except (ValueError, TypeError):
                            pass
            
            current_row += 1
        
        # Subtotal row
        subtotal_row = current_row
        for col in range(1, max_col + 1):
            cell = ws.cell(row=subtotal_row, column=col)
            cell.border = thin_border
            if col == amount_col:
                cell.value = total_amount
                cell.number_format = '#,##0.00'
                cell.font = Font(bold=True)
            else:
                cell.value = ''
        
        current_row += 2
        
        # Aging summary for this section
        ws.cell(row=current_row, column=1).value = 'AGING SUMMARY'
        ws.cell(row=current_row, column=1).font = Font(bold=True)
        current_row += 1
        
        for aging_label in ['CURRENT', 'Over 30 days', 'Over 60 days', 'Over 90 days', 'Over 120 days', 'Over 180 days', 'Over 360 days']:
            ws.cell(row=current_row, column=1).value = aging_label
            aging_cell = ws.cell(row=current_row, column=2)
            aging_cell.value = aging_summary.get(aging_label, 0)
            aging_cell.number_format = '#,##0.00'
            aging_cell.font = Font(underline='single')
            current_row += 1
        
        # Total aging row for section
        total_section_aging = sum(aging_summary.values())
        ws.cell(row=current_row, column=1).value = 'Total'
        ws.cell(row=current_row, column=1).font = Font(bold=True)
        total_aging_cell = ws.cell(row=current_row, column=2)
        total_aging_cell.value = total_section_aging
        total_aging_cell.font = Font(bold=True, underline='single')
        total_aging_cell.number_format = '#,##0.00'
        current_row += 2
    
    # Grand totals - only grand total amount (no grand aging summary)
    current_row += 1
    ws.cell(row=current_row, column=1).value = 'GRAND TOTAL'
    ws.cell(row=current_row, column=1).font = Font(bold=True, size=11)
    grand_total_cell = ws.cell(row=current_row, column=2)
    grand_total_cell.value = grand_total_amount
    grand_total_cell.font = Font(bold=True, size=11, underline='single')
    grand_total_cell.number_format = '#,##0.00'
    current_row += 2
    
    # === Add footer at the end (only once) ===
    italic_fmt = Font(italic=True)
    bold_fmt = Font(bold=True)
    regular_fmt = Font()
    italic_lines = {
        "Thank you for trusting your insurance needs with Philippines First Insurance Co., Inc. (PFIC)",
        "Under the Insurance Code: NO INSURANCE POLICY is VALID & BINDING until it is fully paid.",
        "For your convenience, you may pay your insurance premium using the following payment channels:"
    }
    bold_lines = {
        "1. BDO Bills Payment", "a. BDO Mobile Application", "b. Over the Counter",
        "2. BPI Bills Payment", "a. BPI Mobile Application or BPI Online Banking",
        "b. Over the Counter using BPI Express Assist (BEA) Machine",
        "NOTE: Please make checks payable to PHILIPPINES FIRST INSURANCE CO., INC"
    }
    footer_lines = [
        "Thank you for trusting your insurance needs with Philippines First Insurance Co., Inc. (PFIC)",
        "Under the Insurance Code: NO INSURANCE POLICY is VALID & BINDING until it is fully paid.",
        "For your convenience, you may pay your insurance premium using the following payment channels:",
        "", "1. BDO Bills Payment", "a. BDO Mobile Application",
        "   i. Biller: Philippines First Insurance Co., Inc.", "   ii. Reference Number: Policy Invoice Number (Bill Number)",
        "b. Over the Counter", "   i. Company Name: Philippines First Insurance Co., Inc.",
        "   ii. Subscriber Name: Assured Name", "   iii. Subscriber Account Number: Billing Invoice Number",
        "", "2. BPI Bills Payment", "a. BPI Mobile Application or BPI Online Banking",
        "   i. Biller: Philippines First Insurance Co or PFSINC(for short name)", "   ii. Reference Number: Billing Invoice Number",
        "b. Over the Counter using BPI Express Assist (BEA) Machine",
        "   i. Transaction: Bills Payment", "   ii. Merchant: Other Merchant",
        "   iii. Reference Number: Billing Invoice Number", "",
        "NOTE: Please make checks payable to PHILIPPINES FIRST INSURANCE CO., INC",
    ]
    for line in footer_lines:
        cell = ws.cell(row=current_row, column=1)
        cell.value = line
        if line in italic_lines:
            cell.font = italic_fmt
        elif line in bold_lines:
            cell.font = bold_fmt
        else:
            cell.font = regular_fmt
        current_row += 1
    
    wb.save(file_path)

def extract_soa_reinsurer_cashcall(files):
    """Process SOA for cash call reinsurer"""
    print("=" * 60)
    print("extract_soa_reinsurer_cashcall called")
    print(f"Number of files: {len(files) if files else 0}")
    
    if not files:
        print("ERROR: No files provided")
        return None
    
    if all(f.filename == "" for f in files):
        print("ERROR: All files have empty filenames")
        return None
    
    if len(files) < 2:
        print("ERROR: Less than 2 files provided")
        raise ValueError("Cash Call requires 2 files (bulk and summary)")
    
    temp_dir = tempfile.mkdtemp()
    print(f"Created temp directory: {temp_dir}")
    
    today = datetime.now().strftime("%b %d, %Y")
    print(f"Date string: {today}")
    
    try:
        print("\nCalling process_cashcall...")
        output_dfs = process_cashcall(files)
        print(f"process_cashcall returned {len(output_dfs)} items")
        
        if not output_dfs:
            print("ERROR: No output dataframes created")
            import shutil
            shutil.rmtree(temp_dir, ignore_errors=True)
            return None
        
        excel_files = []
        
        print("\nCreating Excel files...")
        for idx, item in enumerate(output_dfs):
            filename, data, is_merged = item
            
            if is_merged:
                # Merged: data is list of (reinsurer_name, df) tuples
                clean_name = make_filename_safe(filename)
                file_name = f"SOA {clean_name} AS OF {today}.xlsx"
                file_path = os.path.join(temp_dir, file_name)
                
                # Create initial workbook
                initial_df = pd.DataFrame()
                initial_df.to_excel(file_path, index=False, engine='openpyxl')
                
                # Apply merged formatting
                apply_cashcall_formatting_merged(file_path, data)
                
                excel_files.append(file_path)
                print(f"  ✓ Created: {file_name} (Merged - {len(data)} reinsurers)")
            else:
                # Single reinsurer
                clean_name = make_filename_safe(filename)
                file_name = f"SOA {clean_name} AS OF {today}.xlsx"
                file_path = os.path.join(temp_dir, file_name)
                
                data.to_excel(file_path, index=False, engine='openpyxl')
                apply_cashcall_formatting(file_path, filename)
                
                excel_files.append(file_path)
                print(f"  ✓ Created: {file_name}")
        
        zip_filename = f"SOA CASH CALL AS OF {today}.zip"
        zip_path = os.path.join(temp_dir, zip_filename)
        
        print(f"\nCreating ZIP file: {zip_filename}")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in excel_files:
                zipf.write(file_path, os.path.basename(file_path))
                print(f"  Added to ZIP: {os.path.basename(file_path)}")
        
        print(f"\n✓ ZIP file created: {zip_filename}")
        print(f"✓ ZIP path: {zip_path}")
        print(f"✓ Total files in ZIP: {len(excel_files)}")
        
        result = (zip_path, zip_filename, temp_dir)
        print(f"\nReturning: {result}")
        print("=" * 60)
        
        return result
    
    except Exception as e:
        print("=" * 60)
        print(f"ERROR in extract_soa_reinsurer_cashcall: {str(e)}")
        import traceback
        print(traceback.format_exc())
        print("=" * 60)
        
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise e