import os
import tempfile
import zipfile
import pandas as pd
from datetime import datetime
import re

# Constants
PREMIUM_COLUMNS = [
    'Reinsurer', 'Address', 'Currency', 'Currency Rate', 'Line', 'Date', 
    'Aging', 'Our Policy No.', 'Invoice No.', 'Bord Date', 'Inst No.', 
    'Due Date', 'Assured Policy No.', 'Binder No.', 'Balance Due', 'REMARKS'
]

CASHCALL_COLUMNS = [
    'Branch', 'Line', 'Reinsurer', 'Assured', 'Policy Number', 
    'Claim Number', 'FLA Number', 'FLA Date', 'Loss Date', 'Aging',
    'Total Share', 'Total Payments', 'Total Amount Due'
]

def make_filename_safe(name: str) -> str:
    """Clean reinsurer name for use in filenames"""
    name = str(name).strip()
    # Remove illegal characters for filenames
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    # Replace multiple spaces with single space
    name = re.sub(r'\s+', ' ', name)
    return name[:100]  # Limit length

def determine_aging_premium(row):
    """Determine aging based on which column has numerical data for Premium"""
    columns = ['CURRENT', 'OVER 30 DAYS', 'OVER 60 DAYS', 'OVER 90 DAYS', 'OVER 120 DAYS', 'OVER 180 DAYS']
    
    for col in columns:
        if col in row and pd.notna(row[col]) and str(row[col]).strip() != '-':
            try:
                val = float(str(row[col]).replace(',', ''))
                if val != 0:
                    if col in ['CURRENT', 'OVER 30 DAYS', 'OVER 60 DAYS', 'OVER 90 DAYS']:
                        return 'Within 120days-PPW'
                    elif col == 'OVER 120 DAYS':
                        return 'OVER 120 DAYS'
                    elif col == 'OVER 180 DAYS':
                        return 'OVER 180 DAYS'
            except:
                continue
    
    return 'Within 120days-PPW'

def calculate_aging_cashcall(fla_date):
    """Calculate aging based on FLA Date for cash call"""
    try:
        if pd.isna(fla_date):
            return 'CURRENT'
        
        if isinstance(fla_date, str):
            fla_date = pd.to_datetime(fla_date)
        
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

def process_premium(files):
    """Process premium files"""
    print("Processing Premium files...")
    
    # Read the CSV file
    df = pd.read_csv(files[0])
    
    # Add new columns
    df['Aging'] = df.apply(determine_aging_premium, axis=1)
    df['REMARKS'] = ''
    
    # Select and reorder columns
    available_columns = [col for col in PREMIUM_COLUMNS if col in df.columns]
    df_processed = df[available_columns].copy()
    
    # Convert Balance Due to numeric
    if 'Balance Due' in df_processed.columns:
        df_processed['Balance Due'] = pd.to_numeric(
            df_processed['Balance Due'].astype(str).str.replace(',', ''), 
            errors='coerce'
        )
    
    # Group by Reinsurer
    output_dfs = []
    if 'Reinsurer' in df_processed.columns:
        reinsuers = df_processed['Reinsurer'].dropna().unique()
        
        for reinsurer in reinsuers:
            reinsurer_df = df_processed[df_processed['Reinsurer'] == reinsurer].copy()
            
            if 'Balance Due' in reinsurer_df.columns:
                total_balance = reinsurer_df['Balance Due'].sum()
                
                # Create total row
                total_row = pd.DataFrame({col: [''] for col in reinsurer_df.columns})
                total_row['Reinsurer'] = f'TOTAL - {reinsurer}'
                total_row['Balance Due'] = total_balance
                
                reinsurer_with_total = pd.concat([reinsurer_df, total_row], ignore_index=True)
                output_dfs.append((reinsurer, reinsurer_with_total))
            else:
                output_dfs.append((reinsurer, reinsurer_df))
    
    return output_dfs

def process_cashcall(files):
    """Process cashcall files (requires 2 files: bulk and summary)"""
    print("Processing Cash Call files...")
    
    # Identify which file is bulk and which is summary
    bulk_file = None
    summary_file = None
    
    for file in files:
        filename_lower = file.filename.lower()
        if 'bulk' in filename_lower:
            bulk_file = file
        elif 'summary' in filename_lower:
            summary_file = file
    
    # If not identified by name, use order (first = bulk, second = summary)
    if not bulk_file or not summary_file:
        bulk_file = files[0]
        summary_file = files[1] if len(files) > 1 else None
    
    if not summary_file:
        raise ValueError("Cash Call requires 2 files: bulk data and summary with loss date")
    
    # Read both files
    df_bulk = pd.read_csv(bulk_file)
    df_details = pd.read_csv(summary_file)
    
    # Standardize column names for matching
    df_bulk_copy = df_bulk.copy()
    df_details_copy = df_details.copy()
    
    # Create matching keys
    df_bulk_copy['match_reinsurer'] = df_bulk_copy['Reinsurer'].astype(str).str.strip().str.upper()
    df_bulk_copy['match_policy'] = df_bulk_copy['Policy Number'].astype(str).str.strip().str.upper()
    df_bulk_copy['match_claim'] = df_bulk_copy['Claim Number'].astype(str).str.strip().str.upper()
    df_bulk_copy['match_fla_date'] = pd.to_datetime(df_bulk_copy['FLA Date'], errors='coerce')
    
    df_details_copy['match_reinsurer'] = df_details_copy['REINSURER'].astype(str).str.strip().str.upper()
    df_details_copy['match_policy'] = df_details_copy['ASSURED POLICY NO'].astype(str).str.strip().str.upper()
    df_details_copy['match_claim'] = df_details_copy['CLAIM NO'].astype(str).str.strip().str.upper()
    df_details_copy['match_fla_date'] = pd.to_datetime(df_details_copy['FLA DATE'], errors='coerce')
    
    # Merge to add LOSS DATE
    merged_df = df_bulk_copy.merge(
        df_details_copy[['match_reinsurer', 'match_policy', 'match_claim', 'match_fla_date', 'LOSS DATE']],
        on=['match_reinsurer', 'match_policy', 'match_claim', 'match_fla_date'],
        how='left'
    )
    
    # Add Loss Date and Aging
    df_bulk_copy['Loss Date'] = merged_df['LOSS DATE']
    df_bulk_copy['Aging'] = df_bulk_copy['FLA Date'].apply(calculate_aging_cashcall)
    
    # Select final columns
    available_columns = [col for col in CASHCALL_COLUMNS if col in df_bulk_copy.columns]
    df_processed = df_bulk_copy[available_columns].copy()
    
    # Convert Total Amount Due to numeric
    if 'Total Amount Due' in df_processed.columns:
        df_processed['Total Amount Due'] = pd.to_numeric(
            df_processed['Total Amount Due'].astype(str).str.replace(',', ''), 
            errors='coerce'
        )
    
    # Group by Reinsurer
    output_dfs = []
    if 'Reinsurer' in df_processed.columns:
        reinsuers = df_processed['Reinsurer'].dropna().unique()
        
        for reinsurer in reinsuers:
            reinsurer_df = df_processed[df_processed['Reinsurer'] == reinsurer].copy()
            
            if 'Total Amount Due' in reinsurer_df.columns:
                total_amount = reinsurer_df['Total Amount Due'].sum()
                
                # Create total row
                total_row = pd.DataFrame({col: [''] for col in reinsurer_df.columns})
                total_row['Reinsurer'] = f'TOTAL - {reinsurer}'
                total_row['Total Amount Due'] = total_amount
                
                reinsurer_with_total = pd.concat([reinsurer_df, total_row], ignore_index=True)
                output_dfs.append((reinsurer, reinsurer_with_total))
            else:
                output_dfs.append((reinsurer, reinsurer_df))
    
    return output_dfs

def extract_soa_reinsurer(files, file_type=None):
    """
    Main function to process SOA for reinsurer
    
    Args:
        files: List of uploaded files
        file_type: 'premium' or 'cash-call'
    
    Returns:
        Tuple of (zip_path, zip_filename, temp_dir)
    """
    if not files or all(f.filename == "" for f in files):
        return None
    
    # Get file type from request or form
    if not file_type:
        return None
    
    # Create temporary directory
    temp_dir = tempfile.mkdtemp()
    today = datetime.now().strftime("%m-%d-%Y")
    
    try:
        # Process based on file type
        if file_type == 'premium':
            if len(files) < 1:
                raise ValueError("Premium requires at least 1 file")
            output_dfs = process_premium(files)
        elif file_type == 'cash-call':
            if len(files) < 2:
                raise ValueError("Cash Call requires 2 files (bulk and summary)")
            output_dfs = process_cashcall(files)
        else:
            raise ValueError(f"Invalid file type: {file_type}")
        
        # Create individual Excel files
        excel_files = []
        for reinsurer_name, reinsurer_df in output_dfs:
            if reinsurer_df.empty:
                continue
            
            # Clean reinsurer name for filename
            clean_name = make_filename_safe(reinsurer_name)
            
            # Create filename
            file_name = f"SoA of {clean_name} as of {today}.xlsx"
            file_path = os.path.join(temp_dir, file_name)
            
            # Save to Excel
            reinsurer_df.to_excel(file_path, index=False, engine='openpyxl')
            excel_files.append(file_path)
            print(f"✓ Created: {file_name}")
        
        # Create ZIP file
        zip_filename = f"SoA Reinsurance as of {today}.zip"
        zip_path = os.path.join(temp_dir, zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in excel_files:
                zipf.write(file_path, os.path.basename(file_path))
        
        print(f"\n✓ ZIP file created: {zip_filename}")
        print(f"✓ Total files in ZIP: {len(excel_files)}")
        
        return zip_path, zip_filename, temp_dir
    
    except Exception as e:
        # Clean up on error
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise e