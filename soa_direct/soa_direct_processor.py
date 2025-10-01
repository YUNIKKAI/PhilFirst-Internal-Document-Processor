import os
import re
import tempfile
import zipfile
import pandas as pd
from datetime import datetime
from soa_direct.merge_and_agent import ACCOUNTS_TO_MERGE, INTERMEDIARY_TO_AGENT

# Columns to drop
UNNECESSARY_COLUMNS = [#"Issue Date"
    "Eff Date", "Assured No.", "Due Date",
    "Advance", "Current", "Over 30 Days", "Over 60 Days",
    "Over 90 Days", "Over 120 Days", "Over 180 Days", "Over 360 Days"
]
MONEY_COLS = ["Premium Bal Due", "Tax Bal Due", "Balance Due"]

def aging_category(days: int) -> str:
    if days < 90:
        return "Within 90days-Credit Term"
    elif days <= 120:
        return "Over 90 days"
    elif days <= 180:
        return "Over 120 days"
    elif days <= 360:
        return "Over 180 days"
    return "Over 360 days"

# Define all aging categories in order
ALL_AGING_CATEGORIES = [
    "Within 90days-Credit Term",
    "Over 90 days",
    "Over 120 days",
    "Over 180 days",
    "Over 360 days"
]

def make_prefix(name: str) -> str:
    """Build filename prefix from intermediary name."""
    name = str(name).strip()
    suffixes = {"JR", "JR.", "SR", "SR.", "III", "IV", "V"}

    if "," in name:
        parts = [p.strip() for p in name.split(",", 1)]
        surname = parts[0]
        firstname_parts = parts[1].split() if len(parts) > 1 else []
        firstname_parts = [p for p in firstname_parts if p.upper().strip(".") not in suffixes]
        firstname = firstname_parts[0] if firstname_parts else ""
        prefix = f"{surname}, {firstname}".strip()
    else:
        clean_name = re.sub(r"(\S+)\s*&\s*(\S+)", r"\1_&_\2", name)
        words = clean_name.split()
        if words and words[-1].upper().strip(".") in suffixes:
            words = words[:-1]
        prefix = " ".join(words[:2])
        prefix = prefix.replace("_&_", " & ")

    # remove illegal characters
    prefix = re.sub(r"[^0-9A-Za-zÑñÁÉÍÓÚÜáéíóúü,& ]+", "", prefix)

    # fallback if empty
    if not prefix:
        prefix = re.sub(r"[^A-Za-z0-9]", "", name)[:10]

    return prefix

def _build_merge_maps(merge_groups_from_arg):
    """
    Build two structures:
    - merge_groups: dict master -> [aliases...] (preserves order)
    - alias_to_master: dict alias_exact_string -> master (exact match only)
    """
    source = merge_groups_from_arg if merge_groups_from_arg is not None else ACCOUNTS_TO_MERGE

    if isinstance(source, dict):
        merge_groups = {str(k).strip(): [str(i).strip() for i in v] for k, v in source.items()}
    else:
        merge_groups = {}
        for grp in source:
            if not grp:
                continue
            master = str(grp[0]).strip()
            aliases = [str(x).strip() for x in grp]
            merge_groups[master] = aliases

    alias_to_master = {}
    for master, aliases in merge_groups.items():
        for alias in aliases:
            alias_to_master[alias] = master  # exact match only

    return merge_groups, alias_to_master

def _is_blank_df_one_row_all_empty(df: pd.DataFrame) -> bool:
    """Detect if df is a single-row DataFrame where every cell is empty string or NaN."""
    if df.shape[0] != 1:
        return False
    as_str = df.fillna("").astype(str).applymap(lambda v: v.strip())
    return (as_str == "").all().all()

def extract_soa_direct(files, merge_groups=None, agent_folders=None):
    agent_folders = agent_folders or INTERMEDIARY_TO_AGENT
    merge_groups_map, alias_to_master = _build_merge_maps(merge_groups)

    temp_dir = tempfile.mkdtemp()
    excel_files = []
    today = pd.to_datetime(datetime.today().date())
    date_str = today.strftime("%B %d, %Y")

    used_prefixes = {}

    combined_rows = []
    for file in files:
        if not file or getattr(file, "filename", "") == "":
            continue
        df = pd.read_csv(file)

        df["Incept Date"] = pd.to_datetime(df.get("Incept Date"), errors="coerce")
        df["Eff Date"] = pd.to_datetime(df.get("Eff Date"), errors="coerce")
        df["Incept Date"] = df["Incept Date"].where(
            df["Eff Date"] <= df["Incept Date"], df["Eff Date"]
        )
        combined_rows.append(df)

    if not combined_rows:
        zip_filename = os.path.join(temp_dir, f"SoA as of {date_str}.zip")
        with zipfile.ZipFile(zip_filename, "w", zipfile.ZIP_DEFLATED):
            pass
        return zip_filename, os.path.basename(zip_filename), temp_dir

    df_all = pd.concat(combined_rows, ignore_index=True)
    df_all["DaysDiff"] = (today - df_all["Incept Date"]).dt.days
    df_all["Aging"] = df_all["DaysDiff"].apply(aging_category)
    df_all["Incept Date"] = df_all["Incept Date"].dt.strftime("%m/%d/%Y")

    if "Remarks" not in df_all.columns:
        df_all["Remarks"] = ""

    df_all = df_all.drop(columns=[col for col in UNNECESSARY_COLUMNS if col in df_all.columns])
    ordered_cols = [
        "Branch", "Intermediary", "Policy No.", "Issue Date", "Incept Date", "Aging", "Ref Pol No.",
        "Assured Name", "Invoice No.", "Bill No.",
        "Premium Bal Due", "Tax Bal Due", "Balance Due", "Remarks"
    ]
    df_all = df_all[[col for col in ordered_cols if col in df_all.columns]]

    for col in MONEY_COLS:
        if col in df_all.columns:
            df_all[col] = (
                df_all[col].astype(str)
                .str.strip()
                .str.replace(",", "", regex=False)
            )
            df_all[col] = pd.to_numeric(df_all[col], errors="coerce").fillna(0)

    # ✅ SORT by Assured Name then Incept Date
    df_all = df_all.sort_values(by=["Assured Name", "Incept Date"], ascending=[True, True]).reset_index(drop=True)

    # --- formatting helper ---
    def apply_formats(workbook, worksheet, sheet_df, subtotal_rows, aging_summary_rows, inter_name, is_merged=False):
        report_header_fmt = workbook.add_format({"bold": True, "align": "left", "font_size": 12})
        header_fmt = workbook.add_format({"bold": True, "border": 1, "align": "center"})
        text_cell_fmt = workbook.add_format({"border": 1, "align": "left"})
        money_cell_fmt = workbook.add_format({"num_format": "#,##0.00", "border": 1, "align": "right"})
        subtotal_fmt = workbook.add_format({
            "num_format": "#,##0.00", "bold": True, "bottom": 2, "align": "right", "border": 1
        })
        blank_cell_fmt = workbook.add_format({"border": 1})
        aging_text_fmt = workbook.add_format({"bold": True, "align": "left"})
        aging_money_fmt = workbook.add_format({"num_format": "#,##0.00", "bold": True, "align": "right"})
        aging_dash_fmt = workbook.add_format({"bold": True, "align": "right"})
        aging_total_text_fmt = workbook.add_format({"bold": True, "align": "left", "top": 1, "bottom": 2})
        aging_total_money_fmt = workbook.add_format({"num_format": "#,##0.00", "bold": True, "align": "right", "top": 1, "bottom": 2})
        no_border_fmt = workbook.add_format({})
        footer_text_fmt = workbook.add_format({"align": "left", "valign": "top"})
        footer_bold_fmt = workbook.add_format({"align": "left", "valign": "top", "bold": True})
        footer_italic_fmt = workbook.add_format({"align": "left", "valign": "top", "italic": True})

        # Report header
        worksheet.write(0, 0, inter_name, report_header_fmt)
        worksheet.write(1, 0, "STATEMENT OF ACCOUNT", report_header_fmt)
        worksheet.write(2, 0, f"AS OF {date_str.upper()}", report_header_fmt)

        startrow = 4
        data_start_row = startrow + 1
        rows, cols = sheet_df.shape

        # Column widths
        for col_idx, col in enumerate(sheet_df.columns):
            max_len = max(sheet_df[col].astype(str).map(len).max(), len(col)) + 2
            if col == "Assured Name":
                worksheet.set_column(col_idx, col_idx, min(max_len, 40))
            elif col == "Remarks":
                worksheet.set_column(col_idx, col_idx, min(max_len, 30))
            elif col in MONEY_COLS:
                worksheet.set_column(col_idx, col_idx, 15)
            else:
                worksheet.set_column(col_idx, col_idx, max_len)

        # Rewrite headers
        for col_idx, col in enumerate(sheet_df.columns):
            worksheet.write(startrow, col_idx, col, header_fmt)

        # Data + subtotal + aging summary
        for r in range(rows):
            for c in range(cols):
                val = sheet_df.iat[r, c]
                excel_row = data_start_row + r
                col_name = sheet_df.columns[c]

                # Check if this row is a subtotal row
                if r in subtotal_rows and col_name in MONEY_COLS:
                    worksheet.write_number(excel_row, c, float(val or 0), subtotal_fmt)
                # Check if this row is part of aging summary
                elif r in aging_summary_rows:
                    aging_info = aging_summary_rows[r]
                    if aging_info['type'] == 'blank':
                        worksheet.write_blank(excel_row, c, None, no_border_fmt)
                    elif aging_info['type'] == 'detail':
                        if c == 0:
                            worksheet.write(excel_row, c, val, aging_text_fmt)
                        elif c == 1:
                            # Check if value is "-" (dash for zero/missing categories)
                            if val == "-":
                                worksheet.write(excel_row, c, val, aging_dash_fmt)
                            else:
                                worksheet.write_number(excel_row, c, float(val or 0), aging_money_fmt)
                        else:
                            worksheet.write_blank(excel_row, c, None, no_border_fmt)
                    elif aging_info['type'] == 'total':
                        if c == 0:
                            worksheet.write(excel_row, c, "Total", aging_total_text_fmt)
                        elif c == 1:
                            worksheet.write_number(excel_row, c, float(val or 0), aging_total_money_fmt)
                        else:
                            worksheet.write_blank(excel_row, c, None, no_border_fmt)
                    elif aging_info['type'] == 'spacing':
                        worksheet.write_blank(excel_row, c, None, no_border_fmt)
                else:
                    if pd.isna(val) or (isinstance(val, str) and val.strip() == ""):
                        worksheet.write_blank(excel_row, c, None, blank_cell_fmt)
                    elif col_name in MONEY_COLS:
                        worksheet.write_number(excel_row, c, float(val), money_cell_fmt)
                    else:
                        worksheet.write(excel_row, c, val, text_cell_fmt)

        # === Footer with payment instructions ===
        last_data_row = data_start_row + rows
        footer_start_row = last_data_row + 3
        
        # Lines that should be italicized
        italic_lines = {
            "Thank you for trusting your insurance needs with Philippines First Insurance Co., Inc. (PFIC)",
            "Under the Insurance Code: NO INSURANCE POLICY is VALID & BINDING until it is fully paid.",
            "For your convenience, you may pay your insurance premium using the following payment channels:"
        }
        
        bold_lines = {
            "1. BDO Bills Payment",
            "a. BDO Mobile Application",
            "b. Over the Counter",
            "2. BPI Bills Payment",
            "a. BPI Mobile Application or BPI Online Banking",
            "b. Over the Counter using BPI Express Assist (BEA) Machine",
            "NOTE: Please make checks payable to PHILIPPINES FIRST INSURANCE CO., INC"
        }
        
        footer_lines = [
            "Thank you for trusting your insurance needs with Philippines First Insurance Co., Inc. (PFIC)",
            "Under the Insurance Code: NO INSURANCE POLICY is VALID & BINDING until it is fully paid.",
            "For your convenience, you may pay your insurance premium using the following payment channels:",
            "",
            "1. BDO Bills Payment",
            "a. BDO Mobile Application",
            "   i. Biller: Philippines First Insurance Co., Inc.",
            "   ii. Reference Number: Policy Invoice Number (Bill Number)",
            "b. Over the Counter",
            "   i. Company Name: Philippines First Insurance Co., Inc.",
            "   ii. Subscriber Name: Assured Name",
            "   iii. Subscriber Account Number: Billing Invoice Number",
            "",
            "2. BPI Bills Payment",
            "a. BPI Mobile Application or BPI Online Banking",
            "   i. Biller: Philippines First Insurance Co or PFSINC(for short name)",
            "   ii. Reference Number: Billing Invoice Number",
            "b. Over the Counter using BPI Express Assist (BEA) Machine",
            "   i. Transaction: Bills Payment",
            "   ii. Merchant: Other Merchant",
            "   iii. Reference Number: Billing Invoice Number",
            "",
            "NOTE: Please make checks payable to PHILIPPINES FIRST INSURANCE CO., INC",
        ]
        for i, line in enumerate(footer_lines):
            if line in italic_lines:
                fmt = footer_italic_fmt
            elif line in bold_lines:
                fmt = footer_bold_fmt
            else:
                fmt = footer_text_fmt
            worksheet.write(footer_start_row + i, 0, line, fmt)

    # === Pass 1: merged accounts ===
    for master, aliases in merge_groups_map.items():
        merged_rows = df_all[df_all["Intermediary"].astype(str).str.strip().isin(aliases)]
        if merged_rows.empty:
            continue

        prefix = make_prefix(master)
        if prefix not in used_prefixes:
            filename_prefix = prefix
            used_prefixes[prefix] = 1
        else:
            used_prefixes[prefix] += 1
            filename_prefix = f"{prefix}_{used_prefixes[prefix]}"

        agent_folder_name = (agent_folders or {}).get(master)
        if agent_folder_name:
            target_folder = os.path.join(temp_dir, "AGENT", str(agent_folder_name))
        else:
            target_folder = temp_dir
        os.makedirs(target_folder, exist_ok=True)

        output_parts, subtotal_row_indexes, aging_summary_row_indexes, running_len = [], [], {}, 0

        def add_block(block_df: pd.DataFrame, is_last_branch=False):
            nonlocal running_len
            if block_df.empty:
                return
            
            output_parts.append(block_df)
            running_len += len(block_df)
            
            # Add subtotal row
            subtotal = {col: "" for col in block_df.columns}
            for mcol in MONEY_COLS:
                subtotal[mcol] = block_df[mcol].sum()
            subtotal_df = pd.DataFrame([subtotal])
            output_parts.append(subtotal_df)
            subtotal_row_indexes.append(running_len)
            running_len += 1
            
            # Add one blank row before aging summary
            blank_row = pd.DataFrame([{col: "" for col in block_df.columns}])
            output_parts.append(blank_row)
            aging_summary_row_indexes[running_len] = {'type': 'blank'}
            running_len += 1
            
            # Calculate aging summary for this branch - include all categories
            aging_summary = block_df.groupby("Aging")["Balance Due"].sum().reset_index()
            aging_dict = dict(zip(aging_summary["Aging"], aging_summary["Balance Due"]))
            
            # Add aging detail rows for ALL categories (no header)
            for aging_cat in ALL_AGING_CATEGORIES:
                detail_row = {col: "" for col in block_df.columns}
                detail_row[block_df.columns[0]] = aging_cat
                if aging_cat in aging_dict:
                    detail_row[block_df.columns[1]] = aging_dict[aging_cat] if len(block_df.columns) > 1 else ""
                else:
                    detail_row[block_df.columns[1]] = "-"  # Display dash for missing categories
                output_parts.append(pd.DataFrame([detail_row]))
                aging_summary_row_indexes[running_len] = {'type': 'detail'}
                running_len += 1
            
            # Add total row
            total_row = {col: "" for col in block_df.columns}
            total_row[block_df.columns[0]] = "Total"
            # Sum only the numeric values (exclude "-")
            total_sum = sum(aging_dict.values())
            total_row[block_df.columns[1]] = total_sum if len(block_df.columns) > 1 else ""
            output_parts.append(pd.DataFrame([total_row]))
            aging_summary_row_indexes[running_len] = {'type': 'total'}
            running_len += 1
            
            # Add 2 no-border spacing rows after aging summary (if not last branch)
            if not is_last_branch:
                for _ in range(2):
                    spacing_row = pd.DataFrame([{col: "" for col in block_df.columns}])
                    output_parts.append(spacing_row)
                    aging_summary_row_indexes[running_len] = {'type': 'spacing'}
                    running_len += 1

        branch_groups = list(merged_rows.groupby("Branch"))
        for idx, (branch_val, branch_group) in enumerate(branch_groups):
            is_last = (idx == len(branch_groups) - 1)
            add_block(branch_group, is_last_branch=is_last)

        sheet_df = pd.concat(output_parts, ignore_index=True)
        excel_filename = os.path.join(target_folder, f"{filename_prefix}_SOA as of {date_str}.xlsx")
        excel_files.append(excel_filename)

        with pd.ExcelWriter(excel_filename, engine="xlsxwriter") as writer:
            sheet_df.to_excel(writer, index=False, sheet_name="SoA", startrow=4)
            apply_formats(writer.book, writer.sheets["SoA"], sheet_df, subtotal_row_indexes, aging_summary_row_indexes, master, is_merged=True)

    # === Pass 2: non-merged ===
    for (branch, name), group in df_all.groupby(["Branch", "Intermediary"]):
        safe_name = str(name).strip() or "UNNAMED"
        if safe_name in alias_to_master:
            continue

        # Build sheet with data, subtotal, blank, aging summary
        output_parts = []
        subtotal_row_indexes = []
        aging_summary_row_indexes = {}
        running_len = 0
        
        # Add data
        output_parts.append(group)
        running_len += len(group)
        
        # Add subtotal
        totals = {col: "" for col in group.columns}
        for mcol in MONEY_COLS:
            totals[mcol] = group[mcol].sum()
        output_parts.append(pd.DataFrame([totals]))
        subtotal_row_indexes.append(running_len)
        running_len += 1
        
        # Add one blank row before aging summary
        blank_row = pd.DataFrame([{col: "" for col in group.columns}])
        output_parts.append(blank_row)
        aging_summary_row_indexes[running_len] = {'type': 'blank'}
        running_len += 1
        
        # Calculate aging summary - include all categories
        aging_summary = group.groupby("Aging")["Balance Due"].sum().reset_index()
        aging_dict = dict(zip(aging_summary["Aging"], aging_summary["Balance Due"]))
        
        # Add aging detail rows for ALL categories (no header)
        for aging_cat in ALL_AGING_CATEGORIES:
            detail_row = {col: "" for col in group.columns}
            detail_row[group.columns[0]] = aging_cat
            if aging_cat in aging_dict:
                detail_row[group.columns[1]] = aging_dict[aging_cat] if len(group.columns) > 1 else ""
            else:
                detail_row[group.columns[1]] = "-"  # Display dash for missing categories
            output_parts.append(pd.DataFrame([detail_row]))
            aging_summary_row_indexes[running_len] = {'type': 'detail'}
            running_len += 1
        
        # Add total row
        total_row = {col: "" for col in group.columns}
        total_row[group.columns[0]] = "Total"
        # Sum only the numeric values (exclude "-")
        total_sum = sum(aging_dict.values())
        total_row[group.columns[1]] = total_sum if len(group.columns) > 1 else ""
        output_parts.append(pd.DataFrame([total_row]))
        aging_summary_row_indexes[running_len] = {'type': 'total'}
        running_len += 1
        
        sheet_df = pd.concat(output_parts, ignore_index=True)

        prefix = make_prefix(safe_name)
        last_word = re.sub(r"[^A-Za-z0-9]", "", safe_name.split()[-1]) if safe_name else "X"
        branch_val, branch_clean = str(branch).strip(), re.sub(r"[^A-Za-z0-9 ]", "", str(branch).strip())

        if prefix not in used_prefixes:
            filename_prefix = prefix
            used_prefixes[prefix] = {branch_clean: 1}
        elif branch_clean not in used_prefixes[prefix]:
            filename_prefix = f"{prefix} ({branch_clean})"
            used_prefixes[prefix][branch_clean] = 2
        else:
            count = used_prefixes[prefix][branch_clean] + 1
            used_prefixes[prefix][branch_clean] = count
            filename_prefix = f"{prefix} ({branch_clean}-{last_word})"

        agent_folder_name = (agent_folders or {}).get(safe_name)
        if agent_folder_name:
            target_folder = os.path.join(temp_dir, "AGENT", str(agent_folder_name))
        else:
            target_folder = temp_dir
        os.makedirs(target_folder, exist_ok=True)

        excel_filename = os.path.join(target_folder, f"{filename_prefix}_SOA as of {date_str}.xlsx")
        excel_files.append(excel_filename)

        with pd.ExcelWriter(excel_filename, engine="xlsxwriter") as writer:
            sheet_df.to_excel(writer, index=False, sheet_name="SoA", startrow=4)
            apply_formats(writer.book, writer.sheets["SoA"], sheet_df, subtotal_row_indexes, aging_summary_row_indexes, safe_name, is_merged=False)

    # Build zip
    zip_filename = os.path.join(temp_dir, f"SoA as of {date_str}.zip")
    with zipfile.ZipFile(zip_filename, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in excel_files:
            if os.path.exists(file):
                arcname = os.path.relpath(file, temp_dir)
                zipf.write(file, arcname)

    return zip_filename, os.path.basename(zip_filename), temp_dir