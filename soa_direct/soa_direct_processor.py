import os
import re
import tempfile
import zipfile
import pandas as pd
from datetime import datetime
from soa_direct.merge_and_agent import ACCOUNTS_TO_MERGE, INTERMEDIARY_TO_AGENT

# Columns to drop
UNNECESSARY_COLUMNS = [#"Issue Date"
    "Eff Date", "Ref Pol No.", "Assured No.", "Due Date",
    "Advance", "Current", "Over 30 Days", "Over 60 Days",
    "Over 90 Days", "Over 120 Days", "Over 180 Days", "Over 360 Days"
]
MONEY_COLS = ["Premium Bal Due", "Tax Bal Due", "Balance Due"]

def aging_category(days: int) -> str:
    if days < 90:
        return "Within the 90days CTE"
    elif days <= 120:
        return "Over 90 days"
    elif days <= 180:
        return "Over 120 days"
    elif days <= 360:
        return "Over 180 days"
    return "Over 360 days"

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
    # prefer arg if provided, else use imported ACCOUNTS_TO_MERGE
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
    # Convert to str and strip — treat '', 'nan', 'None' as empty
    as_str = df.fillna("").astype(str).applymap(lambda v: v.strip())
    return (as_str == "").all().all()


def extract_soa_direct(files, merge_groups=None, agent_folders=None):
    agent_folders = agent_folders or INTERMEDIARY_TO_AGENT
    merge_groups_map, alias_to_master = _build_merge_maps(merge_groups)

    temp_dir = tempfile.mkdtemp()
    excel_files = []
    today = pd.to_datetime(datetime.today().date())

    # Date string for filenames
    first_day_this_month = today.replace(day=1)
    last_day_prev_month = first_day_this_month - pd.Timedelta(days=1)
    date_str = last_day_prev_month.strftime("%B %d, %Y")

    used_prefixes = {}

    # Load CSV files
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
        "Branch", "Intermediary", "Policy No.", "Issue Date", "Incept Date", "Aging",
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

    # --- formatting helper ---
    def apply_formats(workbook, worksheet, sheet_df, subtotal_rows, inter_name):
        report_header_fmt = workbook.add_format({"bold": True, "align": "left", "font_size": 12})
        header_fmt = workbook.add_format({"bold": True, "border": 1, "align": "center"})
        text_cell_fmt = workbook.add_format({"border": 1, "align": "left"})
        money_cell_fmt = workbook.add_format({"num_format": "#,##0.00", "border": 1, "align": "right"})
        subtotal_fmt = workbook.add_format({
            "num_format": "#,##0.00", "bold": True, "bottom": 2, "align": "right", "border": 1
        })
        blank_cell_fmt = workbook.add_format({"border": 1})  # <-- NEW for empty cells

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

        # Data + subtotal
        for r in range(rows):
            for c in range(cols):
                val = sheet_df.iat[r, c]
                excel_row = data_start_row + r
                col_name = sheet_df.columns[c]

                if r in subtotal_rows and col_name in MONEY_COLS:
                    worksheet.write_number(excel_row, c, float(val or 0), subtotal_fmt)
                else:
                    if pd.isna(val) or (isinstance(val, str) and val.strip() == ""):
                        worksheet.write_blank(excel_row, c, None, blank_cell_fmt)  # <-- keep borders
                    elif col_name in MONEY_COLS:
                        worksheet.write_number(excel_row, c, float(val), money_cell_fmt)
                    else:
                        worksheet.write(excel_row, c, val, text_cell_fmt)

    # === Pass 1: merged accounts ===
    for master, aliases in merge_groups_map.items():
        present_aliases = []
        for alias in aliases:
            sub_rows = df_all[df_all["Intermediary"].astype(str).str.strip() == alias]
            if not sub_rows.empty:
                present_aliases.append((alias, sub_rows))
        if not present_aliases:
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

        output_parts, subtotal_row_indexes, running_len = [], [], 0
        def add_block(block_df: pd.DataFrame):
            nonlocal running_len
            if block_df.empty:
                return
            output_parts.append(block_df)
            running_len += len(block_df)
            subtotal = {col: "" for col in block_df.columns}
            for mcol in MONEY_COLS:
                subtotal[mcol] = block_df[mcol].sum()
            subtotal_df = pd.DataFrame([subtotal])
            output_parts.append(subtotal_df)
            running_len += 1
            subtotal_row_indexes.append(running_len - 1)

        for idx, (_alias, sub_rows) in enumerate(present_aliases):
            add_block(sub_rows)
            if idx < len(present_aliases) - 1:
                blanks = pd.DataFrame([{col: "" for col in sub_rows.columns} for _ in range(2)])
                output_parts.append(blanks)
                running_len += 2

        sheet_df = pd.concat(output_parts, ignore_index=True)
        excel_filename = os.path.join(target_folder, f"{filename_prefix}_SOA as of {date_str}.xlsx")
        excel_files.append(excel_filename)

        with pd.ExcelWriter(excel_filename, engine="xlsxwriter") as writer:
            sheet_df.to_excel(writer, index=False, sheet_name="SoA", startrow=4)
            apply_formats(writer.book, writer.sheets["SoA"], sheet_df, subtotal_row_indexes, master)

    # === Pass 2: non-merged ===
    for (branch, name), group in df_all.groupby(["Branch", "Intermediary"]):
        safe_name = str(name).strip() or "UNNAMED"
        if safe_name in alias_to_master:
            continue

        totals = {col: "" for col in group.columns}
        for mcol in MONEY_COLS:
            totals[mcol] = group[mcol].sum()
        group_with_total = pd.concat([group, pd.DataFrame([totals])], ignore_index=True)

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

        sheet_df = group_with_total
        subtotal_row_indexes = [len(sheet_df) - 1]
        excel_filename = os.path.join(target_folder, f"{filename_prefix}_SOA as of {date_str}.xlsx")
        excel_files.append(excel_filename)

        with pd.ExcelWriter(excel_filename, engine="xlsxwriter") as writer:
            sheet_df.to_excel(writer, index=False, sheet_name="SoA", startrow=4)
            apply_formats(writer.book, writer.sheets["SoA"], sheet_df, subtotal_row_indexes, safe_name)

    # Build zip
    zip_filename = os.path.join(temp_dir, f"SoA as of {date_str}.zip")
    with zipfile.ZipFile(zip_filename, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in excel_files:
            if os.path.exists(file):
                arcname = os.path.relpath(file, temp_dir)
                zipf.write(file, arcname)

    return zip_filename, os.path.basename(zip_filename), temp_dir