import os
import re
import tempfile
import zipfile
import pandas as pd
from datetime import datetime

# Columns to drop
UNNECESSARY_COLUMNS = [
    "Issue Date", "Eff Date", "Ref Pol No.", "Assured No.", "Due Date",
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
        clean_name = re.sub(r"(\w+)\s*&\s*(\w+)", r"\1_&_\2", name)
        words = clean_name.split()
        if words and words[-1].upper().strip(".") in suffixes:
            words = words[:-1]
        prefix = " ".join(words[:2])
        prefix = prefix.replace("_&_", " & ")
    return re.sub(r"[^A-Za-z0-9,& ]+", "", prefix)

def extract_soa_direct(files):
    temp_dir = tempfile.mkdtemp()
    excel_files = []
    today = pd.to_datetime(datetime.today().date())

    # Date string for filenames
    first_day_this_month = today.replace(day=1)
    last_day_prev_month = first_day_this_month - pd.Timedelta(days=1)
    date_str = last_day_prev_month.strftime("%B %d, %Y")

    for file in files:
        if not file or file.filename == "":
            continue

        df = pd.read_csv(file)

        # Convert dates
        df["Incept Date"] = pd.to_datetime(df["Incept Date"], errors="coerce")
        df["Eff Date"] = pd.to_datetime(df["Eff Date"], errors="coerce")
        df["Incept Date"] = df["Incept Date"].where(
            df["Eff Date"] <= df["Incept Date"], df["Eff Date"]
        )

        # Days diff + aging
        df["DaysDiff"] = (today - df["Incept Date"]).dt.days
        df["Aging"] = df["DaysDiff"].apply(aging_category)
        df["Incept Date"] = df["Incept Date"].dt.strftime("%m/%d/%Y")

        # Clean columns
        df = df.drop(columns=[col for col in UNNECESSARY_COLUMNS if col in df.columns])
        ordered_cols = [
            "Branch", "Intermediary", "Policy No.", "Incept Date", "Aging",
            "Assured Name", "Invoice No.", "Bill No.",
            "Premium Bal Due", "Tax Bal Due", "Balance Due", "Remarks"
        ]
        df = df[[col for col in ordered_cols if col in df.columns]]

        for col in MONEY_COLS:
            if col in df.columns:
                df[col] = (
                    df[col].astype(str)
                    .str.strip()
                    .str.replace(",", "", regex=False)
                )
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # Group by intermediary
        for name, group in df.groupby("Intermediary"):
            safe_name = str(name).strip()
            totals = {col: "" for col in group.columns}
            for col in MONEY_COLS:
                if col in group.columns:
                    totals[col] = group[col].sum()
            group_with_total = pd.concat([group, pd.DataFrame([totals])], ignore_index=True)

            prefix = make_prefix(safe_name)
            excel_filename = os.path.join(
                temp_dir, f"{prefix}_SOA as of {date_str}.xlsx"
            )
            excel_files.append(excel_filename)

            with pd.ExcelWriter(excel_filename, engine="xlsxwriter") as writer:
                group_with_total.to_excel(writer, index=False, sheet_name="SoA")
                workbook = writer.book
                worksheet = writer.sheets["SoA"]

                money_fmt = workbook.add_format({"num_format": "#,##0.00"})
                total_fmt = workbook.add_format({
                    "num_format": "#,##0.00",
                    "top": 1, "bottom": 2, "bold": True
                })
                last_row = len(group_with_total) + 1

                for col_idx, col in enumerate(group_with_total.columns):
                    max_len = max(
                        group_with_total[col].astype(str).map(len).max(), len(col)
                    ) + 2
                    worksheet.set_column(col_idx, col_idx, max_len)
                    if col in MONEY_COLS:
                        worksheet.set_column(col_idx, col_idx, 15, money_fmt)
                        worksheet.write(last_row - 1, col_idx, group_with_total.iloc[-1][col], total_fmt)

    # Create ZIP
    zip_filename = os.path.join(temp_dir, f"SoA as of {date_str}.zip")
    with zipfile.ZipFile(zip_filename, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in excel_files:
            if os.path.exists(file):
                zipf.write(file, os.path.basename(file))
    return zip_filename, os.path.basename(zip_filename), temp_dir
