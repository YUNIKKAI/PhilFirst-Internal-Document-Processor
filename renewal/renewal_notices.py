import os, re, calendar, tempfile, shutil, zipfile, logging, sys
from datetime import datetime
from flask import current_app as app
from pypdf import PdfReader, PdfWriter
from werkzeug.utils import secure_filename

# Supported policy prefixes
POLICY_PREFIXES = ["AH", "CA", "CG", "CY", "EN", "FG", "FI", "HL", "MC", "MD", "MN", "MR", "PF", "SU"]

# Regex patterns
def create_policy_regex():
    prefix_pattern = '|'.join([f"{prefix}-" for prefix in POLICY_PREFIXES])
    return re.compile(f'((?:{prefix_pattern})[A-Z0-9\\-]+):\\s*Policy\\s*No', re.IGNORECASE)

policy_re = create_policy_regex()
agent_re = re.compile(r'Agent\s*:(.*?)(?:Remarks\s*:|$)', re.DOTALL | re.IGNORECASE)
insured_re = re.compile(
    f'Insured\\s*:(.*?)(?:Plate\\s*No\\.|(?:{"|".join([f"{p}-" for p in POLICY_PREFIXES])})[A-Z0-9\\-]+:\\s*Policy\\s*No|$)', 
    re.DOTALL | re.IGNORECASE
)

# Utility functions
def is_valid_pdf(file_path):
    try:
        with open(file_path, 'rb') as f:
            return f.read(5) == b'%PDF-'
    except Exception:
        return False

def is_supported_policy_prefix(policy_number):
    if not policy_number or '-' not in policy_number:
        return False
    return policy_number.split('-')[0].upper() in POLICY_PREFIXES

def extract_month_year_from_filename(filename):
    name_parts = filename.replace('.pdf', '').split()
    for i, part in enumerate(name_parts):
        for month_num in range(1, 13):
            if part.lower() in [calendar.month_abbr[month_num].lower(), calendar.month_name[month_num].lower()]:
                for j in range(max(0, i-2), min(len(name_parts), i+3)):
                    if name_parts[j].isdigit() and len(name_parts[j]) == 4:
                        return calendar.month_name[month_num], name_parts[j]
    return None, None

def truncate_insured_name_at_inc(name):
    if not name:
        return name
    pos = name.upper().find('INC.')
    return name[:pos + 4].strip() if pos != -1 else name

def sanitize_folder_name(name):
    for char in '<>:"/\\|?*':
        name = name.replace(char, '_')
    name = name.strip().rstrip('.')
    return name[:196] if len(name) > 196 else name

def extract_agent_name(text):
    match = agent_re.search(text)
    return re.sub(r'\s+', ' ', match.group(1).strip()) if match else None

def extract_agent_name_from_pages(reader, page_num):
    try:
        name = extract_agent_name(reader.pages[page_num].extract_text())
        if name:
            return name, page_num
    except Exception:
        pass
    if page_num + 1 < len(reader.pages):
        try:
            name = extract_agent_name(reader.pages[page_num + 1].extract_text())
            if name:
                return name, page_num + 1
        except Exception:
            pass
    return None, None

def extract_insured_name(text):
    match = insured_re.search(text)
    if match:
        name = re.sub(r'\s*\n\s*', ' ', match.group(1).strip())
        return re.sub(r'\s+', ' ', name).strip()
    return None

def has_important_notice(name):
    upper = name.upper()
    return not any(x in upper for x in ['&/OR', 'AND/OR', '&/ OR', 'AND/ OR'])

# Main extraction logic
def extract_renewal_notices(files):
    extracted_data = []
    temp_dir = tempfile.mkdtemp()
    month_name, year = None, None

    for file in files:
        if file and file.filename:
            month_name, year = extract_month_year_from_filename(file.filename)
            if month_name and year:
                break

    folder_name = f"Renewal Notices {month_name or datetime.now().strftime('%B')} {year or datetime.now().year}"
    main_dir = os.path.join(temp_dir, folder_name)
    agents_dir = os.path.join(main_dir, "Agents")
    all_dir = os.path.join(main_dir, "All Renewal Notices")
    os.makedirs(agents_dir, exist_ok=True)
    os.makedirs(all_dir, exist_ok=True)

    try:
        for file in files:
            if not file or file.filename == '':
                continue
            filename = secure_filename(file.filename)
            if not filename.lower().endswith('.pdf'):
                app.logger.warning(f'Skipped non-PDF: {filename}')
                continue
            path = os.path.join(temp_dir, filename)
            file.save(path)
            if not is_valid_pdf(path):
                app.logger.warning(f'Invalid PDF: {filename}')
                continue
            reader = PdfReader(path)
            for i, page in enumerate(reader.pages):
                try:
                    text = page.extract_text()
                    if not text or "RENEWAL NOTICE" not in text.upper():
                        continue
                    match = policy_re.search(text)
                    if not match:
                        continue
                    policy = match.group(1).strip()
                    if not is_supported_policy_prefix(policy):
                        continue
                    agent, agent_page = extract_agent_name_from_pages(reader, i)
                    insured = extract_insured_name(text)
                    if not agent or not insured:
                        continue
                    truncated = truncate_insured_name_at_inc(insured)
                    safe_agent = sanitize_folder_name(agent)
                    safe_insured = sanitize_folder_name(truncated)
                    agent_folder = os.path.join(agents_dir, safe_agent, safe_insured)
                    os.makedirs(agent_folder, exist_ok=True)
                    pdf_name = sanitize_folder_name(f"{truncated} {policy}") + ".pdf"
                    writer = PdfWriter()
                    pages = []
                    if has_important_notice(insured) and i > 0:
                        pages.append(i - 1)
                    pages.append(i)
                    if agent_page != i:
                        pages.append(agent_page)
                    for p in sorted(set(pages)):
                        writer.add_page(reader.pages[p])
                    insured_path = os.path.join(agent_folder, pdf_name)
                    with open(insured_path, "wb") as f:
                        writer.write(f)
                    all_path = os.path.join(all_dir, pdf_name)
                    with open(all_path, "wb") as f:
                        writer.write(f)
                    extracted_data.append(pdf_name)
                except Exception as e:
                    app.logger.error(f'Error on page {i+1} of {filename}: {e}')
        if not extracted_data:
            shutil.rmtree(temp_dir, ignore_errors=True)
            return None
        zip_path = os.path.join(temp_dir, f"{folder_name}.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(main_dir):
                for file in files:
                    zipf.write(os.path.join(root, file), arcname=os.path.relpath(os.path.join(root, file), temp_dir))
        return zip_path, os.path.basename(zip_path), temp_dir
    except Exception as e:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise e
