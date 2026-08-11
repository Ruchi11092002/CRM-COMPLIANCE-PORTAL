#!/usr/bin/env python3
"""
CRM Compliance Portal Generator - Management Friendly Version

Single Python file. No separate HTML/CSS/JS files.

What this version does:
1. Reads only the Compliance Manager sheet from data/compliance_master.xlsx.
2. If compliance_master.xlsx is not found, it automatically uses the only .xlsx file
   available inside the data folder.
3. Auto-detects the Compliance Manager sheet even if the sheet name has minor
   spacing/case differences.
4. Creates one uploads folder for every compliance as: ID - Compliance activity.
5. Scans all uploaded documents from those folders.
6. Sorts activities from most urgent to least urgent.
7. Alert colour logic: due today = alert, 1-19 days = orange, 20-50 days = blue, >50 days = green.
8. Always overwrites the same output file:
   output/CRM_Compliance_Portal.html
9. Optional watch mode:
   python generate_compliance_portal.py --watch
   This auto-regenerates the same HTML whenever Excel or uploads folder changes.

Important limitation:
A pure static HTML file cannot directly write files into Windows folders reliably
for all users. Therefore, upload is handled by placing files in the relevant
CRM Compliance upload/<ID - Compliance activity>/ folder. If watch mode is running, the portal refreshes
automatically. Otherwise, run the script once again.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import time
import webbrowser
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote

import pandas as pd


# =============================================================================
# CONFIGURATION
# =============================================================================

BASE_DIR = Path(__file__).resolve().parent

EXCEL_PATH = BASE_DIR / "Data" / "compliance_master.xlsx"

# Folder where compliance certificates / documents will be stored.
# Earlier folder name was "uploads". This has now been renamed as requested.
UPLOADS_FOLDER_NAME = "CRM Compliance upload"
UPLOADS_DIR = BASE_DIR / UPLOADS_FOLDER_NAME

# Legacy folder support. If old files exist in the previous "uploads" folder,
# the code will still read/copy them into the new CRM Compliance upload folder.
LEGACY_UPLOADS_FOLDER_NAME = "uploads"
LEGACY_UPLOADS_DIR = BASE_DIR / LEGACY_UPLOADS_FOLDER_NAME

# Local relative URL path used only if SharePoint link mode is disabled.
# Spaces are URL-encoded automatically.
UPLOADS_URL_PREFIX = f"../{quote(UPLOADS_FOLDER_NAME)}/"

# SharePoint / OneDrive web link mode.
# Keep this True for top-management sharing. Document buttons will open files
# through SharePoint web instead of local relative paths.
SHAREPOINT_LINK_MODE = True

# Browser URL base for OneDrive web. Use onedrive.aspx instead of /my
# because file preview links work more reliably when id + parent are passed.
SHAREPOINT_WEB_BASE_URL = (
    "https://jindalstainless0-my.sharepoint.com/"
    "personal/sayak_halder_jindalstainless_com/_layouts/15/onedrive.aspx"
)

# Direct file URL root. This is used for Open Document links.
# It avoids OneDrive render failures seen for other users with onedrive.aspx file links.
SHAREPOINT_SITE_ROOT_URL = "https://jindalstainless0-my.sharepoint.com"

# Root SharePoint path of the CRM Compliance upload folder.
# This was decoded from the browser URL shared by you.
SHAREPOINT_UPLOAD_ROOT_PATH = (
    "/personal/sayak_halder_jindalstainless_com/"
    "Documents/CRM Compliance Project/CRM Compliance upload"
)

OUTPUT_HTML_PATH = BASE_DIR / "output" / "index.html"

# The sheet to be used. Only this sheet is considered.
COMPLIANCE_SHEET_NAME = "Compliance Manager"

# Keep this as CO to match your earlier portal folder IDs like CO-1, CO-2 etc.
# If you want CRM-COMP-001 style IDs later, change this to "CRM-COMP".
COMPLIANCE_ID_PREFIX = "CO"
PAD_NUMERIC_IDS = False  # False => CO-1, CO-2. True => CO-001, CO-002.

ADMIN_PASSWORD = ""
ALERT_DAYS = 20
DATE_FORMAT = "%d-%m-%Y"
AUTO_OPEN_BROWSER = False

# Excel column mapping for Compliance Manager sheet.
COLUMN_MAPPING: Dict[str, str] = {
    "id": "S/N",
    "name": "Activity",
    "area": "Area",
    "responsible": "Responsibility",
    "facilitator": "Facilitator",
    "frequency": "Frequency",
    "last_done": "Compliance Date",
    "next_due": "Next Due Date",
    "status": "Present status",
    "remarks": "Present status",
}


# =============================================================================
# BASIC HELPERS
# =============================================================================

def ensure_project_folders() -> None:
    (BASE_DIR / "data").mkdir(exist_ok=True)
    UPLOADS_DIR.mkdir(exist_ok=True)
    OUTPUT_HTML_PATH.parent.mkdir(exist_ok=True)


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    if text.lower() in {"nan", "nat", "none"}:
        return ""
    return text


def normalize_header(value: str) -> str:
    return re.sub(r"\s+", " ", str(value).strip()).lower()


def parse_date(value: Any) -> Optional[date]:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = clean_text(value)
    if not text:
        return None

    formats = [
        "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y",
        "%Y-%m-%d", "%d-%b-%Y", "%d/%b/%Y", "%d %b %Y",
        "%d-%B-%Y", "%d/%B/%Y", "%d %B %Y",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    try:
        parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
        if pd.notna(parsed):
            return parsed.date()
    except Exception:
        pass

    return None


def format_date(d: Optional[date]) -> str:
    return d.strftime(DATE_FORMAT) if d else ""


def safe_folder_name(value: str) -> str:
    value = clean_text(value)
    value = re.sub(r"[^A-Za-z0-9_-]+", "-", value)
    value = re.sub(r"-+", "-", value).strip("-")
    return value or "UNNAMED"


def safe_activity_folder_text(activity: str, max_len: int = 95) -> str:
    """Return Windows/OneDrive-safe readable activity text for folder names.

    Folder format used by the portal:
        ID - Compliance activity

    The activity text is kept readable for non-technical users but unsafe
    Windows characters are removed/replaced. Long activities are truncated
    to avoid OneDrive/Windows path-length issues.
    """
    text = clean_text(activity)
    # Replace characters not allowed in Windows filenames: < > : " / \ | ? *
    text = re.sub(r'[<>:"/\\|?*]+', ' ', text)
    text = re.sub(r'[\r\n\t]+', ' ', text)
    # Keep readable characters and remove odd symbols that make folders messy.
    text = re.sub(r"[^A-Za-z0-9 .,_()&+'\-]+", ' ', text)
    text = re.sub(r'\s+', ' ', text).strip(' .-_')
    if len(text) > max_len:
        text = text[:max_len].rstrip(' .-_')
    return text or 'Compliance Activity'


def build_upload_folder_name(compliance_id: str, activity_name: str) -> str:
    """Build the final upload folder name: ID - Compliance activity."""
    return f"{compliance_id} - {safe_activity_folder_text(activity_name)}"


def build_compliance_id(raw_id: Any, row_number: int) -> str:
    raw = clean_text(raw_id)

    if raw:
        # Convert 1.0 to 1
        if re.match(r"^\d+\.0$", raw):
            raw = raw.split(".")[0]

        if re.match(r"^\d+$", raw):
            if PAD_NUMERIC_IDS:
                return f"{COMPLIANCE_ID_PREFIX}-{int(raw):03d}"
            return f"{COMPLIANCE_ID_PREFIX}-{int(raw)}"

        return f"{COMPLIANCE_ID_PREFIX}-{safe_folder_name(raw)}"

    # Fallback when S/N is blank
    return f"{COMPLIANCE_ID_PREFIX}-ROW-{row_number + 2}"


def calculate_alert(next_due: Optional[date], status_text: str, today: date) -> Tuple[str, str, int]:
    """Classify compliance alert by next due date.

    This portal is for recurring compliance cycles, so there is no Completed bucket.
    Present Status is shown as information only; alert colour is driven by Next Due Date.

    Logic:
    - Overdue: due date has crossed
    - Due Today: due date is today
    - Due <20 Days: 1 to 19 days remaining, orange
    - Due 20-50 Days: 20 to 50 days remaining, blue
    - Due >50 Days: more than 50 days remaining, green
    - No Due Date: blank/invalid next due date
    """
    if not next_due:
        return "No Due Date", "nodue", 999998

    days = (next_due - today).days

    if days < 0:
        return "Overdue", "overdue", days
    if days == 0:
        return "Due Today", "duetoday", days
    if days < ALERT_DAYS:
        return "Due <20 Days", "duesoon", days
    if days <= 50:
        return "Due 20-50 Days", "midrange", days
    return "Due >50 Days", "safe", days


def file_size_text(size: int) -> str:
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.1f} MB"


def build_sharepoint_path(folder_name: str = "", file_name: str = "") -> str:
    """Build a SharePoint server-relative path for a folder or document."""
    parts = [SHAREPOINT_UPLOAD_ROOT_PATH.strip("/")]
    if folder_name:
        parts.append(folder_name.strip("/"))
    if file_name:
        parts.append(file_name.strip("/"))
    return "/" + "/".join(parts)


def build_sharepoint_direct_file_url(folder_name: str, file_name: str) -> str:
    """Build a direct SharePoint file URL.

    This format is more reliable for users other than the owner because it points
    directly to the file path:
        https://tenant.sharepoint.com/personal/.../Documents/.../file.pdf?web=1

    The previous file link format used:
        _layouts/15/onedrive.aspx?id=<file-path>&parent=<folder-path>

    That can show "unknown render failure" for other users even when they have
    permission and can open the file from the shared folder.
    """
    file_path = build_sharepoint_path(folder_name, file_name)
    return f"{SHAREPOINT_SITE_ROOT_URL}{quote(file_path, safe='/')}?web=1"


def build_upload_url(folder_name: str = "", file_name: str = "") -> str:
    """Return the URL used by the HTML portal to open folders/documents.

    Folder links still use OneDrive folder view:
        onedrive.aspx?id=<folder-path>

    File/checklist links now use direct SharePoint file URLs:
        https://tenant.sharepoint.com/personal/.../Documents/.../file.pdf?web=1
    """
    folder_name = folder_name or ""
    file_name = file_name or ""

    if SHAREPOINT_LINK_MODE:
        if file_name:
            return build_sharepoint_direct_file_url(folder_name, file_name)

        folder_path = build_sharepoint_path(folder_name)
        return f"{SHAREPOINT_WEB_BASE_URL}?id={quote(folder_path, safe='')}"

    if folder_name and file_name:
        return f"{UPLOADS_URL_PREFIX}{quote(folder_name)}/{quote(file_name)}"
    if folder_name:
        return f"{UPLOADS_URL_PREFIX}{quote(folder_name)}/"
    return UPLOADS_URL_PREFIX

# =============================================================================
# EXCEL DISCOVERY AND LOADING
# =============================================================================

def resolve_excel_file() -> Path:
    """Find the Excel file reliably.

    Priority:
    1. data/compliance_master.xlsx
    2. If missing and exactly one .xlsx file exists in data/, use that.
    3. Otherwise raise a clear error.
    """
    ensure_project_folders()

    if EXCEL_PATH.exists():
        return EXCEL_PATH

    data_dir = BASE_DIR / "Data"
    xlsx_files = [p for p in data_dir.glob("*.xlsx") if not p.name.startswith("~$")]

    if len(xlsx_files) == 1:
        print(f"compliance_master.xlsx not found. Using Excel found in data folder: {xlsx_files[0].name}")
        return xlsx_files[0]

    if len(xlsx_files) > 1:
        names = "\n".join(f"- {p.name}" for p in xlsx_files)
        raise FileNotFoundError(
            "Multiple Excel files found in data folder. Rename the master file as compliance_master.xlsx.\n"
            f"Files found:\n{names}"
        )

    raise FileNotFoundError(
        f"Excel file not found: {EXCEL_PATH}\n"
        "Place your master Excel file inside the data folder and rename it to compliance_master.xlsx"
    )


def resolve_sheet_name(excel_file: Path) -> str:
    xls = pd.ExcelFile(excel_file)
    sheets = xls.sheet_names

    # Exact match
    if COMPLIANCE_SHEET_NAME in sheets:
        return COMPLIANCE_SHEET_NAME

    # Case/space-insensitive match
    target_norm = normalize_header(COMPLIANCE_SHEET_NAME)
    for s in sheets:
        if normalize_header(s) == target_norm:
            print(f"Using sheet '{s}' as Compliance Manager sheet.")
            return s

    # Column-based fallback: choose sheet that has Activity and Next Due Date
    required_headers = {normalize_header("Activity"), normalize_header("Next Due Date")}
    for s in sheets:
        preview = pd.read_excel(excel_file, sheet_name=s, nrows=5)
        col_norms = {normalize_header(c) for c in preview.columns}
        if required_headers.issubset(col_norms):
            print(f"Compliance Manager sheet name not found exactly. Auto-detected sheet: {s}")
            return s

    raise ValueError(
        "Could not find the Compliance Manager sheet.\n"
        f"Expected sheet name: {COMPLIANCE_SHEET_NAME}\n"
        f"Available sheets: {sheets}"
    )


def get_column_lookup(df: pd.DataFrame) -> Dict[str, str]:
    """Return normalized header -> actual header mapping."""
    return {normalize_header(c): c for c in df.columns}


def get_value(row: pd.Series, lookup: Dict[str, str], logical_column: str) -> Any:
    wanted = COLUMN_MAPPING.get(logical_column, "")
    if not wanted:
        return ""
    actual = lookup.get(normalize_header(wanted))
    if actual is None:
        return ""
    return row.get(actual, "")


def scan_documents(compliance_id: str, activity_name: str) -> Tuple[List[Dict[str, str]], str, str]:
    """Create and scan the compliance upload folder.

    Current folder format:
        CRM Compliance upload/ID - Compliance activity/

    To avoid breaking old data, the code also reads/copies documents from:
        1. CRM Compliance upload/ID/
        2. uploads/ID - Compliance activity/
        3. uploads/ID/

    Existing files are copied, not deleted, so nothing is lost.
    """
    folder_name = build_upload_folder_name(compliance_id, activity_name)
    folder = UPLOADS_DIR / folder_name
    current_root_id_folder = UPLOADS_DIR / compliance_id
    legacy_root_readable_folder = LEGACY_UPLOADS_DIR / folder_name
    legacy_root_id_folder = LEGACY_UPLOADS_DIR / compliance_id
    folder.mkdir(parents=True, exist_ok=True)

    # Gentle migration from old folders to the new readable folder name.
    # Existing files are copied, not deleted, so nothing is lost.
    migration_sources = [
        current_root_id_folder,
        legacy_root_readable_folder,
        legacy_root_id_folder,
    ]

    try:
        for source_folder in migration_sources:
            if source_folder.exists() and source_folder.is_dir() and source_folder != folder:
                for old_file in source_folder.iterdir():
                    if old_file.is_file():
                        target = folder / old_file.name
                        if not target.exists():
                            try:
                                target.write_bytes(old_file.read_bytes())
                            except Exception:
                                pass
    except Exception:
        pass

    docs: List[Dict[str, str]] = []
    for p in sorted(folder.iterdir(), key=lambda x: x.stat().st_mtime if x.exists() else 0, reverse=True):
        if not p.is_file():
            continue
        stat = p.stat()
        docs.append({
            "name": p.name,
            "url": build_upload_url(folder_name, p.name),
            "folder_url": build_upload_url(folder_name),
            "size": file_size_text(stat.st_size),
            "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%d-%m-%Y %H:%M"),
            "ext": p.suffix.lower().replace(".", "") or "file",
        })
    return docs, folder_name, f"{UPLOADS_FOLDER_NAME}/{folder_name}"


def load_records() -> List[Dict[str, Any]]:
    ensure_project_folders()
    excel_file = resolve_excel_file()
    sheet_name = resolve_sheet_name(excel_file)

    print(f"Reading Excel: {excel_file}")
    print(f"Using sheet: {sheet_name}")

    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    lookup = get_column_lookup(df)

    # Validate that the important columns exist.
    required = ["name", "next_due"]
    missing = []
    for key in required:
        expected = COLUMN_MAPPING[key]
        if normalize_header(expected) not in lookup:
            missing.append(expected)
    if missing:
        raise ValueError(
            "Required columns missing from Compliance Manager sheet.\n"
            f"Missing columns: {missing}\n"
            f"Available columns: {list(df.columns)}"
        )

    today = date.today()
    records: List[Dict[str, Any]] = []

    for idx, row in df.iterrows():
        name = clean_text(get_value(row, lookup, "name"))
        if not name:
            continue

        raw_id = get_value(row, lookup, "id")
        compliance_id = build_compliance_id(raw_id, idx)

        last_done = parse_date(get_value(row, lookup, "last_done"))
        next_due = parse_date(get_value(row, lookup, "next_due"))
        status = clean_text(get_value(row, lookup, "status"))
        remarks = clean_text(get_value(row, lookup, "remarks"))

        alert_status, alert_class, days_remaining = calculate_alert(next_due, status, today)
        documents, folder_name, folder_display = scan_documents(compliance_id, name)

        records.append({
            "id": compliance_id,
            "sheet": sheet_name,
            "name": name,
            "area": clean_text(get_value(row, lookup, "area")),
            "responsible": clean_text(get_value(row, lookup, "responsible")),
            "facilitator": clean_text(get_value(row, lookup, "facilitator")),
            "frequency": clean_text(get_value(row, lookup, "frequency")),
            "last_done": format_date(last_done),
            "next_due": format_date(next_due),
            "status_text": status,
            "remarks": remarks,
            "alert_status": alert_status,
            "alert_class": alert_class,
            "days_remaining": days_remaining,
            "days_display": "-" if alert_class == "nodue" else str(days_remaining),
            "documents": documents,
            "folder_name": folder_name,
            "folder_display": folder_display,
            "folder_url": build_upload_url(folder_name),
        })

    records.sort(key=record_sort_key)
    print(f"Compliance records loaded: {len(records)}")
    print(f"Upload folders available at: {UPLOADS_DIR}")
    return records


def record_sort_key(record: Dict[str, Any]) -> Tuple[int, int, str]:
    rank = {
        "overdue": 0,
        "duetoday": 1,
        "duesoon": 2,
        "midrange": 3,
        "safe": 4,
        "nodue": 5,
    }.get(record.get("alert_class", "safe"), 4)

    days = record.get("days_remaining", 999999)
    try:
        days_int = int(days)
    except Exception:
        days_int = 999999

    return (rank, days_int, record.get("name", ""))


def build_kpis(records: List[Dict[str, Any]]) -> Dict[str, int]:
    kpis = {
        "total": len(records),
        "overdue": 0,
        "due_today": 0,
        "due_soon": 0,
        "midrange": 0,
        "safe": 0,
        "no_due": 0,
        "documents": 0,
    }
    for r in records:
        c = r.get("alert_class")
        if c == "overdue":
            kpis["overdue"] += 1
        elif c == "duetoday":
            kpis["due_today"] += 1
        elif c == "duesoon":
            kpis["due_soon"] += 1
        elif c == "midrange":
            kpis["midrange"] += 1
        elif c == "safe":
            kpis["safe"] += 1
        elif c == "nodue":
            kpis["no_due"] += 1
        kpis["documents"] += len(r.get("documents", []))
    return kpis


# =============================================================================
# HTML GENERATION
# =============================================================================

def generate_html(records: List[Dict[str, Any]]) -> str:
    kpis = build_kpis(records)
    data_json = json.dumps(records, ensure_ascii=False).replace("</", "<\\/")
    kpi_json = json.dumps(kpis)
    generated_on = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    css = r'''

:root{
  --bg:#f6f7f9;
  --surface:#ffffff;
  --surface-2:#fafafa;
  --text:#202124;
  --muted:#6b7280;
  --soft:#9ca3af;
  --line:#e6e8ec;
  --line-dark:#d7dbe2;
  --brand:#1f4e79;
  --brand-soft:#eaf2fb;
  --accent:#355c7d;
  --red:#b91c1c;
  --red-soft:#fee2e2;
  --orange:#b45309;
  --orange-soft:#ffedd5;
  --green:#166534;
  --green-soft:#dcfce7;
  --blue:#1d4ed8;
  --blue-soft:#dbeafe;
  --grey:#475569;
  --grey-soft:#f1f5f9;
  --teal:#0f766e;
  --teal-soft:#ccfbf1;
  --lavender:#eef2ff;
  --cream:#fff8e7;
  --shadow:0 10px 28px rgba(22,34,51,.07);
  --shadow-soft:0 2px 10px rgba(22,34,51,.04);
  --radius:14px;
}

*{box-sizing:border-box}

body{
  margin:0;
  font-family:"Segoe UI",Roboto,Arial,sans-serif;
  background:
    radial-gradient(circle at top left, rgba(31,78,121,.08), transparent 320px),
    linear-gradient(180deg,#fbfcfe 0%,var(--bg) 280px);
  color:var(--text);
  letter-spacing:.01em;
}

body.locked{overflow:hidden}

.portal-content{
  min-height:100vh;
  transition:filter .25s ease, transform .25s ease;
}

.portal-content.locked{
  filter:blur(7px);
  pointer-events:none;
  user-select:none;
}

.header{
  padding:28px 36px 20px;
  display:flex;
  justify-content:space-between;
  gap:24px;
  align-items:flex-start;
  border-bottom:1px solid rgba(206,216,226,.9);
  background:
    linear-gradient(135deg,#edf7ff 0%,#f8fbff 44%,#fff8e8 100%);
  box-shadow:0 2px 12px rgba(31,78,121,.04);
}

.header h1{
  margin:0;
  font-size:30px;
  font-weight:750;
  color:#172033;
  letter-spacing:-.02em;
}

.header p{
  margin:9px 0 0;
  color:#526173;
  font-size:13px;
  line-height:1.5;
}

.header-actions{
  display:flex;
  gap:10px;
  flex-wrap:wrap;
  justify-content:flex-end;
}

.shell{padding:22px 36px 36px}

.btn{
  border:1px solid transparent;
  border-radius:10px;
  padding:9px 13px;
  font-weight:650;
  cursor:pointer;
  background:#edf1f5;
  color:#1f2937;
  text-decoration:none;
  display:inline-flex;
  align-items:center;
  justify-content:center;
  gap:8px;
  white-space:nowrap;
  font-size:13px;
  line-height:1.2;
  transition:background .15s ease, border-color .15s ease, box-shadow .15s ease, transform .15s ease;
}

.btn:hover{
  background:#e7ecf2;
  border-color:#d6dce4;
  box-shadow:var(--shadow-soft);
  transform:translateY(-1px);
}

.btn.primary{background:var(--brand);color:white;border-color:var(--brand)}
.btn.primary:hover{background:#183d61;border-color:#183d61}
.btn.green{background:#0f766e;color:white;border-color:#0f766e}
.btn.ghost{background:white;border-color:var(--line);color:#263445}
.btn.small{padding:7px 10px;border-radius:9px;font-size:12px}

.notice{
  background:linear-gradient(135deg,#fffdf7 0%,#f6fbff 100%);
  border:1px solid #f0dfb9;
  color:#715519;
  border-radius:var(--radius);
  padding:13px 15px;
  font-size:13px;
  margin-bottom:18px;
  line-height:1.5;
  box-shadow:var(--shadow-soft);
}

.notice b{color:#4f3a0b}

.kpis{
  display:grid;
  grid-template-columns:repeat(8,minmax(112px,1fr));
  gap:12px;
  margin-bottom:18px;
}

.kpi{
  background:var(--surface);
  border:1px solid var(--line);
  border-radius:var(--radius);
  padding:15px 16px;
  box-shadow:var(--shadow-soft);
  min-height:88px;
  transition:transform .16s ease, box-shadow .16s ease, border-color .16s ease;
}
.kpi.clickable{cursor:pointer}
.kpi.clickable:hover{transform:translateY(-2px);box-shadow:0 12px 28px rgba(22,34,51,.11);border-color:#cbd5e1}
.kpi.active{outline:3px solid rgba(31,78,121,.18);border-color:#9db8d1}

.kpi .label{
  font-size:11px;
  color:var(--muted);
  font-weight:750;
  text-transform:uppercase;
  letter-spacing:.075em;
}

.kpi .value{
  margin-top:10px;
  font-size:28px;
  font-weight:800;
  letter-spacing:-.025em;
  color:#182333;
}

.kpi.total{border-left:4px solid #5b6b7d;background:linear-gradient(180deg,#ffffff 0%,#f8fafc 100%)}
.kpi.overdue{border-left:4px solid var(--red);background:linear-gradient(180deg,#ffffff 0%,#fff7f7 100%)}
.kpi.duetoday{border-left:4px solid #7f1d1d;background:linear-gradient(180deg,#ffffff 0%,#fff5f5 100%)}
.kpi.duesoon{border-left:4px solid var(--orange);background:linear-gradient(180deg,#ffffff 0%,#fffaf0 100%)}
.kpi.midrange{border-left:4px solid var(--blue);background:linear-gradient(180deg,#ffffff 0%,#f5f9ff 100%)}
.kpi.safe{border-left:4px solid var(--green);background:linear-gradient(180deg,#ffffff 0%,#f5fff8 100%)}
.kpi.nodue{border-left:4px solid var(--grey);background:linear-gradient(180deg,#ffffff 0%,#f8fafc 100%)}
.kpi.overdue .value{color:var(--red)}
.kpi.duetoday .value{color:#7f1d1d}
.kpi.duesoon .value{color:var(--orange)}
.kpi.midrange .value{color:var(--blue)}
.kpi.safe .value{color:var(--green)}
.helpbar{
  display:flex;
  justify-content:space-between;
  align-items:center;
  gap:12px;
  margin:0 0 12px;
  color:#556274;
  font-size:13px;
}
.helpbar strong{color:#263445}

.panel{
  background:var(--surface);
  border:1px solid var(--line);
  border-radius:var(--radius);
  box-shadow:var(--shadow);
  overflow:hidden;
}

.toolbar{
  padding:16px;
  display:grid;
  grid-template-columns:2fr repeat(4,1fr);
  gap:10px;
  border-bottom:1px solid var(--line);
  background:#fcfcfd;
}

.toolbar input,.toolbar select{
  width:100%;
  border:1px solid var(--line-dark);
  border-radius:10px;
  padding:10px 12px;
  font-size:13px;
  background:white;
  color:#1f2937;
  outline:none;
}

.toolbar input::placeholder{color:#9aa3af}

.toolbar input:focus,.toolbar select:focus{
  border-color:#8bb4dd;
  box-shadow:0 0 0 3px rgba(31,78,121,.12);
}

.table-wrap{
  overflow:auto;
  max-height:calc(100vh - 365px);
}

table{
  border-collapse:separate;
  border-spacing:0;
  width:100%;
  min-width:1320px;
  font-size:13px;
}

thead th{
  position:sticky;
  top:0;
  z-index:2;
  background:#f4f6f8;
  border-bottom:1px solid var(--line-dark);
  color:#374151;
  text-align:left;
  padding:12px 11px;
  font-weight:750;
  text-transform:uppercase;
  letter-spacing:.045em;
  font-size:11px;
}

tbody td{
  border-bottom:1px solid #edf0f3;
  padding:11px;
  vertical-align:top;
  color:#2f3744;
}

tbody tr:nth-child(even){background:#fcfcfd}
tbody tr:hover{background:#f5f9fd}

.activity{
  font-weight:750;
  line-height:1.35;
  max-width:390px;
  color:#172033;
}

.status-main{
  max-width:250px;
  color:#263445;
  line-height:1.4;
}

.status-preview{
  display:-webkit-box;
  -webkit-line-clamp:3;
  -webkit-box-orient:vertical;
  overflow:hidden;
  padding:8px 10px;
  border-radius:10px;
  background:#f8fbff;
  border:1px solid #e4edf7;
}

.meta{
  font-size:12px;
  color:var(--muted);
  margin-top:4px;
  line-height:1.4;
}

.badge{
  display:inline-flex;
  align-items:center;
  border-radius:999px;
  padding:5px 10px;
  font-size:11px;
  font-weight:750;
  white-space:nowrap;
  border:1px solid transparent;
}

.badge.overdue{background:var(--red-soft);color:var(--red);border-color:#fecaca}
.badge.duetoday{background:#fee2e2;color:#7f1d1d;border-color:#fecaca}
.badge.duesoon{background:var(--orange-soft);color:var(--orange);border-color:#fed7aa}
.badge.midrange{background:var(--blue-soft);color:var(--blue);border-color:#bfdbfe}
.badge.safe{background:var(--green-soft);color:var(--green);border-color:#bbf7d0}
.badge.nodue{background:var(--grey-soft);color:var(--grey);border-color:#e2e8f0}

.actions{display:flex;gap:6px;flex-wrap:wrap}

.modal{
  display:none;
  position:fixed;
  inset:0;
  background:rgba(16,24,40,.50);
  z-index:1000;
  padding:24px;
  backdrop-filter:blur(2px);
}

.modal.show{
  display:flex;
  align-items:center;
  justify-content:center;
}

.modal-card{
  background:white;
  width:min(980px,96vw);
  max-height:90vh;
  overflow:auto;
  border-radius:18px;
  box-shadow:0 22px 70px rgba(15,23,42,.28);
  border:1px solid rgba(255,255,255,.7);
}

.modal-head{
  position:sticky;
  top:0;
  z-index:1;
  background:rgba(255,255,255,.96);
  backdrop-filter:blur(8px);
  padding:18px 20px;
  border-bottom:1px solid var(--line);
  display:flex;
  justify-content:space-between;
  gap:12px;
  align-items:flex-start;
}

.modal-head h2{
  margin:0;
  font-size:20px;
  line-height:1.3;
  font-weight:750;
  color:#172033;
}

.close{
  border:1px solid var(--line);
  background:#f8fafc;
  border-radius:10px;
  width:38px;
  height:38px;
  cursor:pointer;
  font-size:24px;
  color:#475569;
}

.close:hover{background:#eef2f6}

.modal-body{padding:20px}

.detail-grid{
  display:grid;
  grid-template-columns:repeat(3,1fr);
  gap:12px;
  margin-bottom:18px;
}

.field{
  background:#fbfcfd;
  border:1px solid var(--line);
  border-radius:12px;
  padding:12px;
}

.field .flabel{
  color:var(--muted);
  font-size:10.5px;
  font-weight:750;
  text-transform:uppercase;
  letter-spacing:.07em;
  margin-bottom:7px;
}

.field .fvalue{
  font-size:13px;
  font-weight:650;
  color:#172033;
  line-height:1.4;
  word-break:break-word;
}

.hr{
  height:1px;
  background:var(--line);
  margin:18px 0;
}

.docs{
  display:grid;
  grid-template-columns:repeat(auto-fill,minmax(235px,1fr));
  gap:10px;
}

.doc-card{
  border:1px solid var(--line);
  border-radius:12px;
  padding:12px;
  background:#fff;
  box-shadow:var(--shadow-soft);
}

.doc-name{
  font-weight:750;
  font-size:13px;
  word-break:break-word;
  color:#172033;
}

.doc-meta{
  color:var(--muted);
  font-size:12px;
  margin:7px 0 10px;
}

.doc-actions{
  display:flex;
  gap:8px;
  align-items:center;
  flex-wrap:wrap;
}

.empty{
  color:var(--muted);
  font-size:13px;
  padding:10px 0;
}

.copybox{
  display:flex;
  gap:8px;
  margin-top:10px;
}

.copybox input{
  flex:1;
  border:1px solid var(--line-dark);
  border-radius:10px;
  padding:10px 12px;
  font-size:13px;
  color:#263445;
  background:#fbfcfd;
}

.login-overlay{
  position:fixed;
  inset:0;
  z-index:2000;
  display:flex;
  align-items:center;
  justify-content:center;
  padding:24px;
  background:linear-gradient(135deg,rgba(237,247,255,.78),rgba(255,248,232,.78));
}

.login-card{
  width:min(420px,94vw);
  background:rgba(255,255,255,.96);
  border:1px solid rgba(215,219,226,.95);
  border-radius:20px;
  box-shadow:0 24px 70px rgba(15,23,42,.22);
  padding:24px;
  text-align:left;
}

.login-card h2{
  margin:0;
  font-size:24px;
  color:#172033;
  letter-spacing:-.02em;
}

.login-card p{
  margin:9px 0 18px;
  color:#5f6b7a;
  font-size:13px;
  line-height:1.5;
}

.login-card input{
  width:100%;
  border:1px solid var(--line-dark);
  border-radius:12px;
  padding:12px 13px;
  font-size:15px;
  outline:none;
  margin-bottom:12px;
}

.login-card input:focus{
  border-color:#8bb4dd;
  box-shadow:0 0 0 3px rgba(31,78,121,.12);
}

.login-error{
  min-height:18px;
  color:var(--red);
  font-size:13px;
  font-weight:650;
  margin-top:8px;
}

.toast{
  position:fixed;
  right:24px;
  bottom:24px;
  background:#172033;
  color:white;
  padding:13px 16px;
  border-radius:12px;
  box-shadow:0 16px 44px rgba(15,23,42,.25);
  display:none;
  z-index:2200;
  max-width:450px;
  font-size:13px;
  line-height:1.4;
}

.toast.show{display:block}

@media(max-width:1180px){
  .kpis{grid-template-columns:repeat(4,1fr)}
  .toolbar{grid-template-columns:1fr 1fr}
  .detail-grid{grid-template-columns:1fr 1fr}
}

@media(max-width:720px){
  .header{flex-direction:column;padding:20px}
  .shell{padding:18px 20px 24px}
  .kpis{grid-template-columns:repeat(2,1fr)}
  .toolbar{grid-template-columns:1fr}
  .detail-grid{grid-template-columns:1fr}
  .copybox{flex-direction:column}
}

'''

    js = r'''
const RECORDS = __DATA_JSON__;
const KPIS = __KPI_JSON__;
const PORTAL_PASSWORD = '__ADMIN_PASSWORD__';
let filtered = [...RECORDS];
let currentRecord = null;
let activeCardFilter = '';

function $(id){ return document.getElementById(id); }
function esc(v){ return String(v ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#039;'); }
function toast(msg){ const t=$('toast'); t.textContent=msg; t.classList.add('show'); clearTimeout(window.__toast); window.__toast=setTimeout(()=>t.classList.remove('show'),4500); }

function populateKpis(){
  $('kpiTotal').textContent=KPIS.total;
  $('kpiOverdue').textContent=KPIS.overdue;
  $('kpiDueToday').textContent=KPIS.due_today;
  $('kpiNoDue').textContent=KPIS.no_due;
}
function populateFilters(){
  const areas=[...new Set(RECORDS.map(r=>r.area).filter(Boolean))].sort();
  const resp=[...new Set(RECORDS.map(r=>r.responsible).filter(Boolean))].sort();
  const statuses=[...new Set(RECORDS.map(r=>r.status_text).filter(Boolean))].sort();
  for(const x of areas) $('filterArea').insertAdjacentHTML('beforeend',`<option value="${esc(x)}">${esc(x)}</option>`);
  for(const x of resp) $('filterResponsible').insertAdjacentHTML('beforeend',`<option value="${esc(x)}">${esc(x)}</option>`);
  for(const x of statuses) $('filterStatus').insertAdjacentHTML('beforeend',`<option value="${esc(x)}">${esc(x)}</option>`);
}
function applyFilters(){
  const q=$('searchBox').value.trim().toLowerCase();
  const alert=$('filterAlert').value;
  const area=$('filterArea').value;
  const responsible=$('filterResponsible').value;
  const status=$('filterStatus').value;
  filtered=RECORDS.filter(r=>{
    const hay=`${r.id} ${r.name} ${r.area} ${r.responsible} ${r.facilitator} ${r.frequency} ${r.status_text} ${r.remarks}`.toLowerCase();
    if(q && !hay.includes(q)) return false;
    if(alert && r.alert_class !== alert) return false;
    if(area && r.area !== area) return false;
    if(responsible && r.responsible !== responsible) return false;
    if(status && r.status_text !== status) return false;
    return true;
  });
  renderTable();
}
function clearOtherFilters(){
  $('searchBox').value='';
  $('filterArea').value='';
  $('filterResponsible').value='';
  $('filterStatus').value='';
}
function setActiveCard(card){
  document.querySelectorAll('.kpi.clickable').forEach(el=>el.classList.remove('active'));
  const el=document.querySelector(`.kpi.clickable[data-filter="${card}"]`);
  if(el) el.classList.add('active');
}
function cardFilter(card){
  activeCardFilter = card;
  clearOtherFilters();
  if(card === 'all'){
    $('filterAlert').value='';
    activeCardFilter='';
    setActiveCard('all');
  }else{
    $('filterAlert').value=card;
    setActiveCard(card);
  }
  applyFilters();
  const label = card === 'all' ? 'all compliances' : document.querySelector(`.kpi[data-filter="${card}"] .label`).textContent.toLowerCase();
  toast(`Showing ${label}.`);
}
function renderTable(){
  const body=$('tableBody'); body.innerHTML='';
  if(filtered.length===0){ body.innerHTML=`<tr><td colspan="10" style="text-align:center;color:#64748b;padding:30px;">No matching compliance found.</td></tr>`; return; }
  for(const r of filtered){
    const row=document.createElement('tr');
    row.innerHTML=`
      <td><strong>${esc(r.id)}</strong><div class="meta">${esc(r.sheet)}</div></td>
      <td class="activity">${esc(r.name)}<div class="meta">Folder: ${esc(r.folder_display)}</div></td>
      <td>${esc(r.area || '-')}</td>
      <td>${esc(r.responsible || '-')}</td>
      <td class="status-main"><div class="status-preview">${esc(r.status_text || '-')}</div></td>
      <td>${esc(r.frequency || '-')}</td>
      <td><strong>${esc(r.next_due || '-')}</strong><div class="meta">Last: ${esc(r.last_done || '-')}</div></td>
      <td>${esc(r.days_display)}</td>
      <td><span class="badge ${esc(r.alert_class)}">${esc(r.alert_status)}</span></td>
      <td><div class="actions"><button class="btn small primary" data-act="details">Details</button><a class="btn small ghost" href="${esc(r.folder_url)}" target="_blank">Open Folder</a></div></td>`;
    row.querySelector('[data-act="details"]').addEventListener('click',()=>openModal(r));
    body.appendChild(row);
  }
}
function field(label,value){ return `<div class="field"><div class="flabel">${esc(label)}</div><div class="fvalue">${esc(value || '-')}</div></div>`; }
function openModal(r){
  currentRecord=r;
  $('modalTitle').textContent=r.name;
  $('modalBadge').className=`badge ${r.alert_class}`;
  $('modalBadge').textContent=r.alert_status;
  $('detailGrid').innerHTML=
    field('Compliance ID',r.id)+field('Area',r.area)+field('Responsible',r.responsible)+
    field('Facilitator',r.facilitator)+field('Frequency',r.frequency)+field('Last Done',r.last_done)+
    field('Next Due',r.next_due)+field('Days Remaining',r.days_display)+field('Upload Folder',r.folder_display)+
    field('Present Status',r.status_text)+field('Remarks',r.remarks)+field('Source Sheet',r.sheet);
  renderDocs(r);
  $('folderPath').value=r.folder_display;
  $('openFolderBtn').href=r.folder_url;
  $('modal').classList.add('show');
}
function renderDocs(r){
  const d=$('docsList');
  if(!r.documents || r.documents.length===0){ d.innerHTML=`<div class="empty">No documents uploaded yet. Place files inside <b>${esc(r.folder_display)}</b>. If watch mode is running, this same HTML will update automatically.</div>`; return; }
  d.innerHTML=r.documents.map(doc=>`
    <div class="doc-card">
      <div class="doc-name">${esc(doc.name)}</div>
      <div class="doc-meta">${esc(doc.modified)} · ${esc(doc.size)}</div>
      <div class="doc-actions">
        <a class="btn small ghost" href="${esc(doc.url)}" target="_blank" rel="noopener">Open Document</a>
        <a class="btn small" href="${esc(doc.folder_url || r.folder_url)}" target="_blank" rel="noopener">Open Folder</a>
      </div>
    </div>
  `).join('');
}
function closeModal(){ $('modal').classList.remove('show'); currentRecord=null; }
async function copyFolderPath(){
  try{ await navigator.clipboard.writeText($('folderPath').value); toast('Folder path copied.'); }
  catch(e){ $('folderPath').select(); document.execCommand('copy'); toast('Folder path copied.'); }
}

function unlockPortal(){
  const pwd = $('portalPassword').value;
  const err = $('loginError');
  if(pwd === PORTAL_PASSWORD){
    document.body.classList.remove('locked');
    $('portalContent').classList.remove('locked');
    $('loginOverlay').style.display='none';
    err.textContent='';
    toast('Access granted.');
  }else{
    err.textContent='Incorrect password. Please try again.';
    $('portalPassword').value='';
    $('portalPassword').focus();
  }
}

function lockPortal(){
  closeModal();
  document.body.classList.add('locked');
  $('portalContent').classList.add('locked');
  $('loginOverlay').style.display='flex';
  $('loginError').textContent='';
  $('portalPassword').value='';
  setTimeout(()=>$('portalPassword').focus(), 120);
}

function init(){
  populateKpis(); populateFilters(); renderTable();
  ['searchBox','filterAlert','filterArea','filterResponsible','filterStatus'].forEach(id=>$(id).addEventListener(id==='searchBox'?'input':'change',applyFilters));
  $('closeModal').addEventListener('click',closeModal);
  $('modal').addEventListener('click',e=>{ if(e.target.id==='modal') closeModal(); });
  $('copyFolderBtn').addEventListener('click',copyFolderPath);
  document.querySelectorAll('.kpi.clickable').forEach(card=>card.addEventListener('click',()=>cardFilter(card.dataset.filter)));
  $('unlockBtn').addEventListener('click', unlockPortal);
  $('lockPortalBtn').addEventListener('click', lockPortal);
  $('portalPassword').addEventListener('keydown', e=>{ if(e.key === 'Enter') unlockPortal(); });
  setTimeout(()=>$('portalPassword').focus(), 250);
}
document.addEventListener('DOMContentLoaded',init);
'''

    html = r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CRM Compliance Portal</title>
<style>__CSS__</style>
</head>
<body class="locked">
<div id="portalContent" class="portal-content locked">
<div class="header">
  <div>
    <h1>CRM Compliance Portal</h1>
  </div>
  <div class="header-actions">
    <a class="btn ghost" href="__UPLOADS_URL_PREFIX__" target="_blank">Open CRM Compliance Upload Folder</a>
    <button class="btn ghost" id="lockPortalBtn" type="button">Lock Portal</button>
  </div>
</div>
<div class="shell">
  <div class="kpis">
    <div class="kpi clickable total active" data-filter="all" title="Click to show all compliances"><div class="label">Total</div><div class="value" id="kpiTotal">0</div></div>
    <div class="kpi clickable overdue" data-filter="overdue" title="Click to show overdue compliances"><div class="label">Overdue</div><div class="value" id="kpiOverdue">0</div></div>
    <div class="kpi clickable duetoday" data-filter="duetoday" title="Click to show due today compliances"><div class="label">Due Today</div><div class="value" id="kpiDueToday">0</div></div>
    <div class="kpi clickable nodue" data-filter="nodue" title="Click to show compliances without due date"><div class="label">No Due Date</div><div class="value" id="kpiNoDue">0</div></div>
  </div>
  <div class="helpbar"><div><strong>Tip:</strong> Click Total, Overdue, Due Today, or No Due Date to filter the table instantly.</div><div>Folder naming: <strong>ID - Compliance activity</strong></div></div>
  <div class="panel">
    <div class="toolbar">
      <input id="searchBox" type="text" placeholder="Search compliance, area, owner, present status...">
      <select id="filterAlert"><option value="">All alerts</option><option value="overdue">Overdue</option><option value="duetoday">Due Today</option><option value="duesoon">Due &lt;20 Days</option><option value="midrange">20-50 Days</option><option value="safe">&gt;50 Days</option><option value="nodue">No Due Date</option></select>
      <select id="filterArea"><option value="">All areas</option></select>
      <select id="filterResponsible"><option value="">All responsible</option></select>
      <select id="filterStatus"><option value="">All present status</option></select>
    </div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>ID</th><th>Compliance Activity</th><th>Area</th><th>Responsibility</th><th>Present Status</th><th>Frequency</th><th>Due Date</th><th>Days</th><th>Alert</th><th>Action</th></tr></thead>
        <tbody id="tableBody"></tbody>
      </table>
    </div>
  </div>
</div>
<div class="modal" id="modal">
  <div class="modal-card">
    <div class="modal-head"><div><h2 id="modalTitle"></h2><div style="margin-top:8px"><span id="modalBadge" class="badge"></span></div></div><button id="closeModal" class="close">&times;</button></div>
    <div class="modal-body">
      <div id="detailGrid" class="detail-grid"></div>
      <div class="hr"></div>
      <div style="display:flex;justify-content:space-between;gap:12px;align-items:center;margin-bottom:12px;"><h3 style="margin:0;">Documents</h3><a class="btn small ghost" id="openFolderBtn" href="#" target="_blank">Open This Folder</a></div>
      <div id="docsList" class="docs"></div>
      <div class="hr"></div>
      <h3 style="margin:0 0 8px;">Where to place/upload files</h3>
      <p class="meta">Paste documents into this folder. The portal will show the files after auto-watch refresh or after the next normal generation.</p>
      <div class="copybox"><input id="folderPath" readonly><button id="copyFolderBtn" class="btn primary">Copy Folder Path</button></div>
    </div>
  </div>
</div>
</div>
<div id="loginOverlay" class="login-overlay">
  <div class="login-card">
    <h2>CRM Compliance Portal</h2>
    <p>This portal is password protected. Please enter the access password to view compliance status and documents.</p>
    <input type="password" id="portalPassword" placeholder="Enter password" autocomplete="current-password">
    <button id="unlockBtn" class="btn primary" style="width:100%;justify-content:center;">Open Portal</button>
    <div id="loginError" class="login-error"></div>
  </div>
</div>
<div id="toast" class="toast"></div>
<script>__JS__</script>
</body>
</html>'''

    html = html.replace("__CSS__", css)
    html = html.replace("__JS__", js)
    html = html.replace("__DATA_JSON__", data_json)
    html = html.replace("__KPI_JSON__", kpi_json)
    html = html.replace("__ADMIN_PASSWORD__", ADMIN_PASSWORD.replace("\\", "\\\\").replace("'", "\\'"))
    html = html.replace("__UPLOADS_URL_PREFIX__", build_upload_url())
    html = html.replace("__GENERATED_ON__", generated_on)
    return html


# =============================================================================
# GENERATION AND WATCH MODE
# =============================================================================

def generate_once(open_browser: bool = False) -> None:
    records = load_records()
    html = generate_html(records)
    OUTPUT_HTML_PATH.parent.mkdir(exist_ok=True)
    OUTPUT_HTML_PATH.write_text(html, encoding="utf-8")
    print(f"Portal generated successfully at: {OUTPUT_HTML_PATH}")

    if open_browser:
        try:
            webbrowser.open(OUTPUT_HTML_PATH.resolve().as_uri())
        except Exception:
            pass


def current_watch_state() -> Dict[str, float]:
    state: Dict[str, float] = {}

    excel_file = None
    try:
        excel_file = resolve_excel_file()
    except Exception:
        pass

    if excel_file and excel_file.exists():
        state[str(excel_file)] = excel_file.stat().st_mtime

    for watched_dir in [UPLOADS_DIR, LEGACY_UPLOADS_DIR]:
        if watched_dir.exists():
            for root, dirs, files in os.walk(watched_dir):
                for name in dirs + files:
                    p = Path(root) / name
                    try:
                        state[str(p)] = p.stat().st_mtime
                    except FileNotFoundError:
                        continue

    return state


def watch_loop(interval: float, open_browser: bool = False) -> None:
    print("Starting watch mode...")
    print("Keep this window open. The same HTML file will update when Excel/CRM Compliance upload folder changes.")
    generate_once(open_browser=open_browser)
    previous = current_watch_state()

    while True:
        time.sleep(interval)
        latest = current_watch_state()
        if latest != previous:
            print("Change detected. Regenerating portal...")
            try:
                generate_once(open_browser=False)
                previous = current_watch_state()
            except Exception as exc:
                print(f"Regeneration failed: {exc}")


def main() -> None:
    parser = argparse.ArgumentParser(description="CRM Compliance Portal Generator")
    parser.add_argument("--watch", action="store_true", help="Keep running and auto-regenerate when Excel/CRM Compliance upload folder changes")
    parser.add_argument("--interval", type=float, default=5.0, help="Watch mode refresh interval in seconds")
    parser.add_argument("--no-open", action="store_true", help="Do not auto-open browser after generation")
    args = parser.parse_args()

    ensure_project_folders()

    if args.watch:
        try:
            watch_loop(interval=args.interval, open_browser=not args.no_open and AUTO_OPEN_BROWSER)
        except KeyboardInterrupt:
            print("Watch mode stopped.")
    else:
        generate_once(open_browser=not args.no_open and AUTO_OPEN_BROWSER)


if __name__ == "__main__":
    main()
