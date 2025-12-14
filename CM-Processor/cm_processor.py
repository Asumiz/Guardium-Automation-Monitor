"""
Guardium CM Processor — Complete Health Check
===================================================

This script automates the Guardium Health Check, processing logs and generating detailed 
reports in Word and Excel formats.

Expected Repository Structure:
/
├── CM-Processor/
│   └── cm_processor.py <--- (This file)
└── CM/ <--- (Working folder created at the root)
    ├── Central Management/
    ├── STAP status/
    ...

✔ Clean directory structure (No "Processos Internos" or "Tabelas Internas")
✔ Automatic cleanup of files from previous runs
✔ STAP:
    - Active vs. Inactive count
    - Detailed list of INACTIVE agents in Word (Host + Version)
✔ Aggregation:
    - Detects failures (Purge, Archive, Export)
    - Captures the failure DATE
    - Reports in Word: Collector - Failure - Date

Dependencies:
    pip install pandas openpyxl python-docx
"""

import sys
from pathlib import Path
from datetime import datetime
import pandas as pd

try:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
except Exception:
    # This check ensures the script runs even if docx is not installed, 
    # but print a warning if report generation is attempted.
    Document = None

# =========================================================
# CONFIGURATION
# =========================================================

# Folder names used in the CM directory 
FOLDER_CM = "Central Management"
FOLDER_STAP = "STAP status"
FOLDER_AGGREGATION = "Aggregation Processes"
FOLDER_QUALITY = "Collection Quality"
FOLDER_OUTPUT = "output"

BASE_SUBFOLDERS = [
    FOLDER_CM,
    FOLDER_STAP,
    FOLDER_AGGREGATION,
    FOLDER_QUALITY,
]

ACTIVE_KEYWORDS = ("active", "up", "running", "connected", "online")
INACTIVE_KEYWORDS = ("inactive", "down", "stopped", "disconnected", "offline", "failed", "error")
SUCCESS_KEYWORDS = ("success", "done", "completed", "ok")

# =========================================================
# HELPER FUNCTIONS
# =========================================================

def read_table(path: Path) -> pd.DataFrame:
    try:
        if path.suffix.lower() in (".xls", ".xlsx"):
            return pd.read_excel(path)
        if path.suffix.lower() == ".csv":
            return pd.read_csv(path)
    except Exception as e:
        print(f"⚠ Error reading {path.name}: {e}")
    return pd.DataFrame()

def find_column(df, keywords):
    """Finds a column based on keywords (case insensitive)"""
    for c in df.columns:
        c_str = str(c).lower()
        for k in keywords:
            if k in c_str:
                return c
    return None

def clean_folder_contents(folder: Path, recursive=False):
    """Cleans files from previous runs to ensure fresh data"""
    if not folder.is_dir():
        return
    
    for item in folder.iterdir():
        if item.is_file():
            try:
                item.unlink()
            except Exception:
                pass
        elif recursive and item.is_dir():
            # Only clean children recursively, not the child folders themselves
            clean_folder_contents(item, recursive=False)

# =========================================================
# LOGIC: COLLECTOR IDENTIFICATION
# =========================================================

def extract_collectors(cm_folder: Path):
    collectors = set()
    for file in cm_folder.iterdir():
        if not file.is_file(): continue
        df = read_table(file)
        if df.empty: continue

        unit_name_col = find_column(df, ["unit name"])
        unit_type_col = find_column(df, ["unit type"])

        if not unit_name_col or not unit_type_col: continue

        for _, row in df.iterrows():
            unit_type = str(row.get(unit_type_col, "")).lower()
            unit_name = str(row.get(unit_name_col, "")).strip()
            if "collector" in unit_type and unit_name:
                collectors.add(unit_name)
    return sorted(collectors)

# =========================================================
# LOGIC: STAP STATUS (WITH VERSION & INACTIVE FILTER)
# =========================================================

def process_stap_status(folder: Path):
    records = []
    for file in folder.iterdir():
        if not file.is_file(): continue
        df = read_table(file)
        if df.empty: continue

        status_col = find_column(df, ["status"])
        host_col = find_column(df, ["software stap host", "stap host", "host"])
        ver_col = find_column(df, ["revision", "version", "s-tap revision", "stap revision"])

        if not status_col: continue

        for _, row in df.iterrows():
            status = str(row.get(status_col, "")).strip()
            host = str(row.get(host_col, "")).strip() if host_col else "N/A"
            version = str(row.get(ver_col, "")).strip() if ver_col else "Undef."
            
            if status:
                is_active = False
                status_lower = status.lower()
                if any(k in status_lower for k in ACTIVE_KEYWORDS):
                    is_active = True
                
                records.append({
                    "host": host,
                    "status": status,
                    "version": version,
                    "is_active": is_active
                })

    if not records:
        return pd.DataFrame(), {"active": 0, "inactive": 0, "total": 0}

    df_all = pd.DataFrame(records)
    
    active_count = df_all[df_all["is_active"] == True].shape[0]
    inactive_count = df_all[df_all["is_active"] == False].shape[0]

    return df_all, {
        "active": active_count,
        "inactive": inactive_count,
        "total": len(df_all)
    }

# =========================================================
# LOGIC: AGGREGATION (WITH DATE)
# =========================================================

def analyze_aggregation_errors(base_folder: Path, collectors: list):
    """
    Returns a list of dictionaries with found errors:
    [{'collector': X, 'activity': Y, 'status': Z, 'date': D}, ...]
    """
    issues = []

    for collector in collectors:
        collector_path = base_folder / collector
        if not collector_path.exists():
            continue

        for file in collector_path.iterdir():
            if not file.is_file(): continue

            df = read_table(file)
            if df.empty: continue

            act_col = find_column(df, ["activity type", "activity", "process"])
            status_col = find_column(df, ["status", "execution status"])
            date_col = find_column(df, ["start time", "run time", "timestamp", "date"])

            if not act_col or not status_col:
                continue

            for _, row in df.iterrows():
                status_val = str(row.get(status_col, "")).strip()
                activity_val = str(row.get(act_col, "")).strip()
                
                date_val = str(row.get(date_col, "")).strip() if date_col else "Undef. Date"

                if not status_val or not activity_val:
                    continue

                is_success = any(ok_word in status_val.lower() for ok_word in SUCCESS_KEYWORDS)
                
                if not is_success:
                    issues.append({
                        "collector": collector,
                        "activity": activity_val,
                        "status": status_val,
                        "date": date_val
                    })

    try:
        issues.sort(key=lambda x: (x['collector'], x['date']), reverse=True)
    except:
        pass
        
    return issues

# =========================================================
# MAIN
# =========================================================

def main(base_path="."):
    base = Path(base_path).resolve()
    cm = base / "CM"
    
    print("--- STARTING GUARDIUM HEALTH CHECK ---")

    # 1. CLEANUP
    print("🧹 Cleaning files from previous runs...")
    cm.mkdir(exist_ok=True) # Ensure CM exists before cleaning its subdirs
    clean_folder_contents(cm / FOLDER_CM)
    clean_folder_contents(cm / FOLDER_STAP)
    clean_folder_contents(cm / FOLDER_AGGREGATION, recursive=True)
    clean_folder_contents(cm / FOLDER_QUALITY, recursive=True)
    
    # 2. CREATE STRUCTURE
    for sf in BASE_SUBFOLDERS:
        (cm / sf).mkdir(exist_ok=True)

    print("✔ Directory structure verified.")
    
    # --- PROMPT 1 ---
    input(f"\n➡ Place the Central Management spreadsheet in the '{FOLDER_CM}' folder and press ENTER...")

    collectors = extract_collectors(cm / FOLDER_CM)
    if not collectors:
        print("❌ No Collectors found.")
        return

    print(f"\n✔ {len(collectors)} Collectors identified.")

    # 3. CREATE COLLECTOR SUBFOLDERS
    for collector in collectors:
        (cm / FOLDER_AGGREGATION / collector).mkdir(parents=True, exist_ok=True)
        (cm / FOLDER_QUALITY / collector).mkdir(parents=True, exist_ok=True)

    # --- PROMPT 2 ---
    input("\n➡ Now place the files into the subfolders (STAP status, Aggregation/Collector X) and press ENTER to start processing...")

    # 4. PROCESSING AND ANALYSIS
    print("\n⚙ Processing data...")
    
    # A) STAP
    stap_df, stap_sum = process_stap_status(cm / FOLDER_STAP)
    
    # B) Aggregation
    agg_issues = analyze_aggregation_errors(cm / FOLDER_AGGREGATION, collectors)

    # 5. REPORT GENERATION
    out = cm / FOLDER_OUTPUT
    out.mkdir(exist_ok=True)

    # Excel Output
    with pd.ExcelWriter(out / "CM_report.xlsx", engine="openpyxl") as writer:
        if not stap_df.empty:
            stap_df.to_excel(writer, sheet_name="STAP_Inventory", index=False)
        if agg_issues:
            pd.DataFrame(agg_issues).to_excel(writer, sheet_name="Aggregation_Errors", index=False)

    # Word Output
    if Document:
        doc = Document()
        doc.add_heading("Guardium Health Check Report", 0)
        doc.add_paragraph(f"Generated on: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

        # --- SECTION 1: STAP ---
        doc.add_heading("1. Agent Status (STAP)", level=1)
        doc.add_paragraph(f"Total agents detected: {stap_sum['total']}")
        doc.add_paragraph(f"Active Agents: {stap_sum['active']}")
        p_inativos = doc.add_paragraph()
        p_inativos.add_run(f"Inactive Agents: {stap_sum['inactive']}").bold = True

        # Inactive Table
        if stap_sum['inactive'] > 0:
            doc.add_heading("Detail of Inactive Agents:", level=3)
            
            inativos_df = stap_df[stap_df["is_active"] == False]

            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Host'
            hdr_cells[1].text = 'Status'
            hdr_cells[2].text = 'Version (Revision)'

            for _, row in inativos_df.iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = str(row['host'])
                row_cells[1].text = str(row['status'])
                row_cells[2].text = str(row['version'])
        else:
            doc.add_paragraph("✔ All reported agents are active.")

        # --- SECTION 2: AGGREGATION ---
        doc.add_heading("2. Aggregation Process Failures", level=1)
        
        if agg_issues:
            doc.add_paragraph("The following process failures were detected (Purge, Export, Archive, etc.):")
            
            err_table = doc.add_table(rows=1, cols=3)
            err_table.style = 'Table Grid'
            eh_cells = err_table.rows[0].cells
            eh_cells[0].text = 'Collector / Appliance'
            eh_cells[1].text = 'Failure (Process/Status)'
            eh_cells[2].text = 'Occurrence Date'

            for issue in agg_issues:
                row_cells = err_table.add_row().cells
                row_cells[0].text = str(issue['collector'])
                row_cells[1].text = f"{issue['activity']} ({issue['status']})"
                row_cells[2].text = str(issue['date'])
        else:
            doc.add_paragraph("No critical errors found in the provided aggregation logs.")

        # Using the Portuguese filename as specified in the README.md output section
        word_path = out / "Relatorio_Executivo.docx"
        doc.save(word_path)
        print(f"📄 Word Report generated: {word_path}")

    print("\n✅ PROCESSING COMPLETE")
    print(f"Inactive Agents: {stap_sum['inactive']}")
    print(f"Aggregation Errors: {len(agg_issues)}")

if __name__ == "__main__":
    arg = sys.argv[1] if len(sys.argv) > 1 else "."
    main(arg)
