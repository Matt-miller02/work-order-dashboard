#!/usr/bin/env python3
"""
Work Order Dashboard - Auto Refresh Script
Pulls work order data from a Google Sheet (shared as "anyone with the link
can view") by hitting Google's built-in XLSX export endpoint — no OAuth
or service account needed.

Handles two AppFolio export formats:
  Format A: header row present (has PropertyAbbrev + Assigned User columns)
  Format B: AppFolio scheduled report (full history, property+address in col 0, no header row)
"""

import os, math, tempfile, re
from datetime import datetime, timezone
import urllib.request

# ---- CONFIG ----
# Set DRIVE_FILE_ID as a repo variable/secret if you ever swap sheets;
# otherwise this default is used.
DRIVE_FILE_ID = os.environ.get("DRIVE_FILE_ID", "1KaMScsd7By3MCBjI-Zc0MTpjgvwA9Bz-")
EXPORT_URL = f"https://docs.google.com/spreadsheets/d/{DRIVE_FILE_ID}/export?format=xlsx"

OPEN_STATUSES = {'Assigned', 'New', 'Scheduled', 'Estimate Requested'}

# ---- 1. DOWNLOAD XLSX FROM GOOGLE SHEET (public export link) ----
def download_sheet():
    print(f"  Downloading sheet export (file ID: {DRIVE_FILE_ID})")
    req = urllib.request.Request(EXPORT_URL, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req) as r:
        data = r.read()
    # A private/broken-link sheet returns a small HTML login/error page instead
    # of real xlsx bytes, so we sanity-check size + the xlsx zip signature (PK).
    if not data or len(data) < 1000 or not data.startswith(b"PK"):
        raise Exception(
            "Downloaded file doesn't look like a valid XLSX export. "
            "Check that the sheet is still shared as 'Anyone with the link can view' "
            f"and that the file ID ({DRIVE_FILE_ID}) is correct."
        )
    print(f"  Downloaded {len(data):,} bytes")
    return data

# ---- 2. PROCESS XLSX ----
def process_xlsx(xlsx_bytes):
    import pandas as pd

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        f.write(xlsx_bytes)
        tmp_path = f.name

    raw = pd.read_excel(tmp_path, sheet_name=0, header=None)
    print(f"  Raw shape: {raw.shape}")

    def clean(v):
        if v is None: return None
        if isinstance(v, float) and math.isnan(v): return None
        if hasattr(v, "isoformat"): return str(v)[:10]
        return str(v).strip() if not isinstance(v, (int, float)) else v

    # ---- Detect format ----
    # Format A: has a header row with 'PropertyAbbrev' and 'Assigned User'
    header_row = None
    for i in range(min(30, len(raw))):
        row_vals = [str(v).strip() for v in raw.iloc[i].tolist()]
        if 'PropertyAbbrev' in row_vals and 'Assigned User' in row_vals:
            header_row = i
            print(f"  Format A detected — header at row {i}")
            break

    if header_row is not None:
        # Format A processing — read by column name, not position
        df = pd.read_excel(tmp_path, sheet_name=0, header=header_row)
        df['PropertyAbbrev'] = df['PropertyAbbrev'].ffill()
        records = []
        for _, row in df.iterrows():
            wo = clean(row.get("Work Order Number"))
            status = clean(row.get("Status"))
            if not wo or not status:
                continue
            r = {
                "Property":     clean(row.get("PropertyAbbrev")),
                "Unit":         clean(row.get("Unit")),
                "Status":       status,
                "Priority":     clean(row.get("Priority")),
                "Type":         clean(row.get("Work Order Type")),
                "WONumber":     wo,
                "AssignedUser": clean(row.get("Assigned User")),
                "CreatedAt":    str(row.get("Created At", ""))[:10] if row.get("Created At") else None,
                "Description":  str(row.get("Service Request Description", "") or "")[:300].strip() or None,
                "URL":          clean(row.get("AppFolio Link")) or clean(row.get("WorkOrderURLLink")) or clean(row.get("Link")),
            }
            records.append(r)

    else:
        # Format B: AppFolio scheduled report — positional parsing fallback
        print("  Format B detected — AppFolio scheduled report, parsing by position")

        data_start = 0
        wo_pattern = re.compile(r'^\d{4,6}-\d+$')
        for i in range(len(raw)):
            val = str(raw.iloc[i, 3]).strip()
            if wo_pattern.match(val):
                data_start = i
                print(f"  Data starts at row {i}")
                break

        # Confirmed col layout from live logs:
        # 0=Property(full+address), 1=Priority, 2=WOType, 3=WONumber, 4=Status,
        # 5=Unit, 6=CreatedAt, 7=CreatedBy, 8=AssignedUser, 16=JobDescription

        records = []
        for i in range(data_start, len(raw)):
            row = [clean(v) for v in raw.iloc[i].tolist()]

            wo = str(row[3]).strip() if len(row) > 3 and row[3] else None
            if not wo or not wo_pattern.match(wo):
                continue

            status = str(row[4]).strip() if len(row) > 4 and row[4] else None
            if not status or status not in OPEN_STATUSES:
                continue

            prop_raw = str(row[0]).strip() if row[0] else None
            if prop_raw and ' - ' in prop_raw:
                prop = prop_raw.split(' - ')[0].strip()
            elif prop_raw and prop_raw.lower() != 'nan':
                prop = prop_raw
            else:
                prop = None

            assigned = str(row[8]).strip() if len(row) > 8 and row[8] and str(row[8]).lower() != 'nan' else None
            desc = str(row[16])[:300].strip() if len(row) > 16 and row[16] and str(row[16]).lower() != 'nan' else None

            url = str(row[0]).strip() if row[0] and str(row[0]).startswith('http') else None
            if url and 'appfolio' not in url:
                url = None

            r = {
                "Property":     prop,
                "Unit":         str(row[5]).strip() if len(row) > 5 and row[5] and str(row[5]).lower() != 'nan' else None,
                "Status":       status,
                "Priority":     str(row[1]).strip() if len(row) > 1 and row[1] else None,
                "Type":         str(row[2]).strip() if len(row) > 2 and row[2] else None,
                "WONumber":     wo,
                "AssignedUser": assigned,
                "CreatedAt":    str(row[6])[:10] if len(row) > 6 and row[6] else None,
                "Description":  desc,
                "URL":          url,
            }
            records.append(r)

    os.unlink(tmp_path)

    # Deduplicate — merge rows with same WO number, combining assigned users
    seen = {}
    for r in records:
        wo = r['WONumber']
        if wo not in seen:
            seen[wo] = r
        else:
            existing = seen[wo]['AssignedUser'] or ''
            new_user = r['AssignedUser'] or ''
            if new_user and new_user not in existing:
                seen[wo]['AssignedUser'] = (existing + ', ' + new_user).strip(', ')
            if r['Description'] and (not seen[wo]['Description'] or len(r['Description']) > len(seen[wo]['Description'])):
                seen[wo]['Description'] = r['Description']

    records = list(seen.values())
    print(f"  Valid records after dedup: {len(records)}")
    return records

# ---- 3. BUILD DASHBOARD ----
def build_dashboard(records, date_str):
    import json
    with open("dashboard_template.html") as f:
        template = f.read()
    html = template.replace("DATA_PLACEHOLDER", json.dumps(records))
    html = html.replace("DATE_PLACEHOLDER", date_str)
    html = html.replace("COUNT_PLACEHOLDER", str(len(records)))
    return html

# ---- MAIN ----
if __name__ == "__main__":
    print("📥 Downloading sheet from Google Drive...")
    xlsx_bytes = download_sheet()

    print("⚙️  Processing data...")
    records = process_xlsx(xlsx_bytes)
    print(f"✓ {len(records)} work orders processed")

    date_str = datetime.now(timezone.utc).strftime("%B %-d, %Y")

    print("🏗️  Building dashboard...")
    html = build_dashboard(records, date_str)

    with open("index.html", "w") as f:
        f.write(html)
    print(f"✓ Dashboard written ({len(html):,} chars)")
    print("✅ Done!")
