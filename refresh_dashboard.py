#!/usr/bin/env python3
"""
Work Order Dashboard - Auto Refresh Script
Handles two AppFolio export formats:
  Format A: Beryl's daily email (PowerQueryresult sheet, header at row 20, has PropertyAbbrev + Assigned User)
  Format B: AppFolio scheduled report (full history, property+address in col 0, no header row)
"""

import os, base64, json, math, tempfile, re
from datetime import datetime, timezone
import re
import urllib.request, urllib.parse, re

# ---- CONFIG ----
CLIENT_ID     = os.environ["GMAIL_CLIENT_ID"]
CLIENT_SECRET = os.environ["GMAIL_CLIENT_SECRET"]
REFRESH_TOKEN = os.environ["GMAIL_REFRESH_TOKEN"]

SENDER        = "donotreply@appfolio.com"
SUBJECT_MATCH = "Work Order Automation"

OPEN_STATUSES = {'Assigned', 'New', 'Scheduled', 'Estimate Requested'}

# ---- 1. GET ACCESS TOKEN ----
def get_access_token():
    params = urllib.parse.urlencode({
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "refresh_token": REFRESH_TOKEN,
        "grant_type":    "refresh_token"
    }).encode()
    req = urllib.request.Request("https://oauth2.googleapis.com/token", data=params, method="POST")
    with urllib.request.urlopen(req) as r:
        data = json.loads(r.read())
    if "access_token" not in data:
        raise Exception(f"Failed to get access token: {data}")
    return data["access_token"]

# ---- 2. GMAIL HELPER ----
def gmail(path, token):
    req = urllib.request.Request(
        f"https://gmail.googleapis.com/gmail/v1/users/me/{path}",
        headers={"Authorization": f"Bearer {token}"}
    )
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())

# ---- 3. FIND TODAY'S EMAIL ----
def find_todays_message(token):
    today = datetime.now(timezone.utc).strftime("%Y/%m/%d")
    q = urllib.parse.quote(f'from:{SENDER} subject:"{SUBJECT_MATCH}" after:{today} has:attachment')
    data = gmail(f"messages?q={q}&maxResults=1", token)
    messages = data.get("messages", [])
    if not messages:
        raise Exception(f"No email found from {SENDER} today ({today})")
    return messages[0]["id"]

# ---- 4. DOWNLOAD XLSX ATTACHMENT ----
def get_xlsx_attachment(msg_id, token):
    msg = gmail(f"messages/{msg_id}?format=full", token)
    def find_att(parts):
        for part in parts:
            if part.get("filename","").endswith(".xlsx") or "spreadsheet" in part.get("mimeType",""):
                att_id = part["body"].get("attachmentId")
                if att_id:
                    return att_id, part.get("filename","")
            if "parts" in part:
                r = find_att(part["parts"])
                if r[0]: return r
        return None, None
    att_id, filename = find_att(msg.get("payload",{}).get("parts",[]))
    if not att_id:
        raise Exception("No XLSX attachment found")
    print(f"  Found attachment: {filename}")
    att_data = gmail(f"messages/{msg_id}/attachments/{att_id}", token)
    raw = att_data["data"].replace("-","+").replace("_","/")
    return base64.b64decode(raw + "==")

# ---- 5. PROCESS XLSX ----
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
        # Format A processing
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
                "CreatedAt":    str(row.get("Created At",""))[:10] if row.get("Created At") else None,
                "Description":  str(row.get("Service Request Description","") or "")[:300].strip() or None,
                "URL":          clean(row.get("AppFolio Link")) or clean(row.get("Link")),
            }
            records.append(r)

    else:
        # Format B: AppFolio scheduled report
        # Columns: Property(full) | Priority | WO Type | WO Number | Status | Unit | Created At | Created By | Assigned User
        print("  Format B detected — AppFolio scheduled report, parsing by position")

        # Find where data starts — first row where col 3 looks like a WO number (e.g. 12345-1)
        data_start = 0
        wo_pattern = re.compile(r'^\d{4,6}-\d+$')
        for i in range(len(raw)):
            val = str(raw.iloc[i, 3]).strip()
            if wo_pattern.match(val):
                data_start = i
                print(f"  Data starts at row {i}")
                break

        # Col mapping (0-indexed):
        # 0=URL, 1=PropertyAbbrev(sparse), 2=CreatedAt, 3=WONumber, 4=Unit,
        # 5=Status, 6=AssignedUser, 7=Priority, 8=ResidentRequested, 9=Recurring,
        # 10=CreatedBy, 11=WorkOrderType, 12=Description, 13-15=comments,
        # 16=AppFolioLink

        # Forward-fill PropertyAbbrev (col 1) since it's sparse
        last_prop = None
        prop_col = {}
        for i in range(data_start, len(raw)):
            val = clean(raw.iloc[i, 1])
            if val:
                last_prop = val
            prop_col[i] = last_prop

        # Print first 3 data rows so we can verify column positions
        print("  First 3 data rows for column verification:")
        for di in range(data_start, min(data_start+3, len(raw))):
            print(f"    Row {di}: {[str(v).strip() for v in raw.iloc[di].tolist()]}")

        records = []
        for i in range(data_start, len(raw)):
            row = raw.iloc[i].tolist()
            row = [clean(v) for v in row]

            # Col 3 must be a WO number
            wo = str(row[3]).strip() if row[3] else None
            if not wo or not wo_pattern.match(wo):
                continue

            # Find status — search all columns for a known status value
            status = None
            for col_idx, val in enumerate(row):
                if str(val).strip() in OPEN_STATUSES:
                    status = str(val).strip()
                    break
            if not status:
                continue

            # Col layout (confirmed from logs):
            # 0=Property(full), 1=Priority, 2=WOType, 3=WONumber, 4=Status,
            # 5=Unit, 6=CreatedAt, 7=CreatedBy, 8=AssignedUser(?),
            # 16=JobDescription

            # Find assigned user — search cols 7 onwards for a non-nan person name
            assigned = None
            for ci in range(7, min(12, len(row))):
                v = str(row[ci]).strip() if row[ci] else ''
                if v and v.lower() not in ('nan', 'no', 'yes', '') and not v.startswith('http') and not re.match(r'^\d', v):
                    assigned = v
                    break

            desc = str(row[16])[:300].strip() if len(row) > 16 and row[16] else None
            if desc and desc.lower() == 'nan':
                desc = None

            # URL — col 0 starts with http in this format
            url = str(row[0]).strip() if row[0] and str(row[0]).startswith('http') else None

            # Extract short property name (before the dash)
            prop_raw = prop_col.get(i)
            if prop_raw and ' - ' in prop_raw:
                prop = prop_raw.split(' - ')[0].strip()
            else:
                prop = prop_raw

            r = {
                "Property":     prop,
                "Unit":         str(row[5]).strip() if len(row) > 5 and row[5] else None,
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
    print(f"  Valid records: {len(records)}")
    return records

# ---- 6. BUILD DASHBOARD ----
def build_dashboard(records, date_str):
    with open("dashboard_template.html") as f:
        template = f.read()
    html = template.replace("DATA_PLACEHOLDER", json.dumps(records))
    html = html.replace("DATE_PLACEHOLDER", date_str)
    html = html.replace("COUNT_PLACEHOLDER", str(len(records)))
    return html

# ---- MAIN ----
if __name__ == "__main__":
    print("🔑 Getting access token...")
    token = get_access_token()
    print("✓ Access token obtained")

    print("🔍 Finding today's AppFolio email...")
    msg_id = find_todays_message(token)
    print(f"✓ Found message: {msg_id}")

    print("📎 Downloading XLSX attachment...")
    xlsx_bytes = get_xlsx_attachment(msg_id, token)
    print(f"✓ Downloaded {len(xlsx_bytes):,} bytes")

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
