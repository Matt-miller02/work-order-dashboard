#!/usr/bin/env python3
"""
Work Order Dashboard - Auto Refresh Script
Runs daily via GitHub Actions:
1. Finds today's AppFolio email (donotreply@appfolio.com)
2. Downloads the XLSX attachment
3. Processes the data
4. Injects it into the dashboard HTML template
5. GitHub Actions commits and deploys to GitHub Pages
"""

import os, base64, json, math, tempfile
from datetime import datetime, timezone
import urllib.request, urllib.parse

# ---- CONFIG (all set as GitHub Secrets) ----
CLIENT_ID     = os.environ["GMAIL_CLIENT_ID"]
CLIENT_SECRET = os.environ["GMAIL_CLIENT_SECRET"]
REFRESH_TOKEN = os.environ["GMAIL_REFRESH_TOKEN"]

SENDER        = "donotreply@appfolio.com"
SUBJECT_MATCH = "Work Order Automation"

# ---- 1. GET ACCESS TOKEN ----
def get_access_token():
    params = urllib.parse.urlencode({
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "refresh_token": REFRESH_TOKEN,
        "grant_type":    "refresh_token"
    }).encode()
    req = urllib.request.Request(
        "https://oauth2.googleapis.com/token",
        data=params, method="POST"
    )
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
        raise Exception(f"No email found from {SENDER} with subject '{SUBJECT_MATCH}' today ({today})")
    return messages[0]["id"]

# ---- 4. DOWNLOAD XLSX ATTACHMENT ----
def get_xlsx_attachment(msg_id, token):
    msg = gmail(f"messages/{msg_id}?format=full", token)

    def find_att(parts):
        for part in parts:
            fname = part.get("filename", "")
            if fname.endswith(".xlsx") or "spreadsheet" in part.get("mimeType", ""):
                att_id = part["body"].get("attachmentId")
                if att_id:
                    return att_id, fname
            if "parts" in part:
                result = find_att(part["parts"])
                if result[0]:
                    return result
        return None, None

    parts = msg.get("payload", {}).get("parts", [])
    att_id, filename = find_att(parts)
    if not att_id:
        raise Exception("No XLSX attachment found in email")

    print(f"  Found attachment: {filename}")
    att_data = gmail(f"messages/{msg_id}/attachments/{att_id}", token)
    raw = att_data["data"].replace("-", "+").replace("_", "/")
    return base64.b64decode(raw + "==")

# ---- 5. PROCESS XLSX ----
def process_xlsx(xlsx_bytes):
    import pandas as pd

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        f.write(xlsx_bytes)
        tmp_path = f.name

    # Try the known sheet name first, fall back to first sheet
    try:
        df = pd.read_excel(tmp_path, sheet_name="PowerQueryresult", header=0)
    except Exception:
        df = pd.read_excel(tmp_path, sheet_name=0, header=0)

    print(f"  Columns: {df.columns.tolist()}")
    print(f"  Rows: {len(df)}")

    def clean(v):
        if v is None: return None
        if isinstance(v, float) and math.isnan(v): return None
        if hasattr(v, "isoformat"): return str(v)[:10]
        return str(v) if not isinstance(v, (int, float)) else v

    records = []
    for _, row in df.iterrows():
        r = {
            "Property":     clean(row.get("PropertyAbbrev")) or clean(row.get("Property")),
            "Unit":         clean(row.get("Unit")),
            "Status":       clean(row.get("Status")),
            "Priority":     clean(row.get("Priority")),
            "Type":         clean(row.get("Work Order Type")),
            "WONumber":     clean(row.get("Work Order Number")),
            "AssignedUser": clean(row.get("Assigned User")),
            "CreatedAt":    str(row.get("Created At", ""))[:10] if row.get("Created At") else None,
            "Description":  str(row.get("Service Request Description", "") or "")[:300].strip() or None,
            "URL":          clean(row.get("WorkOrderURLLink")),
        }
        records.append(r)

    os.unlink(tmp_path)
    return records

# ---- 6. BUILD DASHBOARD ----
def build_dashboard(records, date_str):
    with open("dashboard_template.html") as f:
        template = f.read()

    data_json = json.dumps(records)
    html = template.replace("DATA_PLACEHOLDER", data_json)
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
