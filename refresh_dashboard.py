#!/usr/bin/env python3
"""
Work Order Dashboard - Auto Refresh Script
Runs daily via GitHub Actions:
1. Finds today's email from pmassistant@thesitusgroup.com
2. Downloads the XLSX attachment
3. Processes the data
4. Injects it into the dashboard HTML template
5. GitHub Actions commits and deploys to GitHub Pages
"""

import os, base64, json, math, re
from datetime import datetime, timezone
import urllib.request, urllib.parse

# ---- CONFIG ----
GMAIL_TOKEN   = os.environ["GMAIL_TOKEN"]       # OAuth token (set in GitHub secrets)
SENDER        = "donotreply@appfolio.com"
SUBJECT_MATCH = "Work Order Automation"

# ---- 1. FIND TODAY'S EMAIL ----
def gmail_request(path, token):
    req = urllib.request.Request(
        f"https://gmail.googleapis.com/gmail/v1/users/me/{path}",
        headers={"Authorization": f"Bearer {token}"}
    )
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())

def find_todays_message():
    today = datetime.now(timezone.utc).strftime("%Y/%m/%d")
    q = urllib.parse.quote(f'from:{SENDER} subject:"{SUBJECT_MATCH}" after:{today} has:attachment')
    data = gmail_request(f"messages?q={q}&maxResults=1", GMAIL_TOKEN)
    messages = data.get("messages", [])
    if not messages:
        raise Exception(f"No email found from {SENDER} today ({today})")
    return messages[0]["id"]

# ---- 2. GET XLSX ATTACHMENT ----
def get_xlsx_attachment(message_id):
    msg = gmail_request(f"messages/{message_id}?format=full", GMAIL_TOKEN)
    
    def find_xlsx(parts):
        for part in parts:
            if part.get("mimeType") == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                att_id = part["body"].get("attachmentId")
                if att_id:
                    return att_id
            if "parts" in part:
                result = find_xlsx(part["parts"])
                if result:
                    return result
        return None
    
    parts = msg.get("payload", {}).get("parts", [])
    att_id = find_xlsx(parts)
    if not att_id:
        raise Exception("No XLSX attachment found in email")
    
    att_data = gmail_request(f"messages/{message_id}/attachments/{att_id}", GMAIL_TOKEN)
    # Gmail uses URL-safe base64
    raw = att_data["data"].replace("-", "+").replace("_", "/")
    return base64.b64decode(raw + "==")

# ---- 3. PROCESS XLSX ----
def process_xlsx(xlsx_bytes):
    import tempfile, pandas as pd
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        f.write(xlsx_bytes)
        tmp_path = f.name
    
    df = pd.read_excel(tmp_path, sheet_name="PowerQueryresult", header=0)
    
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

# ---- 4. BUILD DASHBOARD ----
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
    print("🔍 Finding today's email...")
    msg_id = find_todays_message()
    print(f"✓ Found message: {msg_id}")
    
    print("📎 Downloading XLSX attachment...")
    xlsx_bytes = get_xlsx_attachment(msg_id)
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
