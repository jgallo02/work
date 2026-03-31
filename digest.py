import os
import requests
from datetime import datetime, timedelta, timezone

SMARTSHEET_API_TOKEN = os.environ["SMARTSHEET_API_TOKEN"]
SLACK_WEBHOOK_URL = os.environ["SLACK_WEBHOOK_URL"]
SHEET_ID = "6459127481296772"

def get_column_map(columns):
    return {col["title"]: col["id"] for col in columns}

def get_events():
    headers = {"Authorization": f"Bearer {SMARTSHEET_API_TOKEN}"}
    url = f"https://api.smartsheet.com/2.0/sheets/{SHEET_ID}"
    
    # Fetch full sheet — no row cap, unlike the MCP tool
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    sheet = response.json()

    col_map = get_column_map(sheet["columns"])
    start_date_col = col_map.get("Start Date")
    desc_col = col_map.get("Description")
    poc_col = col_map.get("Point of Contact")
    tentative_col = col_map.get("Tentative")
    external_col = col_map.get("External Announcement")
    geo_col = col_map.get("Geography")

    today = datetime.now(timezone.utc).date()
    two_weeks = today + timedelta(days=14)

    events = []
    for row in sheet["rows"]:
        cells = {cell["columnId"]: cell for cell in row.get("cells", [])}

        start_raw = cells.get(start_date_col, {}).get("value")
        if not start_raw:
            continue

        try:
            start_date = datetime.strptime(start_raw[:10], "%Y-%m-%d").date()
        except ValueError:
            continue

        if not (today <= start_date <= two_weeks):
            continue

        description = cells.get(desc_col, {}).get("value", "")
        # Skip week-header rows (they're just labels like "Week of Jan. 5...")
        if not description or str(description).startswith("Week of"):
            continue

        poc = cells.get(poc_col, {}).get("displayValue") or cells.get(poc_col, {}).get("value", "")
        tentative = cells.get(tentative_col, {}).get("value", False)
        external = cells.get(external_col, {}).get("value", False)
        geo = cells.get(geo_col, {}).get("displayValue") or ""

        events.append({
            "date": start_date,
            "description": description,
            "poc": poc or "—",
            "tentative": bool(tentative),
            "external": bool(external),
            "geo": geo or "",
        })

    events.sort(key=lambda e: e["date"])
    return events

def format_slack_message(events):
    today = datetime.now(timezone.utc).date()
    two_weeks = today + timedelta(days=14)
    date_range = f"{today.strftime('%b %d')} – {two_weeks.strftime('%b %d, %Y')}"

    if not events:
        return f"📅 *Corporate Affairs Calendar — Next 2 Weeks*\n_{date_range}_\n\nNo events found for this period."

    # Group by date
