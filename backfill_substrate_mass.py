"""
One-off backfill: walks every dated row in the 'GitHub' tab and pushes each
mass value to Datacake with its correct historical timestamp (instead of
'now'). Separate from fetch_substrate_mass.py - run this manually whenever
you need to (re)populate history; it does not run on a schedule.

Required environment variables (same secrets as the daily sync):
  AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
  SHAREPOINT_DRIVE_ID, SHAREPOINT_FILE_PATH
  DATACAKE_API_TOKEN, DATACAKE_DEVICE_SERIAL, DATACAKE_FIELD_KEY
"""

import os
from datetime import datetime, timezone
from urllib.parse import quote
import requests

TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

DRIVE_ID = os.environ["SHAREPOINT_DRIVE_ID"]
FILE_PATH = os.environ["SHAREPOINT_FILE_PATH"]
WORKSHEET_NAME = "GitHub"

DATACAKE_API_TOKEN = os.environ["DATACAKE_API_TOKEN"]
DATACAKE_DEVICE_SERIAL = os.environ["DATACAKE_DEVICE_SERIAL"]  # the device UUID from the dashboard URL
DATACAKE_FIELD_KEY = os.environ["DATACAKE_FIELD_KEY"]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def get_graph_token() -> str:
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    resp = requests.post(url, data=data, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]


def get_item_id(token: str) -> str:
    path = FILE_PATH if FILE_PATH.startswith("/") else f"/{FILE_PATH}"
    encoded_path = quote(path, safe="/")
    url = f"{GRAPH_BASE}/drives/{DRIVE_ID}/root:{encoded_path}?$select=id"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, timeout=30)
    if not resp.ok:
        print(f"Graph API error {resp.status_code}: {resp.text}")
    resp.raise_for_status()
    return resp.json()["id"]


def get_used_range(token: str, item_id: str) -> dict:
    url = (
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{item_id}"
        f"/workbook/worksheets('{WORKSHEET_NAME}')/usedRange"
        f"?$select=text"
    )
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, timeout=30)
    if not resp.ok:
        print(f"Graph API error {resp.status_code}: {resp.text}")
    resp.raise_for_status()
    return resp.json()


def parse_mass(mass_cell: str):
    mass_cell = mass_cell.strip()
    if not mass_cell or mass_cell == "0":
        return None
    try:
        return float(mass_cell.replace(".", "").replace(",", "."))
    except ValueError:
        return None  # non-numeric cell (e.g. a label like "Inkubation")


def build_backfill_entries(used_range: dict):
    text_rows = used_range["text"]
    entries = []
    skipped = 0
    today = datetime.now(timezone.utc).date()

    for row in text_rows[1:]:  # skip header row
        date_cell = row[0].strip() if len(row) > 0 else ""
        mass_cell = row[1] if len(row) > 1 else ""

        try:
            date_obj = datetime.strptime(date_cell, "%d.%m.%Y").replace(tzinfo=timezone.utc)
        except ValueError:
            skipped += 1
            continue  # not a real date row (placeholder "0", blank, label, etc.)

        if date_obj.date() > today:
            skipped += 1
            continue  # forward-planned row - not a real measurement yet, don't backfill it

        mass = parse_mass(mass_cell)
        if mass is None:
            skipped += 1
            continue  # no usable mass value for this date

        entries.append({
            "field": DATACAKE_FIELD_KEY,
            "value": mass,
            "timestamp": str(int(date_obj.timestamp())),
        })

    return entries, skipped


def push_batch(entries: list) -> None:
    url = f"https://api.datacake.co/v1/devices/{DATACAKE_DEVICE_SERIAL}/record/?batch=true"
    headers = {
        "Authorization": f"Token {DATACAKE_API_TOKEN}",
        "Content-Type": "application/json",
    }
    resp = requests.post(url, headers=headers, json=entries, timeout=60)
    if not resp.ok:
        print(f"Datacake API error {resp.status_code}: {resp.text}")
    resp.raise_for_status()
    print("Datacake response:", resp.json())


def main():
    token = get_graph_token()
    item_id = get_item_id(token)
    used_range = get_used_range(token, item_id)

    entries, skipped = build_backfill_entries(used_range)
    print(f"Found {len(entries)} historical rows to backfill, skipped {skipped} rows.")

    if not entries:
        print("Nothing to backfill.")
        return

    for e in entries:
        readable = datetime.fromtimestamp(int(e["timestamp"]), tz=timezone.utc).strftime("%d.%m.%Y")
        print(f"  {readable}: {e['value']}")

    push_batch(entries)
    print(f"Backfilled {len(entries)} data points to {DATACAKE_FIELD_KEY}.")


if __name__ == "__main__":
    main()
