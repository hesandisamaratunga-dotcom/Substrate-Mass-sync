"""
Substrate mass sync (v2) - reads today's row from the 'GitHub' tab of
2025-11-19_Auftragsplanung.xlsx via Microsoft Graph, pushes the mass value
to a new field on the existing Datacake device.

Required environment variables (set as GitHub Actions secrets):
  AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
  SHAREPOINT_SITE_ID, SHAREPOINT_DRIVE_ID, SHAREPOINT_FILE_PATH
  DATACAKE_API_TOKEN, DATACAKE_DEVICE_SERIAL, DATACAKE_FIELD_KEY
"""

import os
import sys
from datetime import datetime
from urllib.parse import quote
import requests

TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]

SITE_ID = os.environ["SHAREPOINT_SITE_ID"]
DRIVE_ID = os.environ["SHAREPOINT_DRIVE_ID"]
FILE_PATH = os.environ["SHAREPOINT_FILE_PATH"]  # e.g. "/2025-11-19_Auftragsplanung.xlsx"
WORKSHEET_NAME = "GitHub"

DATACAKE_API_TOKEN = os.environ["DATACAKE_API_TOKEN"]
DATACAKE_DEVICE_SERIAL = os.environ["DATACAKE_DEVICE_SERIAL"]
DATACAKE_FIELD_KEY = os.environ.get("DATACAKE_FIELD_KEY", "SUBSTRATE_MASS_2")

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
    # Using the item ID directly (instead of the combined root:/path:/workbook
    # syntax) works around the "Could not obtain a WAC access token" error
    # that the path-based form triggers even with full permissions.
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


def find_todays_mass(used_range: dict):
    text_rows = used_range["text"]  # displayed strings, e.g. "28.01.2026"
    today_str = datetime.now().strftime("%d.%m.%Y")  # matches sheet's DD.MM.YYYY format

    for row in text_rows[1:]:  # row 0 is the header (Datum / Mass)
        date_cell = row[0].strip() if len(row) > 0 else ""
        mass_cell = row[1].strip() if len(row) > 1 else ""

        if date_cell == today_str:
            if not mass_cell or mass_cell == "0":
                return None  # date matched, but no mass recorded yet
            # sheet uses German locale: '.' thousands sep, ',' decimal sep
            return float(mass_cell.replace(".", "").replace(",", "."))

    return None  # today's date row not present yet


def push_to_datacake(mass: float) -> None:
    # Matches the original working pipeline's push logic exactly:
    # POST to the /record/ endpoint with a batch list of {field, value} dicts.
    url = f"https://api.datacake.co/v1/devices/{DATACAKE_DEVICE_SERIAL}/record/?batch=true"
    headers = {
        "Authorization": f"Token {DATACAKE_API_TOKEN}",
        "Content-Type": "application/json",
    }
    payload = [{"field": DATACAKE_FIELD_KEY, "value": mass}]
    resp = requests.post(url, headers=headers, json=payload, timeout=30)
    if not resp.ok:
        print(f"Datacake API error {resp.status_code}: {resp.text}")
    resp.raise_for_status()
    print(f"Pushed {DATACAKE_FIELD_KEY} = {mass} to Datacake:", resp.json())


def main():
    token = get_graph_token()
    item_id = get_item_id(token)
    used_range = get_used_range(token, item_id)
    mass = find_todays_mass(used_range)

    if mass is None:
        print("No mass value found for today - skipping push.")
        sys.exit(0)

    push_to_datacake(mass)


if __name__ == "__main__":
    main()
