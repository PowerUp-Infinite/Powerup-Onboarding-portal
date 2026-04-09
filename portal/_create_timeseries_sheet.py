"""
_create_timeseries_sheet.py — run once to create the Time Series spreadsheet
in Drive and print its ID. ID is then written to .env automatically.

    cd "c:/PowerUpInfinite/Pre-Onboarding Portal"
    python portal/_create_timeseries_sheet.py
"""
import os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(__file__))
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

from google.oauth2 import service_account
from googleapiclient.discovery import build
import json

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def _creds():
    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if os.path.exists(raw):
        return service_account.Credentials.from_service_account_file(raw, scopes=SCOPES)
    return service_account.Credentials.from_service_account_info(json.loads(raw), scopes=SCOPES)

def main():
    creds  = _creds()
    sheets = build("sheets", "v4", credentials=creds)
    drive  = build("drive",  "v3", credentials=creds)

    root_folder = os.getenv("DRIVE_ROOT_FOLDER_ID")
    data_folder = None

    # Find or use the Data subfolder inside root
    if root_folder:
        res = drive.files().list(
            q=f"'{root_folder}' in parents and mimeType='application/vnd.google-apps.folder' and name='Data' and trashed=false",
            fields="files(id,name)"
        ).execute()
        if res.get("files"):
            data_folder = res["files"][0]["id"]

    # Create the spreadsheet
    body = {
        "properties": {"title": "PowerUp Portal — Time Series"},
        "sheets": [
            {"properties": {"title": "Lines"}},
            {"properties": {"title": "Invested_Value_Line"}},
        ]
    }
    ss = sheets.spreadsheets().create(body=body).execute()
    ss_id = ss["spreadsheetId"]
    print(f"Created spreadsheet: {ss_id}")

    # Write header rows
    sheets.spreadsheets().values().update(
        spreadsheetId=ss_id, range="Lines!A1",
        valueInputOption="RAW",
        body={"values": [["DATE", "PF_ID", "TYPE", "CURRENT_VALUE"]]}
    ).execute()
    sheets.spreadsheets().values().update(
        spreadsheetId=ss_id, range="Invested_Value_Line!A1",
        valueInputOption="RAW",
        body={"values": [["PF_ID", "DATE", "INVESTED_AMOUNT"]]}
    ).execute()
    print("Headers written.")

    # Move into Data folder if found
    if data_folder:
        f = drive.files().get(fileId=ss_id, fields="parents").execute()
        prev = ",".join(f.get("parents", []))
        drive.files().update(
            fileId=ss_id,
            addParents=data_folder,
            removeParents=prev,
            fields="id,parents"
        ).execute()
        print(f"Moved into Data folder ({data_folder})")

    # Append to .env
    env_path = os.path.join(os.path.dirname(__file__), ".env")
    with open(env_path, "a") as f:
        f.write(f"\nTIMESERIES_SPREADSHEET_ID={ss_id}\n")
    print(f"Written TIMESERIES_SPREADSHEET_ID to .env")
    print(f"\nDone. Spreadsheet ID: {ss_id}")

if __name__ == "__main__":
    main()
