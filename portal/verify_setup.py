"""
verify_setup.py — run this once to confirm auth works and all sheets are accessible.

    cd "c:/PowerUpInfinite/Pre-Onboarding Portal"
    python portal/verify_setup.py
"""

import os, sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(__file__))

from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

from config import (
    MAIN_SPREADSHEET_ID, M3_SPREADSHEET_ID, QUESTIONNAIRE_SPREADSHEET_ID,
    M2_TEMPLATE_ID, M2_RISK_REWARD_TEMPLATE_ID, M3_TEMPLATE_ID,
    M1_OUTPUT_FOLDER_ID, M2_OUTPUT_FOLDER_ID, M3_OUTPUT_FOLDER_ID,
    DRIVE_ROOT_FOLDER_ID, MainSheets, M3Sheets
)

OK  = "  ✓"
ERR = "  ✗"

def _sheets():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    import json

    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if os.path.exists(raw):
        creds = service_account.Credentials.from_service_account_file(
            raw, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly",
                         "https://www.googleapis.com/auth/drive.readonly"])
    else:
        creds = service_account.Credentials.from_service_account_info(
            json.loads(raw),
            scopes=["https://www.googleapis.com/auth/spreadsheets.readonly",
                    "https://www.googleapis.com/auth/drive.readonly"])
    sheets = build("sheets", "v4", credentials=creds)
    drive  = build("drive",  "v3", credentials=creds)
    return sheets, drive


def check_sheet_tabs(svc, spreadsheet_id, label, expected_tabs):
    print(f"\n{label}  (id: {spreadsheet_id})")
    try:
        meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        actual = {s["properties"]["title"] for s in meta["sheets"]}
        for tab in expected_tabs:
            if tab in actual:
                print(f"{OK}  tab '{tab}' found")
            else:
                print(f"{ERR}  tab '{tab}' MISSING")
        extra = actual - set(expected_tabs)
        if extra:
            print(f"     (extra tabs present: {', '.join(sorted(extra))})")
    except Exception as e:
        print(f"{ERR}  Cannot access spreadsheet: {e}")


def check_headers(svc, spreadsheet_id, tab, expected_headers):
    try:
        res = svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"{tab}!1:1"
        ).execute()
        actual = [str(h).strip() for h in res.get("values", [[]])[0]] if res.get("values") else []
        missing = [h for h in expected_headers if h not in actual]
        if not actual:
            print(f"     {tab}: no header row found yet (sheet is empty — OK if not migrated yet)")
        elif missing:
            print(f"     {tab}: missing headers — {missing}")
        else:
            print(f"{OK}  {tab}: all {len(expected_headers)} headers present")
    except Exception as e:
        print(f"{ERR}  {tab}: {e}")


def check_drive_item(drv, item_id, label):
    try:
        f = drv.files().get(fileId=item_id, fields="name,mimeType").execute()
        print(f"{OK}  {label}: '{f['name']}' ({f['mimeType'].split('.')[-1]})")
    except Exception as e:
        print(f"{ERR}  {label}: {e}")


def main():
    print("=" * 60)
    print("PowerUp Portal — Setup Verification")
    print("=" * 60)

    print("\n[1] Loading credentials...")
    try:
        svc, drv = _sheets()
        print(f"{OK}  Service account authenticated")
    except Exception as e:
        print(f"{ERR}  Auth failed: {e}")
        sys.exit(1)

    # ── Main spreadsheet ──────────────────────────────────────
    expected_main = [
        MainSheets.PF_LEVEL, MainSheets.SCHEME_LEVEL, MainSheets.RISKGROUP_LEVEL,
        MainSheets.LINES, MainSheets.RESULTS, MainSheets.INVESTED_VALUE_LINE,
        MainSheets.SCHEME_CATEGORY, MainSheets.BASE_DATA,
    ]
    check_sheet_tabs(svc, MAIN_SPREADSHEET_ID, "[2] Main Spreadsheet", expected_main)

    MAIN_HEADERS = {
        MainSheets.PF_LEVEL: ["PF_ID","PF_XIRR","BM_XIRR","PF_CURRENT_VALUE","INVESTED_VALUE"],
        MainSheets.SCHEME_LEVEL: ["PF_ID","ISIN","FUND_NAME","RISK_GROUP_L0","CURRENT_VALUE","XIRR_VALUE"],
        MainSheets.RISKGROUP_LEVEL: ["PF_ID","RISK_GROUP_L0","CURRENT_VALUE"],
        MainSheets.LINES: ["DATE","PF_ID","TYPE","CURRENT_VALUE"],
        MainSheets.RESULTS: ["PF_ID","TYPE","XIRR","CURRENT_VALUE"],
        MainSheets.INVESTED_VALUE_LINE: ["PF_ID","DATE","INVESTED_AMOUNT"],
        MainSheets.SCHEME_CATEGORY: ["Powerup Broad Category","Proposed Sub-Category"],
        MainSheets.BASE_DATA: ["ISIN","EXPENSE_RATIO"],
    }
    print("\n  Checking headers (spot-check key columns):")
    for tab, hdrs in MAIN_HEADERS.items():
        check_headers(svc, MAIN_SPREADSHEET_ID, tab, hdrs)

    # ── M3 Reference spreadsheet ──────────────────────────────
    expected_m3 = [M3Sheets.AUM, M3Sheets.POWERRANKING,
                   M3Sheets.UPSIDE_DOWNSIDE, M3Sheets.ROLLING_RETURNS]
    check_sheet_tabs(svc, M3_SPREADSHEET_ID, "[3] M3 Reference Spreadsheet", expected_m3)

    M3_HEADERS = {
        M3Sheets.AUM: ["ISIN","FUND_NAME","AUM"],
        M3Sheets.POWERRANKING: ["ISIN","POWERRANK","POWERRATING"],
        M3Sheets.UPSIDE_DOWNSIDE: ["Scheme ISIN","Downside Capture Ratio","Upside Capture Ratio"],
        M3Sheets.ROLLING_RETURNS: ["ENTITYID","RETURN_VALUE","ROLLING_PERIOD"],
    }
    print("\n  Checking headers (spot-check key columns):")
    for tab, hdrs in M3_HEADERS.items():
        check_headers(svc, M3_SPREADSHEET_ID, tab, hdrs)

    # ── Questionnaire spreadsheet ─────────────────────────────
    print(f"\n[4] Questionnaire Spreadsheet  (id: {QUESTIONNAIRE_SPREADSHEET_ID})")
    try:
        meta = svc.spreadsheets().get(spreadsheetId=QUESTIONNAIRE_SPREADSHEET_ID).execute()
        tabs = [s["properties"]["title"] for s in meta["sheets"]]
        print(f"{OK}  Accessible — tabs: {tabs}")
    except Exception as e:
        print(f"{ERR}  Cannot access: {e}")

    # ── Drive items ───────────────────────────────────────────
    print("\n[5] Drive Templates & Folders")
    check_drive_item(drv, M2_TEMPLATE_ID,           "M2 Base Deck template")
    check_drive_item(drv, M2_RISK_REWARD_TEMPLATE_ID,"M2 Risk/Reward template")
    check_drive_item(drv, M3_TEMPLATE_ID,            "M3 Template")
    check_drive_item(drv, M1_OUTPUT_FOLDER_ID,       "M1 Output folder")
    check_drive_item(drv, M2_OUTPUT_FOLDER_ID,       "M2 Output folder")
    check_drive_item(drv, M3_OUTPUT_FOLDER_ID,       "M3 Output folder")
    check_drive_item(drv, DRIVE_ROOT_FOLDER_ID,      "Root PowerUp Portal folder")

    print("\n" + "=" * 60)
    print("Done. Fix any ✗ items above before proceeding to Step 2.")
    print("=" * 60)


if __name__ == "__main__":
    main()
