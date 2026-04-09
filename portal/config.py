"""
config.py — all environment variable definitions for PowerUp Portal.
All IDs, paths, and endpoints live here. Nothing hardcoded elsewhere.

Locally:   values come from .env via python-dotenv
On Cloud:  app.py injects st.secrets into os.environ before this module loads
"""

import os


# ── Google Service Account ────────────────────────────────────
GOOGLE_SERVICE_ACCOUNT_JSON: str | None = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")

# ── M1 ────────────────────────────────────────────────────────
M1_APPS_SCRIPT_URL: str | None = os.environ.get("M1_APPS_SCRIPT_URL")

# ── M2 ────────────────────────────────────────────────────────
M2_TEMPLATE_ID: str | None = os.environ.get("M2_TEMPLATE_ID")
M2_RISK_REWARD_TEMPLATE_ID: str | None = os.environ.get("M2_RISK_REWARD_TEMPLATE_ID")

# ── M3 ────────────────────────────────────────────────────────
M3_TEMPLATE_ID: str | None = os.environ.get("M3_TEMPLATE_ID")

# ── Spreadsheets ──────────────────────────────────────────────
MAIN_SPREADSHEET_ID: str | None = os.environ.get("MAIN_SPREADSHEET_ID")
M3_SPREADSHEET_ID: str | None = os.environ.get("M3_SPREADSHEET_ID")
TIMESERIES_SPREADSHEET_ID: str | None = os.environ.get("TIMESERIES_SPREADSHEET_ID")
TIMESERIES_ROW_LIMIT: int = int(os.environ.get("TIMESERIES_ROW_LIMIT", "1000000"))
QUESTIONNAIRE_SPREADSHEET_ID: str | None = os.environ.get("QUESTIONNAIRE_SPREADSHEET_ID")

# ── M2 Assets on Drive ────────────────────────────────────────
M2_CATEGORIZATION_FILE_ID: str | None = os.environ.get("M2_CATEGORIZATION_FILE_ID")
M2_IMG_INFORM_ID: str | None = os.environ.get("M2_IMG_INFORM_ID")
M2_IMG_ONTRACK_ID: str | None = os.environ.get("M2_IMG_ONTRACK_ID")
M2_IMG_OFFTRACK_ID: str | None = os.environ.get("M2_IMG_OFFTRACK_ID")
M2_IMG_OUTOFFORM_ID: str | None = os.environ.get("M2_IMG_OUTOFFORM_ID")

# ── Drive Folders ─────────────────────────────────────────────
DRIVE_ROOT_FOLDER_ID: str | None = os.environ.get("DRIVE_ROOT_FOLDER_ID")
M1_OUTPUT_FOLDER_ID: str | None = os.environ.get("M1_OUTPUT_FOLDER_ID")
M2_OUTPUT_FOLDER_ID: str | None = os.environ.get("M2_OUTPUT_FOLDER_ID")
M3_OUTPUT_FOLDER_ID: str | None = os.environ.get("M3_OUTPUT_FOLDER_ID")

# ── Sheet tab names (single source of truth) ──────────────────
class MainSheets:
    PF_ID_MAPPING       = "PF_ID_Mapping"
    PF_LEVEL            = "PF_level"
    SCHEME_LEVEL        = "Scheme_level"
    RISKGROUP_LEVEL     = "Riskgroup_level"
    LINES               = "Lines"
    RESULTS             = "Results"
    INVESTED_VALUE_LINE = "Invested_Value_Line"
    SCHEME_CATEGORY     = "Scheme_Category"
    BASE_DATA           = "BASE_DATA"

class TimeSeriesSheets:
    LINES               = "Lines"
    INVESTED_VALUE_LINE = "Invested_Value_Line"

class M3Sheets:
    AUM             = "AUM"
    POWERRANKING    = "Powerranking"
    UPSIDE_DOWNSIDE = "Upside_Downside"
    ROLLING_RETURNS = "Rolling_Returns"

# ── Validation helper ─────────────────────────────────────────
def require(name: str, value: str | None) -> str:
    if not value:
        raise EnvironmentError(
            f"Required environment variable '{name}' is not set. "
            f"Add it to portal/.env or Streamlit Cloud secrets."
        )
    return value
