"""
config.py — all environment variable definitions for PowerUp Portal.
All IDs, paths, and endpoints live here. Nothing hardcoded elsewhere.

Resolution order for each variable:
  1. os.environ / .env file  (local development)
  2. st.secrets[key]         (Streamlit Cloud)
"""

import os
from dotenv import load_dotenv

load_dotenv()


def _get(key: str, default: str | None = None) -> str | None:
    """Read a config value from env vars first, then st.secrets."""
    # 1. Environment variable (set by .env locally, or by system)
    val = os.environ.get(key)
    if val:
        return val
    # 2. Streamlit secrets (set via Streamlit Cloud dashboard)
    try:
        import streamlit as st
        val = st.secrets[key]
        return str(val)
    except (KeyError, AttributeError, FileNotFoundError):
        pass
    return default


# ── Google Service Account ────────────────────────────────────
GOOGLE_SERVICE_ACCOUNT_JSON: str | None = _get("GOOGLE_SERVICE_ACCOUNT_JSON")

# ── M1 ────────────────────────────────────────────────────────
M1_APPS_SCRIPT_URL: str | None = _get("M1_APPS_SCRIPT_URL")

# ── M2 ────────────────────────────────────────────────────────
M2_TEMPLATE_ID: str | None = _get("M2_TEMPLATE_ID")
M2_RISK_REWARD_TEMPLATE_ID: str | None = _get("M2_RISK_REWARD_TEMPLATE_ID")

# ── M3 ────────────────────────────────────────────────────────
M3_TEMPLATE_ID: str | None = _get("M3_TEMPLATE_ID")

# ── Spreadsheets ──────────────────────────────────────────────
MAIN_SPREADSHEET_ID: str | None = _get("MAIN_SPREADSHEET_ID")
M3_SPREADSHEET_ID: str | None = _get("M3_SPREADSHEET_ID")
TIMESERIES_SPREADSHEET_ID: str | None = _get("TIMESERIES_SPREADSHEET_ID")
TIMESERIES_ROW_LIMIT: int = int(_get("TIMESERIES_ROW_LIMIT", "1000000"))
QUESTIONNAIRE_SPREADSHEET_ID: str | None = _get("QUESTIONNAIRE_SPREADSHEET_ID")

# ── M2 Assets on Drive ────────────────────────────────────────
M2_CATEGORIZATION_FILE_ID: str | None = _get("M2_CATEGORIZATION_FILE_ID")
M2_IMG_INFORM_ID: str | None = _get("M2_IMG_INFORM_ID")
M2_IMG_ONTRACK_ID: str | None = _get("M2_IMG_ONTRACK_ID")
M2_IMG_OFFTRACK_ID: str | None = _get("M2_IMG_OFFTRACK_ID")
M2_IMG_OUTOFFORM_ID: str | None = _get("M2_IMG_OUTOFFORM_ID")

# ── Drive Folders ─────────────────────────────────────────────
DRIVE_ROOT_FOLDER_ID: str | None = _get("DRIVE_ROOT_FOLDER_ID")
M1_OUTPUT_FOLDER_ID: str | None = _get("M1_OUTPUT_FOLDER_ID")
M2_OUTPUT_FOLDER_ID: str | None = _get("M2_OUTPUT_FOLDER_ID")
M3_OUTPUT_FOLDER_ID: str | None = _get("M3_OUTPUT_FOLDER_ID")

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
