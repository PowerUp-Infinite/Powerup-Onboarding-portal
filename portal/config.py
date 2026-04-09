"""
config.py — all environment variable definitions for PowerUp Portal.
All IDs, paths, and endpoints live here. Nothing hardcoded elsewhere.

Resolution order:
  1. .env file           (local development — via python-dotenv)
  2. st.secrets          (Streamlit Cloud — injected into os.environ at startup)
"""

import os
from dotenv import load_dotenv

load_dotenv()

# On Streamlit Cloud there is no .env file. Secrets are only accessible
# via st.secrets, NOT os.environ.  Bridge the gap by copying every
# top-level secret into os.environ so the rest of this file (and any
# other module that reads os.getenv) just works.
try:
    import streamlit as st
    for key, val in st.secrets.items():
        if isinstance(val, str) and key not in os.environ:
            os.environ[key] = val
except Exception:
    pass  # not running under Streamlit, or no secrets configured


def _env(key: str, default: str | None = None) -> str | None:
    return os.environ.get(key, default)


# ── Google Service Account ────────────────────────────────────
GOOGLE_SERVICE_ACCOUNT_JSON: str | None = _env("GOOGLE_SERVICE_ACCOUNT_JSON")

# ── M1 ────────────────────────────────────────────────────────
M1_APPS_SCRIPT_URL: str | None = _env("M1_APPS_SCRIPT_URL")

# ── M2 ────────────────────────────────────────────────────────
M2_TEMPLATE_ID: str | None = _env("M2_TEMPLATE_ID")
M2_RISK_REWARD_TEMPLATE_ID: str | None = _env("M2_RISK_REWARD_TEMPLATE_ID")

# ── M3 ────────────────────────────────────────────────────────
M3_TEMPLATE_ID: str | None = _env("M3_TEMPLATE_ID")

# ── Spreadsheets ──────────────────────────────────────────────
MAIN_SPREADSHEET_ID: str | None = _env("MAIN_SPREADSHEET_ID")
M3_SPREADSHEET_ID: str | None = _env("M3_SPREADSHEET_ID")
TIMESERIES_SPREADSHEET_ID: str | None = _env("TIMESERIES_SPREADSHEET_ID")
TIMESERIES_ROW_LIMIT: int = int(_env("TIMESERIES_ROW_LIMIT", "1000000"))
QUESTIONNAIRE_SPREADSHEET_ID: str | None = _env("QUESTIONNAIRE_SPREADSHEET_ID")

# ── M2 Assets on Drive ────────────────────────────────────────
M2_CATEGORIZATION_FILE_ID: str | None = _env("M2_CATEGORIZATION_FILE_ID")
M2_IMG_INFORM_ID: str | None = _env("M2_IMG_INFORM_ID")
M2_IMG_ONTRACK_ID: str | None = _env("M2_IMG_ONTRACK_ID")
M2_IMG_OFFTRACK_ID: str | None = _env("M2_IMG_OFFTRACK_ID")
M2_IMG_OUTOFFORM_ID: str | None = _env("M2_IMG_OUTOFFORM_ID")

# ── Drive Folders ─────────────────────────────────────────────
DRIVE_ROOT_FOLDER_ID: str | None = _env("DRIVE_ROOT_FOLDER_ID")
M1_OUTPUT_FOLDER_ID: str | None = _env("M1_OUTPUT_FOLDER_ID")
M2_OUTPUT_FOLDER_ID: str | None = _env("M2_OUTPUT_FOLDER_ID")
M3_OUTPUT_FOLDER_ID: str | None = _env("M3_OUTPUT_FOLDER_ID")

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
