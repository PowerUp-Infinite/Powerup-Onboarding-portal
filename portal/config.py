"""
config.py — all environment variable definitions for PowerUp Portal.
All IDs, paths, and endpoints live here. Nothing hardcoded elsewhere.

Values are read lazily from os.environ on first access, NOT at import time.
This ensures st.secrets → os.environ injection in app.py runs first.

Locally:   values come from .env via python-dotenv
On Cloud:  app.py injects st.secrets into os.environ at startup
"""

import os

# ── Env-var key mapping ───────────────────────────────────────
# Maps config attribute name → (env var name, default value)
_ENV_KEYS: dict[str, tuple[str, str | None]] = {
    "GOOGLE_SERVICE_ACCOUNT_JSON": ("GOOGLE_SERVICE_ACCOUNT_JSON", None),
    "M1_APPS_SCRIPT_URL":         ("M1_APPS_SCRIPT_URL", None),
    "M2_TEMPLATE_ID":             ("M2_TEMPLATE_ID", None),
    "M2_RISK_REWARD_TEMPLATE_ID": ("M2_RISK_REWARD_TEMPLATE_ID", None),
    "M3_TEMPLATE_ID":             ("M3_TEMPLATE_ID", None),
    "MAIN_SPREADSHEET_ID":        ("MAIN_SPREADSHEET_ID", None),
    "M3_SPREADSHEET_ID":          ("M3_SPREADSHEET_ID", None),
    "TIMESERIES_SPREADSHEET_ID":  ("TIMESERIES_SPREADSHEET_ID", None),
    "TIMESERIES_ROW_LIMIT":       ("TIMESERIES_ROW_LIMIT", "1000000"),
    "QUESTIONNAIRE_SPREADSHEET_ID": ("QUESTIONNAIRE_SPREADSHEET_ID", None),
    "M2_CATEGORIZATION_FILE_ID":  ("M2_CATEGORIZATION_FILE_ID", None),
    "M2_IMG_INFORM_ID":           ("M2_IMG_INFORM_ID", None),
    "M2_IMG_ONTRACK_ID":          ("M2_IMG_ONTRACK_ID", None),
    "M2_IMG_OFFTRACK_ID":         ("M2_IMG_OFFTRACK_ID", None),
    "M2_IMG_OUTOFFORM_ID":        ("M2_IMG_OUTOFFORM_ID", None),
    "DRIVE_ROOT_FOLDER_ID":       ("DRIVE_ROOT_FOLDER_ID", None),
    "M1_OUTPUT_FOLDER_ID":        ("M1_OUTPUT_FOLDER_ID", None),
    "M2_OUTPUT_FOLDER_ID":        ("M2_OUTPUT_FOLDER_ID", None),
    "M3_OUTPUT_FOLDER_ID":        ("M3_OUTPUT_FOLDER_ID", None),
    # Agreement Automation
    "AGREEMENT_ELITE_TEMPLATE_ID":    ("AGREEMENT_ELITE_TEMPLATE_ID", None),
    "AGREEMENT_NONELITE_TEMPLATE_ID": ("AGREEMENT_NONELITE_TEMPLATE_ID", None),
    "AGREEMENT_OUTPUT_FOLDER_ID":     ("AGREEMENT_OUTPUT_FOLDER_ID", None),
}


def __getattr__(name: str):
    """Lazy lookup: read os.environ when the attribute is first accessed."""
    if name in _ENV_KEYS:
        env_key, default = _ENV_KEYS[name]
        val = os.environ.get(env_key, default)
        if name == "TIMESERIES_ROW_LIMIT":
            return int(val) if val else 1000000
        return val
    raise AttributeError(f"module 'config' has no attribute {name!r}")


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
    RATINGS             = "Ratings"

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
