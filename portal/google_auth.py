"""
google_auth.py — service account authentication for Google APIs.

Builds and caches API clients for:
  - Google Sheets  (spreadsheets read/write)
  - Google Drive   (file copy, move, share)
  - Google Slides  (presentation populate)

Credential resolution order:
  1. st.secrets["gcp_service_account"]  — Streamlit Cloud / production
  2. GOOGLE_SERVICE_ACCOUNT_JSON env var as a file path  — local JSON key file
  3. GOOGLE_SERVICE_ACCOUNT_JSON env var as inline JSON  — CI / Docker env

Usage:
    from google_auth import get_sheets_service, get_drive_service, get_slides_service

    sheets = get_sheets_service()
    result = sheets.spreadsheets().values().get(
        spreadsheetId="...", range="Sheet1!A1:Z"
    ).execute()
"""

import json
import os

import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build

# All scopes needed across the three automations.
_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations",
]


def _load_credentials() -> service_account.Credentials:
    """
    Load service account credentials using the resolution order above.
    Raises ValueError with a helpful message if nothing is configured.
    """

    # 1. st.secrets (Streamlit Cloud / production)
    try:
        info = dict(st.secrets["gcp_service_account"])
        return service_account.Credentials.from_service_account_info(
            info, scopes=_SCOPES
        )
    except (KeyError, FileNotFoundError):
        pass

    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        raise ValueError(
            "No Google credentials found.\n"
            "Options:\n"
            "  • Set st.secrets['gcp_service_account'] in .streamlit/secrets.toml\n"
            "  • Set GOOGLE_SERVICE_ACCOUNT_JSON=/path/to/key.json in .env\n"
            "  • Set GOOGLE_SERVICE_ACCOUNT_JSON='{...inline JSON...}' in .env"
        )

    # 2. File path
    if os.path.exists(raw):
        return service_account.Credentials.from_service_account_file(
            raw, scopes=_SCOPES
        )

    # 3. Inline JSON string
    try:
        info = json.loads(raw)
        return service_account.Credentials.from_service_account_info(
            info, scopes=_SCOPES
        )
    except json.JSONDecodeError:
        raise ValueError(
            f"GOOGLE_SERVICE_ACCOUNT_JSON is set but is neither a valid file path "
            f"nor valid JSON. Value starts with: {raw[:60]!r}"
        )


# Cache resource so clients are built once per Streamlit session,
# not on every rerun.
@st.cache_resource(show_spinner=False)
def get_sheets_service():
    """Return an authenticated Google Sheets API client (v4)."""
    return build("sheets", "v4", credentials=_load_credentials())


@st.cache_resource(show_spinner=False)
def get_drive_service():
    """Return an authenticated Google Drive API client (v3)."""
    return build("drive", "v3", credentials=_load_credentials())


@st.cache_resource(show_spinner=False)
def get_slides_service():
    """Return an authenticated Google Slides API client (v1)."""
    return build("slides", "v1", credentials=_load_credentials())


def check_auth() -> tuple[bool, str]:
    """
    Lightweight connectivity check — tries to load credentials.
    Returns (ok: bool, message: str).
    Use in the sidebar or on startup to surface config errors early.
    """
    try:
        _load_credentials()
        return True, "Google credentials loaded successfully."
    except ValueError as e:
        return False, str(e)
