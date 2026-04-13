"""
desktop/portal_shims/google_auth.py

Drop-in replacement for portal/google_auth.py that does NOT import streamlit.

The desktop app is a standalone native app — it has nothing to do with
Streamlit Cloud, so pulling 50+ MB of Streamlit + all its deps into the
PyInstaller bundle is pure waste. Worse, PyInstaller drops streamlit's
package metadata, so the import itself crashes with
"No package metadata was found for streamlit" at runtime.

app_config.py inserts desktop/portal_shims onto sys.path BEFORE portal/,
so when portal/sheets.py does `from google_auth import get_sheets_service`
Python resolves to THIS module, not portal/google_auth.py.

Public API is identical to the cloud version:
  get_sheets_service()  — cached build('sheets', 'v4', ...)
  get_drive_service()   — cached build('drive',  'v3', ...)
  get_slides_service()  — cached build('slides', 'v1', ...)
  check_auth()          — (ok: bool, msg: str)
"""
from __future__ import annotations

import json
import os
from functools import lru_cache

from google.oauth2 import service_account
from googleapiclient.discovery import build


_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations",
]


def _load_credentials() -> service_account.Credentials:
    """Load credentials from GOOGLE_SERVICE_ACCOUNT_JSON (set by app_config.py
    to point at the bundled resources/credentials.json). Accepts either a
    file path or raw inline JSON."""
    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        raise ValueError(
            "No Google credentials found. app_config.py should have pointed "
            "GOOGLE_SERVICE_ACCOUNT_JSON at the bundled "
            "resources/credentials.json — check that file is in the bundle."
        )

    if os.path.exists(raw):
        return service_account.Credentials.from_service_account_file(
            raw, scopes=_SCOPES,
        )

    # Inline JSON fallback (dev setups)
    try:
        info = json.loads(raw)
        return service_account.Credentials.from_service_account_info(
            info, scopes=_SCOPES,
        )
    except json.JSONDecodeError:
        raise ValueError(
            f"GOOGLE_SERVICE_ACCOUNT_JSON is not a valid file path or JSON. "
            f"Starts with: {raw[:60]!r}"
        )


# lru_cache replaces @st.cache_resource — one client per process, same
# caching semantics as the cloud portal. cache_discovery=False because
# googleapiclient's on-disk discovery cache fights with PyInstaller temp
# dirs (same reason we disabled it in the cloud portal).
@lru_cache(maxsize=1)
def get_sheets_service():
    return build("sheets", "v4", credentials=_load_credentials(),
                 cache_discovery=False)


@lru_cache(maxsize=1)
def get_drive_service():
    return build("drive", "v3", credentials=_load_credentials(),
                 cache_discovery=False)


@lru_cache(maxsize=1)
def get_slides_service():
    return build("slides", "v1", credentials=_load_credentials(),
                 cache_discovery=False)


def check_auth() -> tuple[bool, str]:
    try:
        _load_credentials()
        return True, "Google credentials loaded successfully."
    except ValueError as e:
        return False, str(e)
