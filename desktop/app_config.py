"""
config.py — desktop app configuration.

Loads environment variables from portal/.env at build time (baked in) and
resolves bundled resource paths (credentials.json, icons) whether running
from source or from a PyInstaller bundle.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path


def resource_path(*parts: str) -> str:
    """Return an absolute path to a bundled resource.
    Works both in dev (running from source) and in a PyInstaller one-file bundle
    (which extracts to sys._MEIPASS at runtime)."""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        base = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    else:
        base = Path(__file__).resolve().parent
    return str(base.joinpath(*parts))


def repo_root() -> Path:
    """Path to the repo root when running from source. Not used in bundled mode."""
    return Path(__file__).resolve().parent.parent


def _load_env_file() -> None:
    """Populate os.environ from the env file bundled alongside the app.

    Resolution order:
      1. desktop/resources/.env    (bundled at build time — works in both
         dev mode and frozen .exe/.app)
      2. portal/.env               (dev-mode fallback — only exists when
         running directly from the repo source tree)

    First file found wins.
    """
    candidates = [
        Path(resource_path('resources', '.env')),
        repo_root() / 'portal' / '.env',
    ]
    for env_path in candidates:
        if not env_path.exists():
            continue
        for line in env_path.read_text(encoding='utf-8').splitlines():
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            if '=' not in line:
                continue
            k, v = line.split('=', 1)
            k = k.strip()
            v = v.strip().strip('"').strip("'")
            os.environ.setdefault(k, v)
        break  # first hit wins


_load_env_file()


# ── Credentials ───────────────────────────────────────────────
# The service account JSON ships bundled at desktop/resources/credentials.json.
# In bundled mode it is extracted to sys._MEIPASS/resources/credentials.json.
CREDENTIALS_PATH = resource_path('resources', 'credentials.json')

# The portal's google_auth.py looks at GOOGLE_SERVICE_ACCOUNT_JSON — point it
# at our bundled file so the existing portal code works unchanged.
os.environ['GOOGLE_SERVICE_ACCOUNT_JSON'] = CREDENTIALS_PATH


# ── Runtime paths ─────────────────────────────────────────────
# Drive template / folder IDs are read from portal/.env (already loaded above).
def _req(name: str) -> str:
    v = os.environ.get(name, '').strip()
    if not v:
        raise EnvironmentError(
            f"Missing required env var {name!r}. Re-build the app — portal/.env "
            f"must contain this value."
        )
    return v


MAIN_SPREADSHEET_ID         = _req('MAIN_SPREADSHEET_ID')
QUESTIONNAIRE_SPREADSHEET_ID = _req('QUESTIONNAIRE_SPREADSHEET_ID')
TIMESERIES_SPREADSHEET_ID   = _req('TIMESERIES_SPREADSHEET_ID')
M3_SPREADSHEET_ID           = _req('M3_SPREADSHEET_ID')

M1_APPS_SCRIPT_URL          = _req('M1_APPS_SCRIPT_URL')
M1_OUTPUT_FOLDER_ID         = _req('M1_OUTPUT_FOLDER_ID')
M2_OUTPUT_FOLDER_ID         = _req('M2_OUTPUT_FOLDER_ID')
M3_OUTPUT_FOLDER_ID         = _req('M3_OUTPUT_FOLDER_ID')


# ── portal/ import bootstrap ──────────────────────────────────
# The desktop app re-uses portal/m2_engine.py, portal/m3_engine.py and
# portal/sheets.py directly. Those modules are flat scripts that import each
# other by bare name (e.g. `from google_auth import ...`), so we add portal/
# to sys.path whether running from source or from a bundle.
def _prepare_portal_imports() -> None:
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # PyInstaller extracts portal/ alongside desktop/
        portal_path = os.path.join(sys._MEIPASS, 'portal')  # type: ignore[attr-defined]
    else:
        portal_path = str(repo_root() / 'portal')
    if portal_path not in sys.path:
        sys.path.insert(0, portal_path)


_prepare_portal_imports()
