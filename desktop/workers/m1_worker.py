"""
m1_worker.py — M1 client report, desktop flow.

Reuses the existing flow (sync to Main Data sheet, then fire the Apps Script
web app). The Apps Script reads from MAIN_SPREADSHEET_ID and writes the
output to M1_OUTPUT_FOLDER_ID on Drive, so there's no upload step from our
side — we just wait for the Apps Script response and hand the user the URL.

Returns {'url': ..., 'title': ...} or raises.
"""
from __future__ import annotations

import app_config                        # noqa: F401
import pandas as pd
import requests

from workers.common import PROGRESS


# M1 Apps Script only reads SCHEME_DATA / PF_LEVEL / BASE_DATA from the
# Main spreadsheet. BASE_DATA is static reference data that's already in
# Sheets — no need to upload. Riskgroup/Results/Lines/Invested are M2/M3
# data that the M1 report doesn't touch, so syncing them here would just
# pollute the main sheet with unused rows AND make every M1 generation
# 30-60 seconds slower (Lines alone is 1M+ rows on the timeseries sheet).
_TAB_UPSERT: dict[str, str] = {
    'pf_level':     'upsert_pf_level',
    'scheme_level': 'upsert_scheme_level',
}

_TAB_ALIASES: dict[str, str] = {}
for _k in _TAB_UPSERT:
    _TAB_ALIASES[_k.lower()] = _k
    _TAB_ALIASES[_k.lower().replace('_', '')] = _k
    _TAB_ALIASES[_k.lower().replace('_', ' ')] = _k

def _sync_excel_to_sheets(xlsx_path: str, only_pf_id: str) -> list[str]:
    """Upsert PF_level + Scheme_level from xlsx into the Main Data
    spreadsheet. Filters each tab to rows for `only_pf_id` to keep memory
    sane and to avoid clobbering other clients' rows.
    Other tabs (Lines / Invested / Riskgroup / Results) are deliberately
    NOT synced — the M1 Apps Script doesn't read them."""
    import sheets  # type: ignore
    import gc

    synced: list[str] = []
    xl = pd.ExcelFile(xlsx_path)

    for sheet_name in xl.sheet_names:
        canonical = _TAB_ALIASES.get(sheet_name.strip().lower())
        if not canonical:
            continue
        upsert_fn = getattr(sheets, _TAB_UPSERT[canonical])
        try:
            df = pd.read_excel(xl, sheet_name=sheet_name)
            df.columns = [str(c).lstrip('\ufeff').strip() for c in df.columns]
            if df.empty:
                continue
            # Always filter to the selected PF_ID — M1 only ever reads this
            # one client. Avoids overwriting other clients' rows.
            if 'PF_ID' in df.columns:
                df = df[df['PF_ID'].astype(str) == str(only_pf_id)].copy()
                if df.empty:
                    continue
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].dt.strftime('%Y-%m-%d').fillna('')
            PROGRESS(f"  Syncing {canonical} ({len(df)} rows)...")
            upsert_fn(df)
            synced.append(canonical)
        except Exception as e:
            PROGRESS(f"  WARN: couldn't sync tab {sheet_name!r}: {e}")
        finally:
            try:
                del df
            except Exception:
                pass
            gc.collect()

    return synced


def _call_apps_script(pf_id: str) -> dict:
    PROGRESS("Calling M1 Apps Script (this can take up to 2 minutes)...")
    resp = requests.post(
        app_config.M1_APPS_SCRIPT_URL,
        json={"pf_id": pf_id},
        timeout=120,
        allow_redirects=True,
    )
    resp.raise_for_status()
    try:
        data = resp.json()
    except Exception:
        raise ValueError(
            f"Apps Script returned non-JSON (status {resp.status_code}): "
            f"{resp.text[:300]}"
        )
    if data.get('status') == 'error':
        raise ValueError(f"Apps Script error: {data.get('message', 'unknown error')}")
    return data


def generate(xlsx_path: str, pf_id: str) -> dict:
    """Run the full M1 pipeline. Returns {'url', 'title'}."""
    PROGRESS(f"[1/2] Syncing uploaded data for PF_ID {pf_id}...")
    synced = _sync_excel_to_sheets(xlsx_path, pf_id)
    if not synced:
        raise ValueError(
            "No recognised data tabs in the uploaded file. "
            "Expected: PF_level, Scheme_level, Riskgroup_level, Results, "
            "Lines, Invested_Value_Line."
        )
    PROGRESS(f"  Synced {len(synced)} tabs: {', '.join(synced)}")

    PROGRESS(f"[2/2] Generating M1 report for {pf_id}...")
    result = _call_apps_script(pf_id)
    return {
        'url': result.get('url', ''),
        'title': result.get('title', f'M1 Report — {pf_id}'),
    }
