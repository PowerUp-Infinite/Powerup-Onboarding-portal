"""
common.py — shared helpers across M1/M2/M3 workers.

- parse_uploaded_excel: turn any user-uploaded .xlsx into the canonical
  {'pf_level', 'scheme', 'riskgroup', 'lines', 'results', 'invested'} dict
  regardless of whether the tab names have spaces / underscores / odd casing.

- list_clients_in_excel: return [(pf_id, display_name), ...] so the GUI can
  show a dropdown when the upload has multiple clients.

- filter_data_to_pf_id: reduce every tab to a single PF_ID's rows.

- upload_pptx_to_drive: thin wrapper that re-uses portal.sheets.

- PROGRESS: a simple callable hook the GUI can set to surface status
  messages to the user without every worker depending on tkinter.
"""
from __future__ import annotations

import sys
import os
from io import BytesIO
from typing import Any

# Make sure desktop/config.py runs first so env + portal/ sys.path are ready.
import app_config  # noqa: F401  (imported for side-effects: bootstraps env + sys.path)

import pandas as pd


# ── progress hook ─────────────────────────────────────────────
class _ProgressSink:
    """Simple observer the GUI replaces with a Tk callback. Defaults to print."""
    def __init__(self):
        self._fn = lambda msg: print(msg)

    def __call__(self, msg: str) -> None:
        try:
            self._fn(msg)
        except Exception:
            pass

    def set(self, fn) -> None:
        self._fn = fn


PROGRESS = _ProgressSink()


# ── Excel parsing ─────────────────────────────────────────────
# Canonical tab key → list of acceptable sheet-name variants (lowercased,
# whitespace-normalised). The uploaded file may use "PF level" / "PF_level"
# / "pf_level" — all map to the same key.
_TAB_ALIASES: dict[str, tuple[str, ...]] = {
    'pf_level':   ('pf_level', 'pflevel', 'pf level'),
    'scheme':     ('scheme_level', 'schemelevel', 'scheme level', 'scheme'),
    'riskgroup':  ('riskgroup_level', 'riskgrouplevel', 'riskgroup level',
                   'risk_group_level', 'risk group level', 'riskgroup'),
    'lines':      ('lines',),
    'results':    ('results',),
    'invested':   ('invested_value_line', 'investedvalueline',
                   'invested value line', 'invested'),
}


def _normalise_tab_name(name: str) -> str:
    return name.strip().lower().replace('_', ' ').replace('  ', ' ')


def _resolve_tab_name(xlsx: pd.ExcelFile, canonical_key: str) -> str | None:
    """Return the actual sheet name in xlsx that matches the canonical key,
    or None if no tab matches."""
    aliases = _TAB_ALIASES.get(canonical_key, ())
    alias_norm = {_normalise_tab_name(a) for a in aliases}
    for sheet_name in xlsx.sheet_names:
        if _normalise_tab_name(sheet_name) in alias_norm:
            return sheet_name
    return None


def parse_uploaded_excel(xlsx_path: str) -> dict[str, pd.DataFrame]:
    """Read an uploaded M2-shape Excel and return a dict of DataFrames.
    Missing tabs come back as empty DataFrames. Column headers are stripped.
    """
    xl = pd.ExcelFile(xlsx_path)
    out: dict[str, pd.DataFrame] = {}
    for key in _TAB_ALIASES:
        actual = _resolve_tab_name(xl, key)
        if actual is None:
            out[key] = pd.DataFrame()
            continue
        df = pd.read_excel(xl, sheet_name=actual)
        df.columns = [str(c).lstrip('\ufeff').strip() for c in df.columns]
        # Datetime → string so downstream code (which expects sheet-read semantics)
        # sees the same format.
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%Y-%m-%d').fillna('')
        out[key] = df
    return out


# ── Client (PF_ID) discovery ─────────────────────────────────
def list_clients_in_excel(xlsx_path: str) -> list[tuple[str, str]]:
    """Return [(pf_id, display_name), ...] discovered from the PF_level tab
    first, then falling back to the Scheme_level tab. Names are read from
    a 'Name' or 'NAME' column when present; otherwise pf_id is reused."""
    data = parse_uploaded_excel(xlsx_path)

    clients: dict[str, str] = {}

    def _harvest(df: pd.DataFrame) -> None:
        if df.empty or 'PF_ID' not in df.columns:
            return
        # Find a name-ish column
        name_col = None
        for c in df.columns:
            if str(c).strip().lower() in ('name', 'client_name', 'customer_name',
                                          'investor_name'):
                name_col = c
                break
        for _, row in df.drop_duplicates('PF_ID').iterrows():
            pid = str(row.get('PF_ID', '')).strip()
            if not pid or pid.lower() == 'nan':
                continue
            nm = ''
            if name_col:
                nm = str(row.get(name_col, '')).strip()
                if nm.lower() == 'nan':
                    nm = ''
            if pid not in clients or (not clients[pid] and nm):
                clients[pid] = nm

    _harvest(data.get('pf_level', pd.DataFrame()))
    _harvest(data.get('scheme', pd.DataFrame()))

    return [(pid, name or pid)
            for pid, name in sorted(clients.items(), key=lambda x: x[1] or x[0])]


def filter_data_to_pf_id(data: dict[str, pd.DataFrame], pf_id: str
                         ) -> dict[str, pd.DataFrame]:
    """Return a copy of `data` where every tab is filtered to rows whose
    PF_ID matches. Tabs without a PF_ID column are returned unchanged.
    Numeric coercion matches portal/m2_engine.load_data()."""
    out: dict[str, Any] = {}
    pf_id_str = str(pf_id).strip()
    for key, df in data.items():
        if df.empty:
            out[key] = df.copy()
            continue
        if 'PF_ID' in df.columns:
            out[key] = df[df['PF_ID'].astype(str).str.strip() == pf_id_str].copy()
        else:
            out[key] = df.copy()

    # Numeric coercion (same columns the portal leaves as text)
    _text_cols = {
        'PF_ID', 'ISIN', 'NAME', 'FUND_NAME', 'FUND_STANDARD_NAME',
        'FUND_LEGAL_NAME', 'TYPE', 'POWERRATING', 'DISTRIBUTION_STATUS',
        'RISK_GROUP_L0', 'UPDATED_SUBCATEGORY', 'UPDATED_BROAD_CATEGORY_GROUP',
        'BROAD_CATEGORY_GROUP', 'DERIVED_CATEGORY', 'Purchase Mode',
        'BM', 'DIR_ISIN', 'ALT_ISIN_J', 'DATE',
    }
    for key in ('pf_level', 'riskgroup', 'scheme', 'results', 'lines', 'invested'):
        if key not in out:
            continue
        df = out[key]
        if df.empty:
            continue
        for col in df.columns:
            if col in _text_cols:
                continue
            df[col] = pd.to_numeric(df[col], errors='coerce')
        out[key] = df
    return out


# ── Drive upload ──────────────────────────────────────────────
def upload_pptx_to_drive(buf: BytesIO, filename: str, folder_id: str,
                         convert_to_slides: bool = True) -> dict:
    """Thin wrapper around portal/sheets.upload_pptx_to_drive so callers don't
    need to muck with sys.path themselves."""
    import sheets  # type: ignore  # portal/sheets.py via config.py bootstrap
    return sheets.upload_pptx_to_drive(
        buf, filename, folder_id, convert_to_slides=convert_to_slides,
    )


def fetch_questionnaire() -> pd.DataFrame:
    """Fetch the questionnaire responses DataFrame from Google Sheets."""
    import sheets  # type: ignore
    return sheets.read_questionnaire()
