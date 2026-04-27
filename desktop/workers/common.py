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
    'is_demat':   ('is_demat', 'isdemat', 'is demat', 'demat'),
    'name_age':   ('name_age', 'nameage', 'name age', 'name and age'),
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
    """Return [(pf_id, display_name), ...] for every PF_ID in the PF_level tab.

    Source of truth is PF_level only — Scheme_level may carry stragglers that
    don't have a corresponding PF_level row, and including them in the picker
    causes downstream failures (PF_ID '...' not found in PF_level). The
    name column on PF_level is optional; if missing/blank, we fall back to
    the name_age tab matched by USER_ID, then to an empty name (the engine
    will warn and leave slides 1/2 for manual fill)."""
    data = parse_uploaded_excel(xlsx_path)

    pf_level = data.get('pf_level', pd.DataFrame())
    if pf_level.empty or 'PF_ID' not in pf_level.columns:
        return []

    name_col = next(
        (c for c in pf_level.columns
         if str(c).strip().lower() in ('name', 'client_name',
                                       'customer_name', 'investor_name')),
        None,
    )

    # Build name_age lookup (USER_ID -> NAME) from the uploaded Excel for the
    # rows where PF_level doesn't carry a name. Optional — empty if no tab.
    name_age = data.get('name_age', pd.DataFrame())
    na_lookup: dict[str, str] = {}
    if not name_age.empty and 'USER_ID' in name_age.columns:
        na_name_col = next((c for c in name_age.columns
                            if c.upper() == 'NAME'), None)
        if na_name_col:
            for _, r in name_age.iterrows():
                uid = str(r.get('USER_ID', '')).strip()
                nm  = str(r.get(na_name_col, '')).strip()
                if uid and uid.lower() != 'nan' and nm and nm.lower() != 'nan':
                    na_lookup[uid] = nm

    clients: list[tuple[str, str]] = []
    for _, row in pf_level.drop_duplicates('PF_ID').iterrows():
        pid = str(row.get('PF_ID', '')).strip()
        if not pid or pid.lower() == 'nan':
            continue
        nm = ''
        if name_col:
            v = row.get(name_col)
            if v is not None and not (isinstance(v, float) and pd.isna(v)):
                s = str(v).strip()
                if s and s.lower() != 'nan':
                    nm = s
        if not nm:
            nm = na_lookup.get(pid, '')
        clients.append((pid, nm))

    # Sort: rows with names first (alphabetical), then nameless PF_IDs.
    clients.sort(key=lambda x: (not x[1], x[1] or x[0]))
    return clients


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

    # Numeric coercion (same columns the portal leaves as text). Match
    # case-insensitively so columns like 'Name' (PF_level) survive too —
    # otherwise pd.to_numeric turns string names into NaN silently.
    _text_cols_ci = {
        'pf_id', 'isin', 'name', 'fund_name', 'fund_standard_name',
        'fund_legal_name', 'type', 'powerrating', 'distribution_status',
        'risk_group_l0', 'updated_subcategory', 'updated_broad_category_group',
        'broad_category_group', 'derived_category', 'purchase mode',
        'bm', 'dir_isin', 'alt_isin_j', 'date',
    }
    for key in ('pf_level', 'riskgroup', 'scheme', 'results', 'lines', 'invested'):
        if key not in out:
            continue
        df = out[key]
        if df.empty:
            continue
        for col in df.columns:
            if str(col).strip().lower() in _text_cols_ci:
                continue
            df[col] = pd.to_numeric(df[col], errors='coerce')
        out[key] = df

    # is_demat: keep IS_DEMAT as bool, coerce PCT_OF_USER + CURRENT_VALUE.
    if 'is_demat' in out and not out['is_demat'].empty:
        df = out['is_demat']
        for col in df.columns:
            if col in ('PF_ID', 'IS_DEMAT'):
                continue
            df[col] = pd.to_numeric(df[col], errors='coerce')
        out['is_demat'] = df
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
