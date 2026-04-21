"""
sheets.py — Google Sheets read/write layer for PowerUp Portal.

All data access goes through this module. Nothing else imports
google API clients directly — use google_auth.py for that.

Spreadsheets:
  MAIN_SPREADSHEET_ID       — client data (M1 + M2 shared)
  TIMESERIES_SPREADSHEET_ID — Lines + Invested_Value_Line (large, kept separate)
  M3_SPREADSHEET_ID         — M3 reference data (monthly refresh)
  QUESTIONNAIRE_SPREADSHEET_ID — existing questionnaire responses

Primary keys per sheet (for upsert/dedup):
  PF_level            → PF_ID
  Scheme_level        → PF_ID + ISIN
  Riskgroup_level     → PF_ID + RISK_GROUP_L0
  Lines               → PF_ID + DATE + TYPE        (time-series, auto-pruned at 1M rows)
  Results             → PF_ID + TYPE
  Invested_Value_Line → PF_ID + DATE               (time-series, auto-pruned at 1M rows)
  Scheme_Category     → Powerup Broad Category
  BASE_DATA           → ISIN
  AUM                 → ISIN
  Powerranking        → ISIN
  Upside_Downside     → Scheme ISIN
  Rolling_Returns     → ENTITYID + ROLLING_PERIOD

Auto-pruning (Lines + Invested_Value_Line):
  When row count exceeds TIMESERIES_ROW_LIMIT (default 1,000,000), the oldest
  rows are deleted automatically before appending new data. "Oldest" = rows
  with the earliest DATE values. No manual intervention required.
"""

import pandas as pd
import numpy as np
import time
from datetime import datetime, timezone, timedelta
from io import BytesIO
from typing import Optional

from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload
from google_auth import get_sheets_service, get_drive_service


def _execute_with_retry(request, max_attempts: int = 5, base_delay: float = 1.0):
    """Execute a Google API request with exponential backoff on transient errors.
    Retries on 409 (concurrent modification), 429 (rate limit), 5xx (server).
    """
    for attempt in range(max_attempts):
        try:
            return request.execute()
        except HttpError as e:
            status = getattr(e.resp, 'status', None)
            try:
                status = int(status) if status is not None else None
            except (TypeError, ValueError):
                status = None
            retryable = status in (409, 429, 500, 502, 503, 504)
            if not retryable or attempt == max_attempts - 1:
                raise
            delay = base_delay * (2 ** attempt)
            time.sleep(delay)
from config import (
    MAIN_SPREADSHEET_ID, M3_SPREADSHEET_ID,
    QUESTIONNAIRE_SPREADSHEET_ID, TIMESERIES_SPREADSHEET_ID,
    TIMESERIES_ROW_LIMIT,
    M1_OUTPUT_FOLDER_ID, M2_OUTPUT_FOLDER_ID,
    MainSheets, M3Sheets, TimeSeriesSheets,
)

# Re-export IDs so other modules can reference them as sheets.MAIN_SPREADSHEET_ID etc.
__all__ = [
    "MAIN_SPREADSHEET_ID", "M3_SPREADSHEET_ID",
    "TIMESERIES_SPREADSHEET_ID", "QUESTIONNAIRE_SPREADSHEET_ID",
]

# ─────────────────────────────────────────────────────────────
# Generic helpers
# ─────────────────────────────────────────────────────────────

def _sheet_to_df(spreadsheet_id: str, tab: str) -> pd.DataFrame:
    """Read an entire sheet tab and return as a DataFrame. Row 1 = headers."""
    svc = get_sheets_service()
    res = svc.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=tab,
        valueRenderOption="UNFORMATTED_VALUE",
        dateTimeRenderOption="FORMATTED_STRING",
    ).execute()
    values = res.get("values", [])
    if not values:
        return pd.DataFrame()
    headers = [str(h).strip() for h in values[0]]
    rows = values[1:]
    # Pad short rows so all rows have same width as headers
    padded = [r + [""] * (len(headers) - len(r)) for r in rows]
    df = pd.DataFrame(padded, columns=headers)
    # Drop fully-empty rows
    df = df.replace("", np.nan)
    df = df.dropna(how="all")
    df = df.infer_objects(copy=False).replace(np.nan, "")
    return df


def _get_sheet_id(spreadsheet_id: str, tab: str) -> int | None:
    """Return the numeric sheetId for a named tab."""
    svc = get_sheets_service()
    meta = svc.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(sheetId,title))"
    ).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == tab:
            return s["properties"]["sheetId"]
    return None


def _resize_sheet(spreadsheet_id: str, tab: str, rows: int, cols: int):
    """
    Resize a sheet to exactly `rows` rows and `cols` columns.
    Done in two passes to avoid Google Sheets' 10M cell limit being hit
    mid-operation (it would otherwise expand rows before shrinking cols).
    Pass 1: shrink columns to target.
    Pass 2: expand/shrink rows to target.
    """
    sheet_id = _get_sheet_id(spreadsheet_id, tab)
    if sheet_id is None:
        return
    svc = get_sheets_service()

    def _update(row_count=None, col_count=None):
        props = {"sheetId": sheet_id, "gridProperties": {}}
        fields = []
        if col_count is not None:
            props["gridProperties"]["columnCount"] = max(col_count, 1)
            fields.append("gridProperties.columnCount")
        if row_count is not None:
            props["gridProperties"]["rowCount"] = max(row_count, 1)
            fields.append("gridProperties.rowCount")
        svc.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [{"updateSheetProperties": {
                "properties": props,
                "fields": ",".join(fields),
            }}]}
        ).execute()

    # Pass 1: set columns first (reduces cell count if cols < current)
    _update(col_count=cols)
    # Pass 2: set rows (safe now — cols already minimised)
    _update(row_count=rows)


def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Convert all non-JSON-serializable types in a DataFrame before calling tolist().
    Vectorized — much faster than cell-by-cell iteration."""
    df = df.copy()
    for col in df.columns:
        dtype = df[col].dtype
        if pd.api.types.is_datetime64_any_dtype(dtype):
            df[col] = df[col].astype(str).replace("NaT", "")
        elif pd.api.types.is_bool_dtype(dtype):
            df[col] = df[col].astype(object)
    # Replace remaining NaN/NaT
    df = df.fillna("")
    return df


def _df_to_sheet(spreadsheet_id: str, tab: str, df: pd.DataFrame) -> int:
    """
    Overwrite a sheet tab completely with df (headers + data).
    Resizes the sheet to exact dimensions first to stay under the
    Google Sheets 10M cell limit. Writes in batches for large datasets.
    Returns number of rows written (excluding header).
    """
    svc = get_sheets_service()

    # Resize to exact dimensions (rows+1 for header, actual col count)
    n_rows = len(df) + 1
    n_cols = len(df.columns) if not df.empty else 1
    _resize_sheet(spreadsheet_id, tab, n_rows, n_cols)

    # Clear existing content
    svc.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id, range=tab
    ).execute()

    if df.empty:
        return 0

    clean = _sanitize_df(df)
    all_values = [clean.columns.tolist()] + clean.values.tolist()

    # Write in batches of 50,000 rows — RAW is faster (no formula parsing)
    BATCH = 50_000
    for i in range(0, len(all_values), BATCH):
        chunk = all_values[i:i + BATCH]
        start_row = i + 1  # 1-based
        svc.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{tab}!A{start_row}",
            valueInputOption="RAW",
            body={"values": chunk},
        ).execute()

    return len(df)


def _upsert(
    spreadsheet_id: str,
    tab: str,
    new_df: pd.DataFrame,
    key_cols: list[str],
) -> dict:
    """
    Upsert new_df into the sheet:
    - Rows whose key_cols values match existing rows → replaced entirely
    - Rows with new key_cols values → appended
    Returns {"replaced": n, "added": n, "total": n}
    """
    existing = _sheet_to_df(spreadsheet_id, tab)

    if existing.empty:
        written = _df_to_sheet(spreadsheet_id, tab, new_df)
        return {"replaced": 0, "added": written, "total": written}

    # Ensure key cols exist in both frames
    for col in key_cols:
        if col not in existing.columns:
            existing[col] = ""
        if col not in new_df.columns:
            raise ValueError(f"Key column '{col}' not found in uploaded data")

    # Build composite key strings for matching
    def _key(df):
        return df[key_cols].astype(str).agg("||".join, axis=1)

    new_keys = set(_key(new_df))
    mask_replace = _key(existing).isin(new_keys)
    n_replaced = int(mask_replace.sum())

    # Keep existing rows that are NOT being replaced, then append new
    kept = existing[~mask_replace]
    merged = pd.concat([kept, new_df], ignore_index=True).infer_objects(copy=False)
    _df_to_sheet(spreadsheet_id, tab, merged)

    return {
        "replaced": n_replaced,
        "added": max(len(new_df) - n_replaced, 0),
        "total": len(merged),
    }


# ─────────────────────────────────────────────────────────────
# Main Spreadsheet — READ
# ─────────────────────────────────────────────────────────────

def read_pf_level() -> pd.DataFrame:
    return _sheet_to_df(MAIN_SPREADSHEET_ID, MainSheets.PF_LEVEL)

def read_scheme_level() -> pd.DataFrame:
    return _sheet_to_df(MAIN_SPREADSHEET_ID, MainSheets.SCHEME_LEVEL)

def read_riskgroup_level() -> pd.DataFrame:
    return _sheet_to_df(MAIN_SPREADSHEET_ID, MainSheets.RISKGROUP_LEVEL)

def read_results() -> pd.DataFrame:
    return _sheet_to_df(MAIN_SPREADSHEET_ID, MainSheets.RESULTS)

def read_scheme_category() -> pd.DataFrame:
    return _sheet_to_df(MAIN_SPREADSHEET_ID, MainSheets.SCHEME_CATEGORY)

def read_base_data() -> pd.DataFrame:
    return _sheet_to_df(MAIN_SPREADSHEET_ID, MainSheets.BASE_DATA)

def read_is_demat() -> pd.DataFrame:
    """SOA/Demat split per PF_ID. Same shape as the 'Is demat' tab in the
    uploaded Excel — used by do_slide4 to render the SOA % / Demat % fields.
    Empty DataFrame if the tab is missing or has only a header row."""
    return _sheet_to_df(MAIN_SPREADSHEET_ID, MainSheets.IS_DEMAT)

def read_questionnaire() -> pd.DataFrame:
    """Read the questionnaire responses sheet (flat, single tab)."""
    return _sheet_to_df(QUESTIONNAIRE_SPREADSHEET_ID, "Sheet1")


# ─────────────────────────────────────────────────────────────
# PF_ID ↔ Questionnaire Name mapping
# ─────────────────────────────────────────────────────────────

def read_pf_id_mapping() -> dict[str, str]:
    """
    Read the PF_ID_Mapping sheet. Returns {pf_id: questionnaire_name}.
    Creates the sheet if it doesn't exist.
    """
    try:
        df = _sheet_to_df(MAIN_SPREADSHEET_ID, MainSheets.PF_ID_MAPPING)
    except Exception:
        # Sheet doesn't exist yet — create it
        _create_mapping_sheet()
        return {}

    if df.empty:
        return {}

    mapping = {}
    for _, row in df.iterrows():
        pid = str(row.get("PF_ID", "")).strip()
        qname = str(row.get("Questionnaire_Name", "")).strip()
        if pid and qname and qname.lower() != "nan":
            mapping[pid] = qname
    return mapping


def save_pf_id_mapping(pf_id: str, questionnaire_name: str):
    """Save or update a single PF_ID → questionnaire name mapping."""
    svc = get_sheets_service()

    # Try to read existing mappings
    try:
        df = _sheet_to_df(MAIN_SPREADSHEET_ID, MainSheets.PF_ID_MAPPING)
    except Exception:
        _create_mapping_sheet()
        df = pd.DataFrame(columns=["PF_ID", "Questionnaire_Name"])

    if df.empty:
        df = pd.DataFrame(columns=["PF_ID", "Questionnaire_Name"])

    # Update existing or add new
    if pf_id in df["PF_ID"].astype(str).values:
        df.loc[df["PF_ID"].astype(str) == pf_id, "Questionnaire_Name"] = questionnaire_name
    else:
        new_row = pd.DataFrame([{"PF_ID": pf_id, "Questionnaire_Name": questionnaire_name}])
        df = pd.concat([df, new_row], ignore_index=True)

    _df_to_sheet(MAIN_SPREADSHEET_ID, MainSheets.PF_ID_MAPPING, df)


def _create_mapping_sheet():
    """Create the PF_ID_Mapping tab in the main spreadsheet."""
    svc = get_sheets_service()
    try:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=MAIN_SPREADSHEET_ID,
            body={"requests": [{
                "addSheet": {"properties": {"title": MainSheets.PF_ID_MAPPING}}
            }]}
        ).execute()
        # Write headers
        svc.spreadsheets().values().update(
            spreadsheetId=MAIN_SPREADSHEET_ID,
            range=f"{MainSheets.PF_ID_MAPPING}!A1:B1",
            valueInputOption="RAW",
            body={"values": [["PF_ID", "Questionnaire_Name"]]},
        ).execute()
    except Exception:
        pass  # sheet may already exist

def get_pf_ids() -> list[str]:
    """Return sorted list of all PF_IDs — used to populate dropdowns in M1/M2 tabs."""
    df = read_pf_level()
    if df.empty or "PF_ID" not in df.columns:
        return []
    return sorted(df["PF_ID"].astype(str).unique().tolist())

def get_client_name(pf_id: str) -> Optional[str]:
    """
    Return a display name for a PF_ID from Scheme_level (FUND_NAME is per scheme,
    so we look at the questionnaire or PF_level for a client name).
    M2 app.py derives name from the questionnaire — we do the same here.
    Returns None if not found.
    """
    try:
        q = read_questionnaire()
        if q.empty:
            return None
        # Find the PF_ID column (may be named differently)
        pf_col = next((c for c in q.columns if "pf" in c.lower() and "id" in c.lower()), None)
        name_col = next((c for c in q.columns if "name" in c.lower()), None)
        if not pf_col or not name_col:
            return None
        row = q[q[pf_col].astype(str) == str(pf_id)]
        if row.empty:
            return None
        return str(row.iloc[0][name_col]).strip() or None
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────
# Time-Series Spreadsheet — READ
# ─────────────────────────────────────────────────────────────

def _col_letter(n: int) -> str:
    """0-indexed column number -> A1 letter (0=A, 1=B, ..., 26=AA)."""
    s = ''
    n += 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _sheet_to_df_filtered(spreadsheet_id: str, tab: str, key_col: str,
                          key_val: str) -> pd.DataFrame:
    """
    Read only the rows of `tab` where column `key_col` == `key_val`.
    Avoids materializing the entire sheet in memory — critical for
    Streamlit Cloud's 1 GB limit when `Lines` has >1M rows.

    Strategy:
      1. Fetch headers only (row 1).
      2. Fetch only the key column (A:A or whichever), find row indices
         whose cell value matches key_val.
      3. If matches > 0, fetch only those specific rows via batchGet.
      4. Return a DataFrame with just the matching rows.
    """
    svc = get_sheets_service()

    # Step 1: headers
    hdr_res = svc.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f'{tab}!1:1',
    ).execute()
    headers = [str(h).strip() for h in hdr_res.get("values", [[]])[0]]
    if not headers:
        return pd.DataFrame()
    try:
        key_idx = headers.index(key_col)
    except ValueError:
        # Key column doesn't exist — fall back to unfiltered read
        return _sheet_to_df(spreadsheet_id, tab)

    # Step 2: key column only
    key_letter = _col_letter(key_idx)
    key_res = svc.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f'{tab}!{key_letter}2:{key_letter}',
        majorDimension='COLUMNS',
    ).execute()
    col_vals = key_res.get("values", [[]])
    col_vals = col_vals[0] if col_vals else []

    # Row numbers are 1-indexed; header is row 1, so data starts at row 2.
    target = str(key_val).strip()
    matching_rows = [i + 2 for i, v in enumerate(col_vals)
                     if str(v).strip() == target]

    if not matching_rows:
        return pd.DataFrame(columns=headers)

    # Step 3: batchGet only matching rows. Collapse contiguous runs into
    # single A:row:row ranges to minimise request size.
    last_col = _col_letter(len(headers) - 1)
    ranges = []
    run_start = matching_rows[0]
    prev = run_start
    for r in matching_rows[1:]:
        if r == prev + 1:
            prev = r
            continue
        ranges.append(f'{tab}!A{run_start}:{last_col}{prev}')
        run_start = r
        prev = r
    ranges.append(f'{tab}!A{run_start}:{last_col}{prev}')

    batch_res = svc.spreadsheets().values().batchGet(
        spreadsheetId=spreadsheet_id,
        ranges=ranges,
        valueRenderOption="UNFORMATTED_VALUE",
        dateTimeRenderOption="FORMATTED_STRING",
    ).execute()

    rows = []
    for value_range in batch_res.get("valueRanges", []):
        for r in value_range.get("values", []):
            rows.append(r + [""] * (len(headers) - len(r)))

    if not rows:
        return pd.DataFrame(columns=headers)

    df = pd.DataFrame(rows, columns=headers)
    df = df.replace("", np.nan)
    df = df.dropna(how="all")
    df = df.infer_objects(copy=False).replace(np.nan, "")
    return df


def read_lines(pf_id: str | None = None) -> pd.DataFrame:
    """Read Lines timeseries. If pf_id is provided, only that client's rows
    are fetched — drastically reduces memory usage on Streamlit Cloud."""
    if pf_id:
        return _sheet_to_df_filtered(
            TIMESERIES_SPREADSHEET_ID, TimeSeriesSheets.LINES, 'PF_ID', pf_id
        )
    return _sheet_to_df(TIMESERIES_SPREADSHEET_ID, TimeSeriesSheets.LINES)


def read_invested_value_line(pf_id: str | None = None) -> pd.DataFrame:
    """Read Invested_Value_Line timeseries. If pf_id is provided, only that
    client's rows are fetched."""
    if pf_id:
        return _sheet_to_df_filtered(
            TIMESERIES_SPREADSHEET_ID, TimeSeriesSheets.INVESTED_VALUE_LINE,
            'PF_ID', pf_id,
        )
    return _sheet_to_df(TIMESERIES_SPREADSHEET_ID, TimeSeriesSheets.INVESTED_VALUE_LINE)


# ─────────────────────────────────────────────────────────────
# Time-Series Spreadsheet — UPSERT with auto-pruning
# ─────────────────────────────────────────────────────────────

def _get_row_count(spreadsheet_id: str, tab: str) -> int:
    """Get current data row count from sheet metadata (no data read)."""
    svc = get_sheets_service()
    meta = svc.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        ranges=[tab],
        fields="sheets(properties(sheetId,title,gridProperties))"
    ).execute()
    props = next(
        (s["properties"] for s in meta["sheets"]
         if s["properties"]["title"] == tab), None
    )
    if not props:
        return 0
    return max(props["gridProperties"]["rowCount"] - 1, 0)  # subtract header


def _read_pf_id_column(spreadsheet_id: str, tab: str) -> list[str]:
    """Read only the PF_ID column (col A) to avoid loading entire sheet."""
    svc = get_sheets_service()
    res = svc.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{tab}!A:A",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()
    values = res.get("values", [])
    if len(values) <= 1:
        return []
    return [str(v[0]).strip() if v else "" for v in values[1:]]


def _delete_rows_for_pf_ids(spreadsheet_id: str, tab: str, pf_ids_to_remove: set[str]) -> int:
    """
    Delete all rows in tab where PF_ID (col A) is in pf_ids_to_remove.
    Works without reading the full sheet — reads only PF_ID column.
    Returns number of rows deleted.
    """
    if not pf_ids_to_remove:
        return 0

    pf_ids = _read_pf_id_column(spreadsheet_id, tab)
    if not pf_ids:
        return 0

    # Find 0-based data-row indices to delete (rows where PF_ID matches)
    rows_to_delete = [i for i, pid in enumerate(pf_ids) if pid in pf_ids_to_remove]
    if not rows_to_delete:
        return 0

    sheet_id = _get_sheet_id(spreadsheet_id, tab)
    if sheet_id is None:
        return 0

    # Build delete requests in reverse order (bottom-up to keep indices stable)
    # Batch contiguous ranges for efficiency
    svc = get_sheets_service()
    requests = []
    # Convert to 1-based sheet row indices (add 1 for header)
    sheet_rows = sorted([r + 1 for r in rows_to_delete], reverse=True)

    # Group contiguous runs
    ranges = []
    i = 0
    while i < len(sheet_rows):
        end = sheet_rows[i]
        start = end
        while i + 1 < len(sheet_rows) and sheet_rows[i + 1] == start - 1:
            start = sheet_rows[i + 1]
            i += 1
        ranges.append((start, end))
        i += 1

    for start, end in ranges:
        requests.append({
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": start,
                    "endIndex": end + 1,
                }
            }
        })

    # Execute in batches of 50 to avoid request size limits
    total_deleted = 0
    for batch_start in range(0, len(requests), 50):
        batch = requests[batch_start:batch_start + 50]
        svc.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": batch}
        ).execute()
        total_deleted += sum(
            r["deleteDimension"]["range"]["endIndex"] - r["deleteDimension"]["range"]["startIndex"]
            for r in batch
        )

    return total_deleted


def _append_rows(spreadsheet_id: str, tab: str, df: pd.DataFrame) -> int:
    """Append rows to the bottom of a sheet (no header write). Returns rows appended.
    Uses _execute_with_retry because immediately after _delete_rows_for_pf_ids the
    Sheets backend can return 409 "operation aborted" while it reindexes."""
    if df.empty:
        return 0

    svc = get_sheets_service()
    clean = _sanitize_df(df)
    values = clean.values.tolist()

    BATCH = 50_000
    appended = 0
    for i in range(0, len(values), BATCH):
        chunk = values[i:i + BATCH]
        req = svc.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=tab,
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": chunk},
        )
        _execute_with_retry(req)
        appended += len(chunk)

    return appended


def _prune_oldest_rows(spreadsheet_id: str, tab: str, limit: int) -> int:
    """
    If sheet has more than `limit` data rows, delete the oldest rows
    (from the top, after header) to bring count back to limit.
    Returns number of rows deleted.
    """
    current = _get_row_count(spreadsheet_id, tab)
    if current <= limit:
        return 0

    rows_to_delete = current - limit
    sheet_id = _get_sheet_id(spreadsheet_id, tab)
    if sheet_id is None:
        return 0

    svc = get_sheets_service()
    svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": 1,
                    "endIndex": 1 + rows_to_delete,
                }
            }
        }]}
    ).execute()

    return rows_to_delete


def _upsert_timeseries(tab: str, date_col: str, new_df: pd.DataFrame, key_cols: list[str]) -> dict:
    """
    Upsert time-series data without reading the full sheet:
    1. Get PF_IDs from new data
    2. Delete ALL existing rows for those PF_IDs (avoids partial data)
    3. Append new rows
    4. Prune oldest rows if total exceeds TIMESERIES_ROW_LIMIT
    Returns {"replaced": n, "added": n, "total": n, "pruned": n}
    """
    new_pf_ids = set(new_df["PF_ID"].astype(str).unique()) if "PF_ID" in new_df.columns else set()

    # Step 1: Delete existing rows for these PF_IDs
    n_deleted = _delete_rows_for_pf_ids(TIMESERIES_SPREADSHEET_ID, tab, new_pf_ids)

    # Step 2: Append new data
    n_appended = _append_rows(TIMESERIES_SPREADSHEET_ID, tab, new_df)

    # Step 3: Prune if over limit
    pruned = _prune_oldest_rows(TIMESERIES_SPREADSHEET_ID, tab, TIMESERIES_ROW_LIMIT)

    total = _get_row_count(TIMESERIES_SPREADSHEET_ID, tab)

    return {
        "replaced": n_deleted,
        "added":    n_appended,
        "total":    total,
        "pruned":   pruned,
    }


def upsert_lines(df: pd.DataFrame) -> dict:
    return _upsert_timeseries(TimeSeriesSheets.LINES, "DATE", df, ["PF_ID", "DATE", "TYPE"])

def upsert_invested_value_line(df: pd.DataFrame) -> dict:
    return _upsert_timeseries(TimeSeriesSheets.INVESTED_VALUE_LINE, "DATE", df, ["PF_ID", "DATE"])


# ─────────────────────────────────────────────────────────────
# M3 Reference Spreadsheet — READ
# ─────────────────────────────────────────────────────────────

def read_m3_aum() -> pd.DataFrame:
    return _sheet_to_df(M3_SPREADSHEET_ID, M3Sheets.AUM)

def read_m3_powerranking() -> pd.DataFrame:
    return _sheet_to_df(M3_SPREADSHEET_ID, M3Sheets.POWERRANKING)

def read_m3_upside_downside() -> pd.DataFrame:
    return _sheet_to_df(M3_SPREADSHEET_ID, M3Sheets.UPSIDE_DOWNSIDE)

def read_m3_rolling_returns() -> pd.DataFrame:
    return _sheet_to_df(M3_SPREADSHEET_ID, M3Sheets.ROLLING_RETURNS)


# ─────────────────────────────────────────────────────────────
# Main Spreadsheet — UPSERT (used by Data Manager tab)
# ─────────────────────────────────────────────────────────────

def upsert_pf_level(df: pd.DataFrame) -> dict:
    return _upsert(MAIN_SPREADSHEET_ID, MainSheets.PF_LEVEL, df, ["PF_ID"])

def upsert_scheme_level(df: pd.DataFrame) -> dict:
    return _upsert(MAIN_SPREADSHEET_ID, MainSheets.SCHEME_LEVEL, df, ["PF_ID", "ISIN"])

def upsert_riskgroup_level(df: pd.DataFrame) -> dict:
    return _upsert(MAIN_SPREADSHEET_ID, MainSheets.RISKGROUP_LEVEL, df, ["PF_ID", "RISK_GROUP_L0"])

def upsert_results(df: pd.DataFrame) -> dict:
    return _upsert(MAIN_SPREADSHEET_ID, MainSheets.RESULTS, df, ["PF_ID", "TYPE"])

def upsert_scheme_category(df: pd.DataFrame) -> dict:
    return _upsert(MAIN_SPREADSHEET_ID, MainSheets.SCHEME_CATEGORY, df, ["Powerup Broad Category"])


# ─────────────────────────────────────────────────────────────
# M3 Reference Spreadsheet — WRITE (monthly refresh, full replace)
# ─────────────────────────────────────────────────────────────

def write_m3_aum(df: pd.DataFrame) -> int:
    return _df_to_sheet(M3_SPREADSHEET_ID, M3Sheets.AUM, df)

def write_m3_powerranking(df: pd.DataFrame) -> int:
    return _df_to_sheet(M3_SPREADSHEET_ID, M3Sheets.POWERRANKING, df)

def write_m3_upside_downside(df: pd.DataFrame) -> int:
    return _df_to_sheet(M3_SPREADSHEET_ID, M3Sheets.UPSIDE_DOWNSIDE, df)

def write_m3_rolling_returns(df: pd.DataFrame) -> int:
    return _df_to_sheet(M3_SPREADSHEET_ID, M3Sheets.ROLLING_RETURNS, df)


# ─────────────────────────────────────────────────────────────
# Drive upload — upload generated files (PPTX → Google Slides)
# ─────────────────────────────────────────────────────────────

def upload_pptx_to_drive(
    buf: BytesIO,
    filename: str,
    folder_id: str,
    convert_to_slides: bool = True,
) -> dict:
    """
    Upload a PPTX BytesIO buffer to Google Drive.

    Args:
        buf: BytesIO containing the PPTX file.
        filename: Display name for the file in Drive.
        folder_id: Drive folder ID to upload into.
        convert_to_slides: If True, convert to Google Slides format.

    Returns:
        {"id": file_id, "url": web_view_link, "name": filename}
    """
    drive = get_drive_service()

    mime_pptx = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    mime_slides = "application/vnd.google-apps.presentation"

    file_metadata = {
        "name": filename.replace(".pptx", "") if convert_to_slides else filename,
        "parents": [folder_id],
    }
    if convert_to_slides:
        file_metadata["mimeType"] = mime_slides

    media = MediaIoBaseUpload(buf, mimetype=mime_pptx, resumable=True)

    created = drive.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, name, webViewLink",
        supportsAllDrives=True,
    ).execute()

    return {
        "id": created["id"],
        "url": created.get("webViewLink", f"https://docs.google.com/presentation/d/{created['id']}/edit"),
        "name": created["name"],
    }


def upload_docx_to_drive(
    buf: BytesIO,
    filename: str,
    folder_id: str,
    convert_to_gdoc: bool = False,
) -> dict:
    """
    Upload a DOCX BytesIO buffer to Google Drive.

    Returns:
        {"id": file_id, "url": web_view_link, "name": filename}
    """
    drive = get_drive_service()

    mime_docx = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    file_metadata = {
        "name": filename.replace(".docx", "") if convert_to_gdoc else filename,
        "parents": [folder_id],
    }
    if convert_to_gdoc:
        file_metadata["mimeType"] = "application/vnd.google-apps.document"

    media = MediaIoBaseUpload(buf, mimetype=mime_docx, resumable=True)

    created = drive.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, name, webViewLink",
        supportsAllDrives=True,
    ).execute()

    return {
        "id": created["id"],
        "url": created.get("webViewLink", f"https://drive.google.com/file/d/{created['id']}/view"),
        "name": created["name"],
    }


def export_drive_file_as_pdf(file_id: str) -> BytesIO:
    """Export a Google Drive file (Docs/Slides/Sheets) as PDF."""
    drive = get_drive_service()
    request = drive.files().export_media(
        fileId=file_id,
        mimeType="application/pdf",
    )
    from googleapiclient.http import MediaIoBaseDownload
    buf = BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────
# Drive cleanup — auto-delete M1/M2 outputs older than 7 days
# ─────────────────────────────────────────────────────────────

def cleanup_old_outputs(max_age_days: int = 7) -> dict[str, int]:
    """
    Delete files in M1 and M2 output folders that are older than max_age_days.
    Returns {"m1": count_deleted, "m2": count_deleted}.
    """
    drive = get_drive_service()
    cutoff = (datetime.now(timezone.utc) - timedelta(days=max_age_days)).isoformat()
    deleted = {"m1": 0, "m2": 0}

    folders = [
        ("m1", M1_OUTPUT_FOLDER_ID),
        ("m2", M2_OUTPUT_FOLDER_ID),
    ]

    for key, folder_id in folders:
        if not folder_id:
            continue
        try:
            # Find files older than cutoff in this folder
            query = (
                f"'{folder_id}' in parents"
                f" and createdTime < '{cutoff}'"
                f" and trashed = false"
            )
            page_token = None
            while True:
                resp = drive.files().list(
                    q=query,
                    fields="nextPageToken, files(id, name, createdTime)",
                    pageSize=100,
                    pageToken=page_token,
                ).execute()

                for f in resp.get("files", []):
                    try:
                        drive.files().delete(fileId=f["id"]).execute()
                        deleted[key] += 1
                    except Exception:
                        pass  # skip files we can't delete (permissions, etc.)

                page_token = resp.get("nextPageToken")
                if not page_token:
                    break
        except Exception:
            pass  # folder not accessible, skip silently

    return deleted


# ─────────────────────────────────────────────────────────────
# Dataset auto-detection (used by Data Manager tab)
# ─────────────────────────────────────────────────────────────

# Maps a frozenset of expected columns → (sheet tab name, upsert function, key cols)
_DATASET_SIGNATURES = [
    (
        {"PF_ID", "PF_XIRR", "PF_CURRENT_VALUE", "INVESTED_VALUE"},
        MainSheets.PF_LEVEL, upsert_pf_level, ["PF_ID"],
    ),
    (
        {"PF_ID", "ISIN", "FUND_NAME", "XIRR_VALUE", "CURRENT_VALUE"},
        MainSheets.SCHEME_LEVEL, upsert_scheme_level, ["PF_ID", "ISIN"],
    ),
    (
        {"PF_ID", "RISK_GROUP_L0", "XIRR_VALUE"},
        MainSheets.RISKGROUP_LEVEL, upsert_riskgroup_level, ["PF_ID", "RISK_GROUP_L0"],
    ),
    (
        {"DATE", "PF_ID", "TYPE", "CURRENT_VALUE"},
        TimeSeriesSheets.LINES, upsert_lines, ["PF_ID", "DATE", "TYPE"],
    ),
    (
        {"PF_ID", "TYPE", "XIRR", "TOTAL_TAX_PAID"},
        MainSheets.RESULTS, upsert_results, ["PF_ID", "TYPE"],
    ),
    (
        {"PF_ID", "DATE", "INVESTED_AMOUNT"},
        TimeSeriesSheets.INVESTED_VALUE_LINE, upsert_invested_value_line, ["PF_ID", "DATE"],
    ),
    (
        {"Powerup Broad Category", "Proposed Sub-Category", "V1 Risk Group L0"},
        MainSheets.SCHEME_CATEGORY, upsert_scheme_category, ["Powerup Broad Category"],
    ),
    # M3 reference (full replace — no upsert needed)
    ({"ISIN", "FUND_NAME", "AUM", "DAILYNAV"}, M3Sheets.AUM, write_m3_aum, None),
    ({"ISIN", "POWERRANK", "POWERRATING"}, M3Sheets.POWERRANKING, write_m3_powerranking, None),
    ({"Scheme ISIN", "Downside Capture Ratio", "Upside Capture Ratio"}, M3Sheets.UPSIDE_DOWNSIDE, write_m3_upside_downside, None),
    ({"ENTITYID", "RETURN_VALUE", "ROLLING_PERIOD"}, M3Sheets.ROLLING_RETURNS, write_m3_rolling_returns, None),
]


def detect_dataset(df: pd.DataFrame) -> Optional[tuple]:
    """
    Given a DataFrame, identify which dataset it is by matching column signatures.
    Returns (tab_name, write_fn, key_cols) or None if unrecognised.
    key_cols is None for M3 reference sheets (full replace, no upsert).
    """
    cols = set(df.columns)
    for sig, tab, fn, keys in _DATASET_SIGNATURES:
        if sig.issubset(cols):
            return tab, fn, keys
    return None
