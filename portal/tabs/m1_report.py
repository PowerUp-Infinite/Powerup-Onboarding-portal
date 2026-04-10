"""
tabs/m1_report.py — M1 Client Report Sheet tab.

Flow (primary):
  1. User uploads an Excel file containing PF_level / Scheme_level data
  2. Portal extracts PF_IDs + client names from the file
  3. If multiple PF_IDs → user picks from a name-based dropdown
  4. "Generate Report" → HTTP POST to Apps Script Web App
  5. Portal shows clickable link to the generated sheet

Flow (fallback):
  If no file is uploaded, a dropdown populated from Google Sheets is shown instead.
"""

import pandas as pd
import requests
import streamlit as st

import sheets
from config import M1_APPS_SCRIPT_URL, M1_OUTPUT_FOLDER_ID


# ── Tab name → upsert function mapping ──────────────────────
_TAB_UPSERT_MAP = {
    'PF_level':            sheets.upsert_pf_level,
    'Scheme_level':        sheets.upsert_scheme_level,
    'Riskgroup_level':     sheets.upsert_riskgroup_level,
    'Results':             sheets.upsert_results,
    'Lines':               sheets.upsert_lines,
    'Invested_Value_Line': sheets.upsert_invested_value_line,
}

_TAB_ALIASES = {}
for _canonical in _TAB_UPSERT_MAP:
    _TAB_ALIASES[_canonical.lower()] = _canonical
    _TAB_ALIASES[_canonical.lower().replace('_', '')] = _canonical
    _TAB_ALIASES[_canonical.lower().replace('_', ' ')] = _canonical


def _sync_upload_to_sheets(uploaded_file, only_pf_id: str | None = None) -> list[str]:
    """Parse uploaded Excel and upsert each recognised tab into Google Sheets.
    Filters large per-client tabs (Lines, Invested_Value_Line) to only the
    selected PF_ID — critical for memory on Streamlit Cloud.
    """
    import gc
    synced = []
    try:
        xls = pd.ExcelFile(uploaded_file)
    except Exception:
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, low_memory=False)
            df.columns = [c.strip() for c in df.columns]
            if 'PF_ID' in df.columns:
                sheets.upsert_pf_level(df)
                synced.append('PF_level')
        except Exception:
            pass
        return synced

    LARGE_PER_CLIENT_TABS = {'Lines', 'Invested_Value_Line'}
    sheet_order = sorted(
        xls.sheet_names,
        key=lambda s: 1 if _TAB_ALIASES.get(s.strip().lower()) in LARGE_PER_CLIENT_TABS else 0,
    )

    for sheet_name in sheet_order:
        canonical = _TAB_ALIASES.get(sheet_name.strip().lower())
        if not canonical:
            continue
        upsert_fn = _TAB_UPSERT_MAP[canonical]
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            df.columns = [c.lstrip('\ufeff').strip() for c in df.columns]
            if df.empty:
                continue
            if canonical in LARGE_PER_CLIENT_TABS and only_pf_id and 'PF_ID' in df.columns:
                df = df[df['PF_ID'].astype(str) == str(only_pf_id)].copy()
                if df.empty:
                    continue
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].dt.strftime('%Y-%m-%d').fillna('')
            upsert_fn(df)
            synced.append(canonical)
        except Exception as e:
            st.warning(f"Could not sync tab '{sheet_name}': {e}")
        finally:
            try:
                del df
            except Exception:
                pass
            gc.collect()

    uploaded_file.seek(0)
    return synced


# ── helpers ──────────────────────────────────────────────────


def _find_col(columns, *candidates) -> str | None:
    """Case-insensitive column lookup. Returns the actual column name or None."""
    upper_map = {c.upper(): c for c in columns}
    for cand in candidates:
        if cand.upper() in upper_map:
            return upper_map[cand.upper()]
    return None


def _clients_from_upload(uploaded_file) -> list[tuple[str, str]]:
    """
    Parse an uploaded Excel/CSV and return [(pf_id, display_name), ...].
    Looks across all sheets for a PF_ID column.  NAME is used for display
    if present, otherwise falls back to PF_ID.
    """
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
            return _extract_clients(df)

        xls = pd.ExcelFile(uploaded_file)
        all_clients: dict[str, str] = {}  # pf_id → name

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name)
            df.columns = [c.strip() for c in df.columns]

            pf_col = _find_col(df.columns, "PF_ID", "PFID", "pf_id")
            if not pf_col:
                continue

            name_col = _find_col(
                df.columns, "NAME", "CLIENT_NAME", "CUSTOMER_NAME",
                "INVESTOR_NAME", "name", "Name",
            )

            for _, row in df.drop_duplicates(pf_col).iterrows():
                pid = str(row.get(pf_col, "")).strip()
                if not pid or pid.lower() == "nan":
                    continue
                name = ""
                if name_col:
                    name = str(row.get(name_col, "")).strip()
                    if name.lower() == "nan":
                        name = ""
                if pid not in all_clients or (not all_clients[pid] and name):
                    all_clients[pid] = name

        return [
            (pid, name if name else pid)
            for pid, name in sorted(all_clients.items(), key=lambda x: x[1] or x[0])
        ]
    except Exception as e:
        st.error(f"Could not parse uploaded file: {e}")
        return []


def _extract_clients(df: pd.DataFrame) -> list[tuple[str, str]]:
    """Extract (pf_id, label) pairs from a single DataFrame."""
    df.columns = [c.strip() for c in df.columns]

    pf_col = _find_col(df.columns, "PF_ID", "PFID", "pf_id")
    if not pf_col:
        return []

    name_col = _find_col(
        df.columns, "NAME", "CLIENT_NAME", "CUSTOMER_NAME",
        "INVESTOR_NAME", "name", "Name",
    )

    clients = []
    for _, row in df.drop_duplicates(pf_col).iterrows():
        pid = str(row.get(pf_col, "")).strip()
        if not pid or pid.lower() == "nan":
            continue
        name = ""
        if name_col:
            name = str(row.get(name_col, "")).strip()
            if name.lower() == "nan":
                name = ""
        label = name if name else pid
        clients.append((pid, label))

    return sorted(clients, key=lambda x: x[1])


@st.cache_data(ttl=300, show_spinner=False)
def _load_clients_from_sheets() -> list[tuple[str, str]]:
    """
    Fallback: load client list from Google Sheets.
    Returns [(pf_id, display_label), ...] sorted by name.
    """
    pf_df = sheets.read_pf_level()
    scheme_df = sheets.read_scheme_level()

    if pf_df.empty:
        return []

    # Build name map from Scheme_level (case-insensitive column lookup)
    name_map: dict[str, str] = {}
    if not scheme_df.empty:
        name_col = _find_col(
            scheme_df.columns, "NAME", "CLIENT_NAME", "CUSTOMER_NAME",
            "INVESTOR_NAME",
        )
        pf_col_s = _find_col(scheme_df.columns, "PF_ID", "PFID")
        if name_col and pf_col_s:
            for _, row in scheme_df.drop_duplicates(pf_col_s).iterrows():
                pid = str(row.get(pf_col_s, "")).strip()
                name = str(row.get(name_col, "")).strip()
                if pid and name and name.lower() not in ("nan", ""):
                    name_map[pid] = name

    # Also check PF_level itself for a name column
    if not name_map:
        name_col_pf = _find_col(
            pf_df.columns, "NAME", "CLIENT_NAME", "CUSTOMER_NAME",
            "INVESTOR_NAME",
        )
        pf_col_pf = _find_col(pf_df.columns, "PF_ID", "PFID")
        if name_col_pf and pf_col_pf:
            for _, row in pf_df.drop_duplicates(pf_col_pf).iterrows():
                pid = str(row.get(pf_col_pf, "")).strip()
                name = str(row.get(name_col_pf, "")).strip()
                if pid and name and name.lower() not in ("nan", ""):
                    name_map[pid] = name

    pf_col = _find_col(pf_df.columns, "PF_ID", "PFID") or "PF_ID"
    clients = []
    for pid in pf_df[pf_col].astype(str).unique():
        name = name_map.get(pid, "")
        if not name:
            continue  # skip PF_IDs with no client name
        clients.append((pid, name))

    return sorted(clients, key=lambda x: x[1])


def _call_apps_script(pf_id: str) -> dict:
    """
    POST to the Apps Script Web App with {"pf_id": pf_id}.
    Returns the parsed JSON response from the script.
    """
    resp = requests.post(
        M1_APPS_SCRIPT_URL,
        json={"pf_id": pf_id},
        timeout=120,
        allow_redirects=True,
    )
    resp.raise_for_status()

    try:
        data = resp.json()
    except Exception:
        raise ValueError(
            f"Apps Script returned non-JSON response "
            f"(status {resp.status_code}): {resp.text[:300]}"
        )

    if data.get("status") == "error":
        raise ValueError(f"Apps Script error: {data.get('message', 'unknown error')}")

    return data


# ── main render ──────────────────────────────────────────────


def render():
    st.header("M1 — Client Report Sheet")
    st.caption(
        "Upload a client Excel file or select from existing data "
        "to generate a personalised report as a new Google Sheet."
    )
    st.divider()

    # ── Config check ──────────────────────────────────────────
    if not M1_APPS_SCRIPT_URL:
        st.error(
            "**M1_APPS_SCRIPT_URL is not configured.**\n\n"
            "Add it to `portal/.env`:\n"
            "```\nM1_APPS_SCRIPT_URL=https://script.google.com/macros/s/.../exec\n```"
        )
        return

    # ── Primary: load clients from Google Sheets ────────────────
    clients: list[tuple[str, str]] = []
    source = ""

    with st.spinner("Loading client list..."):
        try:
            clients = _load_clients_from_sheets()
            source = "sheets"
        except Exception as e:
            st.error(f"Could not load client list from Google Sheets: {e}")

    # ── Fallback: file upload ─────────────────────────────────
    if not clients:
        st.info("No clients found in Google Sheets. Upload a client file instead.")

    uploaded = st.file_uploader(
        "Or upload client Excel / CSV",
        type=["xlsx", "xls", "csv"],
        help="Upload a file containing PF_ID (and optionally NAME) columns. Overrides the list above.",
    )

    if uploaded:
        upload_clients = _clients_from_upload(uploaded)
        if upload_clients:
            clients = upload_clients
            source = "upload"
        else:
            st.warning("No PF_ID column found in the uploaded file.")

    if not clients:
        st.warning(
            "No clients found. Upload a client Excel file above, "
            "or add client data via the Data Manager tab."
        )
        return

    # ── Client selector ───────────────────────────────────────
    labels = [label for _, label in clients]
    pf_ids = [pid for pid, _ in clients]

    if len(clients) == 1 and source == "upload":
        selected_pf_id = pf_ids[0]
        st.info(f"Client: **{labels[0]}**")
    else:
        selected_label = st.selectbox(
            "Select client",
            options=labels,
            index=0,
            help=(
                "From uploaded file."
                if source == "upload"
                else "From Google Sheets. Upload a file above for a different client."
            ),
        )
        selected_pf_id = pf_ids[labels.index(selected_label)]

    # ── Output folder info ────────────────────────────────────
    if M1_OUTPUT_FOLDER_ID:
        folder_url = f"https://drive.google.com/drive/folders/{M1_OUTPUT_FOLDER_ID}"
        st.caption(f"Reports saved to: [M1 Output folder]({folder_url})")

    st.divider()

    # ── Generate button ───────────────────────────────────────
    if st.button("Generate Report", type="primary", use_container_width=True):
        display = labels[pf_ids.index(selected_pf_id)]

        # ── Sync uploaded data to Google Sheets ──────────────
        if source == "upload" and uploaded:
            with st.spinner("Syncing uploaded data to Google Sheets..."):
                try:
                    uploaded.seek(0)
                    synced = _sync_upload_to_sheets(uploaded, only_pf_id=selected_pf_id)
                    if synced:
                        st.success(f"Synced {len(synced)} tabs: {', '.join(synced)}")
                except Exception as e:
                    st.error(f"Failed to sync uploaded data: {e}")
                    return

        with st.spinner(f"Generating report for {display}... (this can take up to 60 seconds)"):
            try:
                result = _call_apps_script(selected_pf_id)
            except requests.Timeout:
                st.error(
                    "Request timed out after 2 minutes. "
                    "The Apps Script may still be running — check the M1 output folder."
                )
                return
            except requests.HTTPError as e:
                st.error(f"HTTP error calling Apps Script: {e}")
                return
            except ValueError as e:
                st.error(str(e))
                return
            except Exception as e:
                st.error(f"Unexpected error: {e}")
                return

        sheet_url = result.get("url", "")
        sheet_title = result.get("title", "Generated Report")

        if not sheet_url:
            st.warning(
                "Report generated but no URL returned. "
                f"Check the [M1 output folder](https://drive.google.com/drive/folders/{M1_OUTPUT_FOLDER_ID})."
            )
            return

        st.success("Report generated successfully.")
        st.markdown(
            f"### [{sheet_title}]({sheet_url})",
            help="Click to open the generated Google Sheet.",
        )
        st.link_button("Open report", sheet_url, type="primary")
