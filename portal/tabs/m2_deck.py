"""
tabs/m2_deck.py — M2 Client Strategy Deck tab.

Flow:
  1. Load client list from Google Sheets (primary) or uploaded Excel (override)
  2. User selects a client (name-based dropdown)
  3. User confirms/selects matching questionnaire response
  4. "Generate Deck" → m2_engine builds PPTX in memory
  5. PPTX uploaded to Google Drive (M2_OUTPUT_FOLDER_ID), converted to Slides
  6. Portal shows clickable link to the generated presentation
"""

from difflib import SequenceMatcher

import pandas as pd
import streamlit as st

import sheets
from config import M2_OUTPUT_FOLDER_ID
from m2_engine import generate_deck, load_data


# ── Tab name → upsert function mapping ──────────────────────
_TAB_UPSERT_MAP = {
    'PF_level':            sheets.upsert_pf_level,
    'Scheme_level':        sheets.upsert_scheme_level,
    'Riskgroup_level':     sheets.upsert_riskgroup_level,
    'Results':             sheets.upsert_results,
    'Lines':               sheets.upsert_lines,
    'Invested_Value_Line': sheets.upsert_invested_value_line,
}

# Case-insensitive lookup for common tab name variations
_TAB_ALIASES = {}
for _canonical in _TAB_UPSERT_MAP:
    _TAB_ALIASES[_canonical.lower()] = _canonical
    _TAB_ALIASES[_canonical.lower().replace('_', '')] = _canonical
    _TAB_ALIASES[_canonical.lower().replace('_', ' ')] = _canonical


def _sync_upload_to_sheets(uploaded_file, only_pf_id: str | None = None) -> list[str]:
    """
    Parse the uploaded Excel and upsert each recognised tab into Google Sheets.
    If only_pf_id is provided, large per-client tabs (Lines, Invested_Value_Line)
    are filtered to that PF_ID only — critical for memory on Streamlit Cloud.
    Returns list of tab names that were synced.
    """
    import gc
    synced = []
    try:
        xls = pd.ExcelFile(uploaded_file)
    except Exception:
        # Single-sheet CSV — try PF_level as a guess
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

    # Sync small tabs first, large per-client tabs last
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
            # For large per-client tabs, filter to only the selected PF_ID
            if canonical in LARGE_PER_CLIENT_TABS and only_pf_id and 'PF_ID' in df.columns:
                df = df[df['PF_ID'].astype(str) == str(only_pf_id)].copy()
                if df.empty:
                    continue
            # Convert Timestamp/datetime columns to strings for JSON serialization
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].dt.strftime('%Y-%m-%d').fillna('')
            upsert_fn(df)
            synced.append(canonical)
        except Exception as e:
            st.warning(f"Could not sync tab '{sheet_name}': {e}")
        finally:
            # Free memory between tabs (critical on Streamlit Cloud)
            try:
                del df
            except Exception:
                pass
            gc.collect()

    # Reset file pointer so _clients_from_upload can re-read it
    uploaded_file.seek(0)
    return synced


# ── helpers ──────────────────────────────────────────────────


def _find_col(columns, *candidates) -> str | None:
    """Case-insensitive column lookup."""
    upper_map = {c.upper(): c for c in columns}
    for cand in candidates:
        if cand.upper() in upper_map:
            return upper_map[cand.upper()]
    return None


def _clients_from_upload(uploaded_file) -> list[tuple[str, str]]:
    """
    Parse uploaded Excel/CSV → [(pf_id, display_name), ...].
    Scans all sheets for PF_ID column; uses NAME for display if present.
    """
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
            return _extract_clients(df)

        xls = pd.ExcelFile(uploaded_file)
        all_clients: dict[str, str] = {}

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
def _load_questionnaire_names() -> list[str]:
    """Load all respondent names from the questionnaire sheet."""
    try:
        qdf = sheets.read_questionnaire()
        if qdf.empty:
            return []
        name_col = next((c for c in qdf.columns if c.lower() == "name"), None)
        if not name_col:
            return []
        names = qdf[name_col].astype(str).str.strip().tolist()
        return [n for n in names if n and n.lower() not in ("nan", "")]
    except Exception:
        return []


def _best_match(target: str, choices: list[str]) -> int:
    """Return index of best fuzzy match from choices, or 0 if no good match."""
    if not choices or not target:
        return 0
    target_l = target.lower().strip()
    best_idx, best_score = 0, 0.0
    for i, c in enumerate(choices):
        score = SequenceMatcher(None, target_l, c.lower().strip()).ratio()
        if score > best_score:
            best_score = score
            best_idx = i
    return best_idx if best_score > 0.3 else 0


@st.cache_data(ttl=300, show_spinner=False)
def _load_clients_from_sheets() -> list[tuple[str, str]]:
    """
    Load client list from Google Sheets.
    Returns [(pf_id, display_label), ...] sorted by name.
    """
    pf_df = sheets.read_pf_level()
    scheme_df = sheets.read_scheme_level()

    if pf_df.empty:
        return []

    # Build name map from Scheme_level first, then PF_level
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


# ── main render ──────────────────────────────────────────────


def render():
    st.header("M2 — Client Strategy Deck")
    st.caption(
        "Select a client to generate their personalised strategy "
        "presentation as a Google Slides deck."
    )
    st.divider()

    # ── Config check ──────────────────────────────────────────
    if not M2_OUTPUT_FOLDER_ID:
        st.error(
            "**M2_OUTPUT_FOLDER_ID is not configured.**\n\n"
            "Add it to `portal/.env`:\n"
            "```\nM2_OUTPUT_FOLDER_ID=your_folder_id_here\n```"
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
        key="m2_upload",
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
        selected_name = labels[0]
        st.info(f"Client: **{labels[0]}**")
    else:
        selected_label = st.selectbox(
            "Select client",
            options=labels,
            index=0,
            key="m2_client",
            help=(
                "From uploaded file."
                if source == "upload"
                else "From Google Sheets. Upload a file above for a different client."
            ),
        )
        selected_pf_id = pf_ids[labels.index(selected_label)]
        selected_name = selected_label

    # ── Questionnaire matching ────────────────────────────────
    q_names = _load_questionnaire_names()
    saved_mapping = sheets.read_pf_id_mapping()
    saved_q_name = saved_mapping.get(selected_pf_id)

    questionnaire_name = None

    if q_names:
        options = sorted(set(q_names))

        # Pre-select: saved mapping first, then fuzzy match on client name
        if saved_q_name and saved_q_name in options:
            default_idx = options.index(saved_q_name)
        else:
            default_idx = _best_match(selected_name, options)

        questionnaire_name = st.selectbox(
            "Questionnaire response",
            options=options,
            index=default_idx,
            key=f"m2_questionnaire_{selected_pf_id}",
            help="Pre-filled with best match. Click and type to search.",
        )
    else:
        st.caption("No questionnaire responses found.")

    # ── Output folder info ────────────────────────────────────
    folder_url = f"https://drive.google.com/drive/folders/{M2_OUTPUT_FOLDER_ID}"
    st.caption(f"Decks saved to: [M2 Output folder]({folder_url})")

    st.divider()

    # ── Generate button ───────────────────────────────────────
    if st.button("Generate Deck", type="primary", use_container_width=True, key="m2_generate"):
        display = selected_name
        customer_name = display if display != selected_pf_id else "Client"

        # Save the questionnaire mapping if user selected one
        if questionnaire_name and questionnaire_name != saved_q_name:
            try:
                sheets.save_pf_id_mapping(selected_pf_id, questionnaire_name)
            except Exception:
                pass  # non-critical — don't block generation

        # ── Sync uploaded data to Google Sheets ──────────────
        if source == "upload" and uploaded:
            with st.spinner("Syncing uploaded data to Google Sheets..."):
                try:
                    uploaded.seek(0)
                    synced = _sync_upload_to_sheets(uploaded, only_pf_id=selected_pf_id)
                    if synced:
                        st.success(f"Synced {len(synced)} tabs: {', '.join(synced)}")
                        # Clear cached data so load_data picks up new rows
                        _load_clients_from_sheets.clear()
                    else:
                        st.warning(
                            "No recognised data tabs found in the uploaded file. "
                            "Expected tabs: PF_level, Scheme_level, Riskgroup_level, "
                            "Results, Lines, Invested_Value_Line"
                        )
                except Exception as e:
                    st.error(f"Failed to sync uploaded data: {e}")
                    return

        with st.spinner(f"Loading data for {display}..."):
            try:
                data = load_data()
            except Exception as e:
                st.error(f"Failed to load data from Google Sheets: {e}")
                return

        with st.spinner(f"Generating strategy deck for {display}... (this can take up to 2 minutes)"):
            try:
                buf, filename = generate_deck(
                    selected_pf_id, customer_name, data=data,
                    questionnaire_name=questionnaire_name,
                )
            except Exception as e:
                st.error(f"Deck generation failed: {e}")
                return

        with st.spinner("Uploading to Google Drive..."):
            try:
                result = sheets.upload_pptx_to_drive(
                    buf, filename, M2_OUTPUT_FOLDER_ID, convert_to_slides=True,
                )
            except Exception as e:
                st.error(f"Upload to Drive failed: {e}")
                # Offer local download as fallback
                buf.seek(0)
                st.download_button(
                    "Download PPTX locally",
                    data=buf,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
                return

        slide_url = result["url"]
        slide_name = result["name"]

        st.success("Strategy deck generated and uploaded successfully.")
        st.markdown(
            f"### [{slide_name}]({slide_url})",
            help="Click to open the generated Google Slides presentation.",
        )
        st.link_button("Open deck", slide_url, type="primary")
