"""
tabs/data_manager.py — Data Manager tab (Tab 4).

Flow:
  1. User uploads a CSV or Excel file
  2. Portal auto-detects which dataset(s) it is from column headers
  3. Simple pass/fail shown per dataset
  4. User clicks Upload — data written to Google Sheets
"""

import io
import pandas as pd
import streamlit as st

import sheets
from config import MainSheets, M3Sheets, TimeSeriesSheets

# ── Human-readable labels for each dataset ────────────────────
_TAB_LABELS = {
    MainSheets.PF_LEVEL:            "PF Level",
    MainSheets.SCHEME_LEVEL:        "Scheme Level",
    MainSheets.RISKGROUP_LEVEL:     "Risk Group Level",
    MainSheets.LINES:               "Lines",
    MainSheets.RESULTS:             "Results",
    MainSheets.INVESTED_VALUE_LINE: "Invested Value Line",
    MainSheets.SCHEME_CATEGORY:     "Scheme Category",
    M3Sheets.AUM:                   "AUM",
    M3Sheets.POWERRANKING:          "Power Ranking",
    M3Sheets.UPSIDE_DOWNSIDE:       "Upside / Downside",
    M3Sheets.ROLLING_RETURNS:       "Rolling Returns",
}

_M3_FULL_REPLACE = {M3Sheets.AUM, M3Sheets.POWERRANKING,
                    M3Sheets.UPSIDE_DOWNSIDE, M3Sheets.ROLLING_RETURNS}

_TIMESERIES_TABS = {TimeSeriesSheets.LINES, TimeSeriesSheets.INVESTED_VALUE_LINE}


def _load_file(uploaded) -> dict[str, pd.DataFrame]:
    """Parse an uploaded file into {sheet_name: DataFrame}."""
    name = uploaded.name.lower()
    raw  = uploaded.read()

    if name.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(raw), low_memory=False, encoding="utf-8-sig")
        df.columns = [c.strip().lstrip("\ufeff") for c in df.columns]
        return {uploaded.name: df}

    xf = pd.ExcelFile(io.BytesIO(raw))
    result = {}
    for sheet in xf.sheet_names:
        df = xf.parse(sheet)
        df.columns = [str(c).strip() for c in df.columns]
        result[sheet] = df
    return result


def render():
    st.header("Data Manager")
    st.caption("Upload a CSV or Excel file to refresh shared data in Google Sheets.")
    st.divider()

    uploaded = st.file_uploader(
        "Upload file",
        type=["csv", "xlsx", "xls"],
        help="CSV or Excel. Multi-sheet Excel files are processed automatically.",
    )

    if uploaded is None:
        return

    # ── Parse ─────────────────────────────────────────────────
    with st.spinner("Reading file..."):
        try:
            sheets_in_file = _load_file(uploaded)
        except Exception as e:
            st.error(f"Could not read file: {e}")
            return

    # ── Detect each sheet ─────────────────────────────────────
    detected: list[dict] = []

    for sheet_name, df in sheets_in_file.items():
        if df.empty:
            continue
        result = sheets.detect_dataset(df)
        if result is None:
            continue
        tab, fn, key_cols = result
        detected.append({
            "sheet_name": sheet_name,
            "label":      _TAB_LABELS.get(tab, tab),
            "tab":        tab,
            "df":         df,
            "fn":         fn,
            "key_cols":   key_cols,
            "full_replace": tab in _M3_FULL_REPLACE,
            "is_timeseries": tab in _TIMESERIES_TABS,
        })

    if not detected:
        st.error("No recognisable datasets found in this file.")
        return

    # ── Simple summary ────────────────────────────────────────
    st.success(f"Detected **{len(detected)}** dataset(s). Ready to upload.")

    for item in detected:
        df = item["df"]
        mode = "Full replace" if item["full_replace"] else "Upsert"
        icon = ":white_check_mark:"
        st.markdown(
            f"{icon} **{item['label']}** — {len(df):,} rows — {mode}"
        )

    st.divider()

    # ── Upload button ─────────────────────────────────────────
    if st.button("Upload to Google Sheets", type="primary", use_container_width=True):
        progress = st.progress(0, text="Writing...")
        errors = []

        for i, item in enumerate(detected):
            progress.progress(
                i / len(detected),
                text=f"Writing {item['label']}..."
            )
            try:
                result = item["fn"](item["df"])
                if isinstance(result, dict):
                    msg = f"**{item['label']}** — {result.get('added', 0)} added, {result.get('replaced', 0)} replaced"
                    pruned = result.get("pruned", 0)
                    if pruned:
                        msg += f", {pruned:,} old rows pruned"
                    st.success(msg)
                else:
                    st.success(f"**{item['label']}** — {result} rows written")
            except Exception as e:
                errors.append((item["label"], str(e)))
                st.error(f"**{item['label']}** failed: {e}")

        progress.progress(1.0, text="Done.")

        if not errors:
            st.balloons()
        else:
            st.warning(f"{len(errors)} dataset(s) failed — check errors above.")
