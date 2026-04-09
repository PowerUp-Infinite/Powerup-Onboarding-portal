"""
app.py — PowerUp Portal, main Streamlit entry point.

Run with:
    streamlit run portal/app.py

Four tabs:
  Tab 1  M1 · Client Report Sheet   — trigger Apps Script, get Google Sheet link
  Tab 2  M2 · Client Strategy Deck  — select client, generate Google Slides deck
  Tab 3  M3 · Transition Deck       — upload client Excel, generate Google Slides deck
  Tab 4  Data Manager               — upload + deduplicate + write shared data to Sheets
"""

import streamlit as st
from dotenv import load_dotenv

load_dotenv()

from google_auth import check_auth
from tabs import data_manager, m1_report, m2_deck, m3_deck
import sheets

# ─────────────────────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PowerUp Portal",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────────────────────
st.title("⚡ PowerUp Portal")
st.caption("Internal tool — PowerUp Infinite · Wealth Management")

# ─────────────────────────────────────────────────────────────
# Sidebar — auth status
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Configuration")
    auth_ok, auth_msg = check_auth()
    if auth_ok:
        st.success("Google auth ✓")
    else:
        st.error("Google auth ✗")
        st.caption(auth_msg)
    st.divider()
    st.caption("PowerUp Infinite — Internal use only")

# ─────────────────────────────────────────────────────────────
# Auto-cleanup old M1/M2 outputs (once per session)
# ─────────────────────────────────────────────────────────────
if "cleanup_done" not in st.session_state:
    try:
        deleted = sheets.cleanup_old_outputs(max_age_days=7)
        total = deleted["m1"] + deleted["m2"]
        if total:
            st.toast(f"Cleaned up {total} old report(s) (M1: {deleted['m1']}, M2: {deleted['m2']})")
        st.session_state.cleanup_done = True
    except Exception:
        st.session_state.cleanup_done = True

# ─────────────────────────────────────────────────────────────
# Tabs
# ─────────────────────────────────────────────────────────────
tab_m1, tab_m2, tab_m3, tab_data = st.tabs([
    "📋  M1 · Client Report",
    "📊  M2 · Strategy Deck",
    "🔄  M3 · Transition Deck",
    "🗄️  Data Manager",
])

# ── Tab 1 ─────────────────────────────────────────────────────
with tab_m1:
    m1_report.render()

# ── Tab 2 ─────────────────────────────────────────────────────
with tab_m2:
    m2_deck.render()

# ── Tab 3 ─────────────────────────────────────────────────────
with tab_m3:
    m3_deck.render()

# ── Tab 4 ─────────────────────────────────────────────────────
with tab_data:
    data_manager.render()
