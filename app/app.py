"""
Streamlit Visual Data Tool – Entry point.
Run from repo root: streamlit run app/app.py

No modifications to existing src/ code; this app only imports from it.
"""
import sys
from pathlib import Path

# Ensure repo root is on path so we can import src
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import streamlit as st

from app.config import APP_TITLE
from app.ui import (
    render_file_loader,
    render_column_analyzer,
    render_dictionary_manager,
    render_merge_engine,
    render_pivot_engine,
    render_data_explorer,
    render_export_tools,
)

# Page config (must be first Streamlit command)
st.set_page_config(
    page_title=APP_TITLE,
    layout="wide",
    initial_sidebar_state="expanded",
)

# Optional dashboard CSS
try:
    css_path = Path(__file__).parent / "assets" / "style.css"
    if css_path.exists():
        st.markdown(f"<style>{css_path.read_text()}</style>", unsafe_allow_html=True)
except Exception:
    pass

# Session state initialization
def _init_session_state():
    if "folder_path" not in st.session_state:
        st.session_state.folder_path = ""
    if "include_subfolders" not in st.session_state:
        st.session_state.include_subfolders = False
    if "file_list" not in st.session_state:
        st.session_state.file_list = []
    if "file_headers" not in st.session_state:
        st.session_state.file_headers = {}  # path key -> list of column names
    if "column_mapping" not in st.session_state:
        st.session_state.column_mapping = {}
    if "merged_df" not in st.session_state:
        st.session_state.merged_df = None
    if "pivot_df" not in st.session_state:
        st.session_state.pivot_df = None
    if "export_source_df" not in st.session_state:
        st.session_state.export_source_df = None  # current view to export (explored or merged/pivot)
    if "active_tab" not in st.session_state:
        st.session_state.active_tab = 0

_init_session_state()

# Sidebar: step navigation and folder
st.sidebar.title(APP_TITLE)
st.sidebar.markdown("---")
st.sidebar.subheader("Steps")
steps = [
    "1. Files",
    "2. Columns",
    "3. Dictionary",
    "4. Merge",
    "5. Pivot",
    "6. Explore",
    "7. Export",
]
# Status indicators (simple)
file_ready = len(st.session_state.get("file_list", [])) > 0
cols_ready = file_ready and len(st.session_state.get("file_headers", {})) > 0
merge_ready = st.session_state.get("merged_df") is not None
for i, label in enumerate(steps):
    if i == 0:
        status = "Ready" if file_ready else "—"
    elif i == 1:
        status = "Ready" if cols_ready else ("—" if not file_ready else "Open")
    elif i == 2:
        status = "Ready" if file_ready else "—"
    elif i == 3:
        status = "Ready" if merge_ready and file_ready else "—"
    elif i == 4:
        status = "Ready" if merge_ready else "—"
    elif i == 5:
        status = "Ready" if merge_ready else "—"
    elif i == 6:
        status = "Ready" if (merge_ready or st.session_state.get("pivot_df") is not None) else "—"
    else:
        status = "Ready" if st.session_state.get("export_source_df") is not None or merge_ready else "—"
    st.sidebar.markdown(f"**{label}** — {status}")

st.sidebar.markdown("---")
st.sidebar.subheader("Folder")
folder_path = st.sidebar.text_input(
    "Paste or type folder path here",
    value=st.session_state.folder_path,
    key="sidebar_folder_path",
)
if folder_path != st.session_state.folder_path:
    st.session_state.folder_path = folder_path
    st.session_state.file_list = []
    st.session_state.file_headers = {}

include_subfolders = st.sidebar.checkbox(
    "Include subfolders",
    value=st.session_state.include_subfolders,
    key="sidebar_include_subfolders",
)
if include_subfolders != st.session_state.include_subfolders:
    st.session_state.include_subfolders = include_subfolders
    st.session_state.file_list = []
    st.session_state.file_headers = {}

st.sidebar.markdown("---")
if st.sidebar.button("Go to Export"):
    st.session_state.active_tab = 6
    st.rerun()

# Main area: tabbed workflow
tab_names = ["1. Files", "2. Columns", "3. Dictionary", "4. Merge", "5. Pivot", "6. Explore", "7. Export"]
tabs = st.tabs(tab_names)

with tabs[0]:
    render_file_loader()

with tabs[1]:
    render_column_analyzer()

with tabs[2]:
    render_dictionary_manager()

with tabs[3]:
    render_merge_engine()

with tabs[4]:
    render_pivot_engine()

with tabs[5]:
    render_data_explorer()

with tabs[6]:
    render_export_tools()
