"""
Step 4: Merge files.
Apply column mapping, read all files via FileMerger, concatenate. Show progress and row count.
Warn if rows exceed Excel limit.
"""
import streamlit as st
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.file_merger import FileMerger
from app.config import EXCEL_MAX_ROWS


def render_merge_engine():
    st.subheader("Merge files")
    raw_list = st.session_state.get("file_list") or []
    file_list = [Path(p) for p in raw_list]
    column_mapping = st.session_state.get("column_mapping") or {}

    if not file_list:
        st.info("Load files in step 1 first.")
        return

    # Use selected folder as root for FileMerger
    folder_path = (st.session_state.get("folder_path") or "").strip()
    root_dir = Path(folder_path) if folder_path else Path(file_list[0]).parent
    if not root_dir.exists():
        root_dir = Path(file_list[0]).parent

    if st.button("Merge files", type="primary", key="merge_btn"):
        merger = FileMerger(
            str(root_dir),
            clean_columns=True,
            column_aliases=column_mapping,
            validate_structure=False,
        )
        dataframes = []
        progress = st.progress(0.0)
        n = len(file_list)
        for i, path in enumerate(file_list):
            df = merger.read_file(Path(path))
            if df is not None:
                dataframes.append(df)
            progress.progress((i + 1) / n)
        progress.progress(1.0)

        if not dataframes:
            st.error("Could not read any file.")
            return

        merged = merger.merge_dataframes(dataframes)
        st.session_state.merged_df = merged
        st.session_state.export_source_df = merged
        st.success(f"Merged **{len(merged)}** rows from **{len(dataframes)}** files.")

        if len(merged) >= EXCEL_MAX_ROWS:
            st.warning(
                "This dataset is larger than Excel limits (~1,048,576 rows). "
                "Consider using Pivot or filtering before export."
            )
        st.rerun()

    if st.session_state.get("merged_df") is not None:
        df = st.session_state.merged_df
        st.metric("Total rows", f"{len(df):,}")
        st.metric("Total columns", len(df.columns))
