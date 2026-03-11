"""
Step 1: File browser / folder selection.
Detect CSV and Excel files; show count, types, sizes. Option: include subfolders.
"""
import streamlit as st
from pathlib import Path

from app.utils import get_supported_files_from_folder


def render_file_loader():
    st.subheader("Files")
    folder = (st.session_state.get("folder_path") or "").strip()
    include_subfolders = st.session_state.get("include_subfolders", False)

    if not folder:
        st.info("Enter a folder path in the sidebar, then we'll scan for CSV and Excel files here.")
        return

    path = Path(folder)
    if not path.exists():
        st.error(f"Folder does not exist: {folder}")
        return
    if not path.is_dir():
        st.error(f"Path is not a directory: {folder}")
        return

    files = get_supported_files_from_folder(path, include_subfolders)

    if not files:
        st.warning("No CSV or Excel files found in this folder.")
        st.session_state.file_list = []
        return

    # Store as list of Paths (Streamlit may serialize; normalize on use elsewhere)
    st.session_state.file_list = [Path(p) for p in files]
    folder_name = path.name or path.as_posix()

    st.success(f"**Files detected:** {folder_name}")
    st.caption(f"Total: {len(files)} file(s)" + (" (including subfolders)" if include_subfolders else ""))

    # File list with size and type
    st.markdown("**File list**")
    rows = []
    for p in files:
        try:
            size = p.stat().st_size
            size_kb = size / 1024
            if size_kb >= 1024:
                size_str = f"{size_kb / 1024:.1f} MB"
            else:
                size_str = f"{size_kb:.1f} KB"
        except OSError:
            size_str = "—"
        rows.append((p.name, p.suffix.lower(), size_str))

    import pandas as pd
    df_display = pd.DataFrame(rows, columns=["File name", "Type", "Size"])
    st.dataframe(df_display, use_container_width=True, hide_index=True)
