"""
Step 2: Column inspection.
Scan each file for column names; show column x file matrix and match/mismatch.
Uses structure_validator for validation; header-only reads where possible.
"""
import streamlit as st
from pathlib import Path
import pandas as pd
import sys

# Repo root on path (app.py adds it)
ROOT = Path(__file__).resolve().parent.parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.structure_validator import validate_column_structure, StructureValidationResult


def _get_columns_only(file_path: Path) -> list:
    """Read only column names (header) to save memory. Uses pandas directly."""
    suf = file_path.suffix.lower()
    try:
        if suf == ".csv":
            df = pd.read_csv(file_path, nrows=0, encoding="utf-8", on_bad_lines="skip")
        elif suf in (".xlsx", ".xls"):
            df = pd.read_excel(file_path, engine="openpyxl" if suf == ".xlsx" else "xlrd", header=0)
            df = df.head(0)
        else:
            return []
        return [str(c) for c in df.columns]
    except Exception:
        for enc in ("latin-1", "windows-1252"):
            try:
                if suf == ".csv":
                    df = pd.read_csv(file_path, nrows=0, encoding=enc, on_bad_lines="skip")
                    return [str(c) for c in df.columns]
            except Exception:
                continue
    return []


def render_column_analyzer():
    st.subheader("Columns")
    raw_list = st.session_state.get("file_list") or []
    file_list = [Path(p) for p in raw_list]
    column_mapping = st.session_state.get("column_mapping") or {}

    if not file_list:
        st.info("Select a folder and load files in step 1 first.")
        return

    # Get headers for each file (cached in session). Use resolved path as key.
    def _path_key(p):
        return str(Path(p).resolve())
    file_headers = st.session_state.get("file_headers") or {}
    need_scan = not all(_path_key(p) in file_headers for p in file_list)
    if need_scan:
        with st.spinner("Scanning column names…"):
            for p in file_list:
                key = _path_key(p)
                if key not in file_headers:
                    file_headers[key] = _get_columns_only(Path(p))
        st.session_state.file_headers = file_headers

    # Apply mapping for display (variant -> canonical)
    def canonical(name: str) -> str:
        return column_mapping.get(name, name)

    all_columns = set()
    for cols in file_headers.values():
        for c in cols:
            all_columns.add(canonical(c))

    if not all_columns:
        st.warning("No columns found in the scanned files.")
        return

    # Build column x file matrix
    file_names = [p.name for p in file_list]
    rows = []
    for col in sorted(all_columns):
        # Which files have this column (after mapping)?
        row = {"Column name": col}
        match = True
        for p in file_list:
            cols = file_headers.get(_path_key(p), [])
            present = any(canonical(c) == col for c in cols)
            row[p.name] = "Yes" if present else "No"
            if not present:
                match = False
        row["Match"] = "Yes" if match else "No"
        rows.append(row)

    df_matrix = pd.DataFrame(rows)
    st.dataframe(df_matrix, use_container_width=True, hide_index=True)

    # Validation using structure_validator (with header-only DataFrames)
    dfs_for_validation = []
    for p in file_list:
        cols = file_headers.get(_path_key(p), [])
        dfs_for_validation.append(pd.DataFrame(columns=cols))

    result = validate_column_structure(
        [Path(p) for p in file_list],
        dfs_for_validation,
        allow_missing_columns=True,
    )

    if not result.success:
        st.warning("**Columns do not match across files.** Use the Dictionary tab to map them.")

    # Preview first rows of a selected file
    st.markdown("---")
    st.caption("Preview first rows of a file")
    chosen = st.selectbox("Choose file", options=[p.name for p in file_list], key="col_preview_file")
    if chosen:
        path = next(p for p in file_list if p.name == chosen)
        try:
            if path.suffix.lower() == ".csv":
                preview = pd.read_csv(path, nrows=5, encoding="utf-8", on_bad_lines="skip")
            else:
                preview = pd.read_excel(path, engine="openpyxl", header=0).head(5)
            st.dataframe(preview, use_container_width=True, hide_index=True)
        except Exception as e:
            st.caption(f"Could not preview: {e}")
