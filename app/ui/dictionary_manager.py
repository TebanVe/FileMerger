"""
Step 3: Column dictionary (visual mapping).
Map source column names to canonical names; persist as YAML-style mapping (session + optional download).
"""
import streamlit as st
import yaml
import io
from pathlib import Path


def render_dictionary_manager():
    st.subheader("Column dictionary")
    file_headers = st.session_state.get("file_headers") or {}
    column_mapping = st.session_state.get("column_mapping") or {}

    all_sources = set()
    for cols in file_headers.values():
        all_sources.update(cols)

    if not all_sources:
        st.info("Load files and open the Columns step first so we know which columns exist.")
        return

    st.caption("Map different column names to one canonical name so they merge correctly.")

    # Current mappings table
    if column_mapping:
        st.markdown("**Current mappings**")
        rows = [{"Source column": k, "Map to (canonical)": v} for k, v in column_mapping.items()]
        import pandas as pd
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        if st.button("Clear all mappings", key="dict_clear"):
            st.session_state.column_mapping = {}
            st.rerun()

    # Add or edit mapping
    st.markdown("---")
    st.markdown("**Add or change mapping**")
    source_col = st.selectbox(
        "Source column",
        options=sorted(all_sources),
        key="dict_source",
    )
    existing_canonical = list(set(column_mapping.values()) | {source_col})
    canonical_name = st.text_input(
        "Map to (canonical name)",
        value=column_mapping.get(source_col, source_col),
        key="dict_canonical",
    )
    if st.button("Apply mapping", key="dict_apply"):
        if source_col and canonical_name:
            st.session_state.column_mapping = st.session_state.get("column_mapping") or {}
            st.session_state.column_mapping[source_col] = canonical_name.strip()
            st.rerun()

    # Advanced: view / download YAML
    with st.expander("Advanced: View / download YAML"):
        yaml_data = {"column_aliases": column_mapping} if column_mapping else {}
        yaml_str = yaml.dump(yaml_data, default_flow_style=False, allow_unicode=True)
        st.code(yaml_str, language="yaml")
        st.download_button(
            label="Download YAML",
            data=yaml_str,
            file_name="column_aliases.yaml",
            mime="application/x-yaml",
            key="dict_download_yaml",
        )
