"""
Step 7: Export. Show summary (rows, columns) and optional preview; CSV, Excel, optional Parquet.
"""
import streamlit as st
import pandas as pd
import io
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    import pyarrow as pa
    import pyarrow.parquet as pq
    HAS_PARQUET = True
except ImportError:
    HAS_PARQUET = False


def _format_size(num_bytes: float) -> str:
    if num_bytes >= 1024 * 1024:
        return f"{num_bytes / (1024*1024):.1f} MB"
    if num_bytes >= 1024:
        return f"{num_bytes / 1024:.1f} KB"
    return f"{num_bytes:.0f} B"


def render_export_tools():
    st.subheader("Export")
    export_df = st.session_state.get("export_source_df")
    merged_df = st.session_state.get("merged_df")
    pivot_df = st.session_state.get("pivot_df")

    if export_df is None:
        if merged_df is not None:
            export_df = merged_df
            st.session_state.export_source_df = merged_df
        elif pivot_df is not None:
            export_df = pivot_df
            st.session_state.export_source_df = pivot_df

    if export_df is None or export_df.empty:
        st.info("Merge or create a pivot first, or use Explore to choose the data to export.")
        return

    # Summary
    st.markdown("**You are about to export**")
    c1, c2 = st.columns(2)
    with c1:
        st.metric("Rows", f"{len(export_df):,}")
    with c2:
        st.metric("Columns", len(export_df.columns))

    try:
        approx = export_df.memory_usage(deep=True).sum()
        st.caption(f"Approximate size: {_format_size(approx)}")
    except Exception:
        pass

    # Small preview
    with st.expander("This is what you will download (first 10 rows)"):
        st.dataframe(export_df.head(10), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("**Download**")

    # CSV
    buf_csv = io.StringIO()
    export_df.to_csv(buf_csv, index=False)
    csv_bytes = buf_csv.getvalue().encode("utf-8")
    st.download_button(
        label="Export as CSV",
        data=csv_bytes,
        file_name="exported_data.csv",
        mime="text/csv",
        key="export_csv",
    )

    # Excel (avoid for huge rows)
    if len(export_df) < 1_048_576:
        buf_xlsx = io.BytesIO()
        export_df.to_excel(buf_xlsx, index=False, engine="openpyxl")
        buf_xlsx.seek(0)
        st.download_button(
            label="Export as Excel",
            data=buf_xlsx,
            file_name="exported_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="export_xlsx",
        )
    else:
        st.caption("Excel export skipped (row count exceeds Excel limit). Use CSV or Parquet.")

    if HAS_PARQUET:
        buf_pq = io.BytesIO()
        export_df.to_parquet(buf_pq, index=False)
        buf_pq.seek(0)
        st.download_button(
            label="Export as Parquet",
            data=buf_pq,
            file_name="exported_data.parquet",
            mime="application/octet-stream",
            key="export_parquet",
        )
