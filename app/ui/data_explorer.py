"""
Step 6: Interactive data exploration + preview before export.
Filters, sort, search; KPI cards; optional charts; "Filtered data" / "Preview of data to export" table.
"""
import streamlit as st
import pandas as pd
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    import plotly.express as px
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

PAGE_SIZE = 100


def render_data_explorer():
    st.subheader("Explore data")
    merged_df = st.session_state.get("merged_df")
    pivot_df = st.session_state.get("pivot_df")

    # Which dataset to explore
    source_options = []
    if merged_df is not None and not merged_df.empty:
        source_options.append("Merged data")
    if pivot_df is not None and not pivot_df.empty:
        source_options.append("Pivot table")

    if not source_options:
        st.info("Merge or create a pivot in steps 4–5 first.")
        return

    chosen = st.radio(
        "Explore",
        options=source_options,
        key="explore_source",
    )
    df = merged_df if chosen == "Merged data" else pivot_df
    if df is None or df.empty:
        return

    # Filters in sidebar-style (above table)
    st.markdown("**Filters**")
    filtered = df.copy()
    for col in df.columns:
        if df[col].dtype == "object" or pd.api.types.is_string_dtype(df[col]):
            uniq = df[col].dropna().unique().tolist()
            if len(uniq) <= 100:
                sel = st.multiselect(f"«{col}»", options=sorted(uniq, key=str), key=f"filter_{col}")
                if sel:
                    filtered = filtered[filtered[col].isin(sel)]
            else:
                search = st.text_input(f"Search «{col}»", key=f"search_{col}")
                if search:
                    filtered = filtered[filtered[col].astype(str).str.contains(search, case=False, na=False)]
        elif pd.api.types.is_numeric_dtype(df[col]):
            lo = float(df[col].min()) if len(df) else 0
            hi = float(df[col].max()) if len(df) else 1
            r = st.slider(f"«{col}» range", lo, hi, (lo, hi), key=f"slider_{col}")
            filtered = filtered[(filtered[col] >= r[0]) & (filtered[col] <= r[1])]

    # Sort
    sort_col = st.selectbox("Sort by", options=[""] + filtered.columns.tolist(), key="explore_sort_col")
    if sort_col:
        asc = st.checkbox("Ascending", value=True, key="explore_asc")
        filtered = filtered.sort_values(sort_col, ascending=asc)

    # Set export source to this view
    st.session_state.export_source_df = filtered

    # KPI row
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Total rows", f"{len(filtered):,}")
    with c2:
        st.metric("Columns", len(filtered.columns))
    with c3:
        st.metric("Records (after filter)", f"{len(filtered):,}")

    # Optional chart (one numeric vs one category/date)
    if HAS_PLOTLY and len(filtered) > 0 and len(filtered) <= 10_000:
        num_cols = filtered.select_dtypes(include=["number"]).columns.tolist()
        other_cols = [c for c in filtered.columns if c not in num_cols]
        if num_cols and other_cols:
            with st.expander("Chart preview"):
                x_col = st.selectbox("Chart X", options=other_cols[:20], key="chart_x")
                y_col = st.selectbox("Chart Y", options=num_cols[:20], key="chart_y")
                if x_col and y_col:
                    try:
                        fig = px.bar(
                            filtered.head(500),
                            x=x_col,
                            y=y_col,
                            title=f"{y_col} by {x_col}",
                        )
                        fig.update_layout(margin=dict(l=20, r=20, t=40, b=20))
                        st.plotly_chart(fig, use_container_width=True)
                    except Exception:
                        pass

    # Preview of data to export
    st.markdown("---")
    st.markdown("**Preview of data to export**")
    st.caption(f"Showing 1–{min(PAGE_SIZE, len(filtered))} of {len(filtered):,} rows.")
    st.dataframe(
        filtered.head(PAGE_SIZE),
        use_container_width=True,
        hide_index=True,
    )
