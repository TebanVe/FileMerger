"""
Step 5: Pivot builder.
Group by columns + aggregate columns (Sum, Average, Count, Min, Max). Create pivot and store in session.
"""
import streamlit as st
import pandas as pd

AGG_LABELS = {"Sum": "sum", "Average": "mean", "Count": "count", "Min": "min", "Max": "max"}


def render_pivot_engine():
    st.subheader("Pivot table")
    merged_df = st.session_state.get("merged_df")

    if merged_df is None or merged_df.empty:
        st.info("Merge data in step 4 first.")
        return

    cols = merged_df.columns.tolist()
    numeric_cols = merged_df.select_dtypes(include=["number"]).columns.tolist()

    group_by = st.multiselect("Group by", options=cols, default=cols[:1] if cols else [], key="pivot_group")
    agg_cols = st.multiselect("Columns to aggregate", options=numeric_cols or cols, key="pivot_agg_cols")

    if not group_by:
        st.caption("Select at least one group-by column.")
        return

    # Aggregation type per column (default Sum for numeric)
    agg_config = {}
    if agg_cols:
        for c in agg_cols:
            default = "Sum" if c in numeric_cols else "Count"
            agg_config[c] = st.selectbox(
                f"Aggregation for «{c}»",
                options=list(AGG_LABELS.keys()),
                index=list(AGG_LABELS.keys()).index(default) if default in AGG_LABELS else 0,
                key=f"pivot_agg_{c}",
            )

    if st.button("Create pivot", type="primary", key="pivot_create"):
        if not group_by:
            st.warning("Select at least one group-by column.")
            return
        agg_dict = {}
        for c in agg_cols or numeric_cols or cols:
            if c in group_by:
                continue
            label = agg_config.get(c, "Sum")
            agg_dict[c] = AGG_LABELS.get(label, "sum")
        if not agg_dict:
            other = [x for x in cols if x not in group_by]
            if other:
                agg_dict[other[0]] = "count"
        if not agg_dict:
            st.warning("Select at least one column to aggregate.")
            return
        try:
            pivot_df = merged_df.groupby(group_by, as_index=False).agg(agg_dict)
            st.session_state.pivot_df = pivot_df
            st.session_state.export_source_df = pivot_df
            st.success(f"Pivot created: {len(pivot_df)} rows.")
            st.rerun()
        except Exception as e:
            st.error(str(e))

    if st.session_state.get("pivot_df") is not None:
        st.dataframe(st.session_state.pivot_df.head(100), use_container_width=True, hide_index=True)
        st.caption(f"Showing first 100 of {len(st.session_state.pivot_df)} rows.")
