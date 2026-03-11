"""
App configuration constants.
Used by the Streamlit visual data tool (no changes to existing src/ code).
"""

# Excel row limit; show warning when merged data exceeds this
EXCEL_MAX_ROWS = 1_048_576

# Supported file extensions for scan (must match file_merger.py)
SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".csv")

# Default page title
APP_TITLE = "Data Merge Tool"
