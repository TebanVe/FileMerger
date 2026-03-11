# Data Merge Tool (Streamlit app)

Visual, step-by-step tool to merge CSV and Excel files, fix column mismatches, pivot, explore, and export. No coding required—everything is done by clicking and selecting in the browser.

## How to start the app

1. Install dependencies from the project root: `pip install -r requirements.txt` (includes Streamlit, Plotly, PyYAML, PyArrow, and the rest).
2. Open a terminal in the **project root** (the folder that contains `app/` and `src/`).
3. Run:
   ```bash
   streamlit run app/app.py
   ```
   Or with a specific port:
   ```bash
   streamlit run app/app.py --server.port 8501
   ```
4. Your browser will open to the app (e.g. `http://localhost:8501`). If not, open that URL manually.

**Optional:** Create a shortcut (e.g. a `.bat` or `.cmd` file on Windows) that runs the command above so you can start the app without typing in a terminal.

## How to use the app (end users)

1. **Folder:** In the sidebar, paste or type the full path to the folder that contains your CSV or Excel files. Check **Include subfolders** if your files are in subfolders.
2. **Files (tab 1):** The app will list all CSV and Excel files found. If the list is empty, check the folder path.
3. **Columns (tab 2):** See which columns each file has and whether they match. Use **Dictionary (tab 3)** to map different column names to one name (e.g. "Campaign Name" → "Campaign").
4. **Dictionary (tab 3):** Add mappings so that different column names from different files are treated as the same (e.g. "Cost USD" → "Cost").
5. **Merge (tab 4):** Click **Merge files**. The app combines all files and shows the total row count. If you see a warning about Excel row limits, use Pivot or filters before exporting.
6. **Pivot (tab 5):** Optionally create a pivot table: choose **Group by** columns and which columns to aggregate (Sum, Average, Count, etc.), then click **Create pivot**.
7. **Explore (tab 6):** Filter, sort, and preview the data. The table at the bottom is the **preview of what you will export**.
8. **Export (tab 7):** Choose CSV, Excel, or Parquet and click the download button. The summary at the top shows how many rows and columns you are exporting.

You never need to type commands or edit code—only use the sidebar and tabs.
