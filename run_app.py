"""
Launcher for the Data Merge Tool (Streamlit app).
Run this from the FileMerger folder:  python run_app.py

This sets the project folder correctly so the app can find all its files.
"""
import sys
import os
from pathlib import Path

# Project root = folder that contains app/, src/, run_app.py
ROOT = Path(__file__).resolve().parent
os.chdir(ROOT)
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# Start Streamlit with app/app.py
import streamlit.web.cli as stcli
sys.argv = ["streamlit", "run", str(ROOT / "app" / "app.py"), "--server.headless", "true"]
stcli.main()
