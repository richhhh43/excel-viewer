import streamlit as st
import pandas as pd
from pathlib import Path
import os
import time

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("Excel Viewer")

csv_path = Path("data/latest.csv")

if not csv_path.exists():
    st.warning("No data yet. Run publish_sheet.py to generate data/latest.csv")
    st.stop()

# Use file modified time to force Streamlit to refresh data when CSV changes
mtime = os.path.getmtime(csv_path)
st.caption(f"Last updated: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(mtime))}")

df = pd.read_csv(csv_path)
st.dataframe(df, use_container_width=True, hide_index=True)
