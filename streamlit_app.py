import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("Excel Viewer")

csv_path = Path("data/latest.csv")

if not csv_path.exists():
    st.warning("No data yet. Run publish_sheet.py to generate data/latest.csv")
    st.stop()

df = pd.read_csv(csv_path)
st.caption(f"Loaded: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Rows: {len(df):,} | Cols: {len(df.columns):,}")
st.dataframe(df, use_container_width=True, hide_index=True)
