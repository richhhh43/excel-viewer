import time
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("Excel Viewer")

CSV_PATH = Path("data/latest.csv")

@st.cache_data(ttl=30)
def load_csv():
    if not CSV_PATH.exists():
        return pd.DataFrame()
    return pd.read_csv(CSV_PATH)

if st.button("Refresh now"):
    st.cache_data.clear()
    st.rerun()

df = load_csv()

if CSV_PATH.exists():
    st.caption(f"Published: {time.strftime('%Y-%m-%d %H:%M:%S %Z', time.localtime(CSV_PATH.stat().st_mtime))}")

if df.empty:
    st.warning("No data yet. Run publish.bat to generate data/latest.csv")
    st.stop()

st.dataframe(df, hide_index=True, use_container_width=True)
