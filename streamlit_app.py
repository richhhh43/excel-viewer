import time
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("Excel Viewer")

CSV_PATH = Path("data/latest.csv")

@st.cache_data(ttl=30)
def load_csv(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    return pd.read_csv(path)

if st.button("Refresh now"):
    st.cache_data.clear()
    st.rerun()

df = load_csv(CSV_PATH)

if CSV_PATH.exists():
    st.caption(f"Published: {time.strftime('%Y-%m-%d %H:%M:%S %Z', time.localtime(CSV_PATH.stat().st_mtime))}")
else:
    st.warning(f"CSV not found: {CSV_PATH.as_posix()}")

if df.empty:
    st.stop()

st.dataframe(df, hide_index=True, use_container_width=True)

