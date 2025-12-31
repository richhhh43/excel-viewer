import streamlit as st
import pandas as pd
from pathlib import Path

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("Excel Viewer")

csv_path = Path("data/latest.csv")
ts_path = Path("data/updated_at.txt")

if not csv_path.exists():
    st.warning("No data yet. Run publish_sheet.py to generate data/latest.csv")
    st.stop()

if ts_path.exists():
    st.caption(f"Published: {ts_path.read_text(encoding='utf-8').strip()}")

df = pd.read_csv(csv_path)
st.dataframe(df, use_container_width=True, hide_index=True)
