import streamlit as st
import pandas as pd
import requests
from urllib.parse import quote

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("Excel Viewer")

# --- EDIT THIS TO MATCH YOUR REPO ---
REPO = "richhhh43/excel-viewer"
BRANCH = "main"
BASE = f"https://raw.githubusercontent.com/{REPO}/{BRANCH}/data"
TS_URL = f"{BASE}/updated_at.txt"
CSV_URL = f"{BASE}/latest.csv"

# Cache fetches briefly so you don't hammer GitHub
@st.cache_data(ttl=60)
def fetch_text(url: str) -> str:
    r = requests.get(url, timeout=20)
    r.raise_for_status()
    return r.text.strip()

@st.cache_data(ttl=60)
def fetch_csv(url: str) -> pd.DataFrame:
    return pd.read_csv(url)

if st.button("Refresh now"):
    st.cache_data.clear()
    st.rerun()

published = fetch_text(TS_URL)
st.caption(f"Published: {published}")

# Use published timestamp as a cache-buster so GitHub CDN won't serve stale CSV
cachebuster = quote(published)
df = fetch_csv(f"{CSV_URL}?v={cachebuster}")

st.caption(f"Rows: {len(df):,} | Cols: {len(df.columns):,}")
st.dataframe(df, use_container_width=True, hide_index=True)
