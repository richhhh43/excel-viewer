
import time
from io import StringIO
import streamlit as st
import pandas as pd
import requests

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("Excel Viewer")

REPO = "richhhh43/excel-viewer"
BRANCH = "main"
BASE = f"https://raw.githubusercontent.com/{REPO}/{BRANCH}/data"

def u(name: str, token: str) -> str:
    return f"{BASE}/{name}?v={token}"

if "token" not in st.session_state:
    st.session_state.token = str(int(time.time()))

if st.button("Refresh now"):
    st.session_state.token = str(int(time.time()))
    st.cache_data.clear()
    st.rerun()

@st.cache_data(ttl=60)
def fetch_text(url: str) -> str:
    r = requests.get(url, timeout=20, headers={"Cache-Control": "no-cache"})
    r.raise_for_status()
    return r.text.strip()

@st.cache_data(ttl=60)
def fetch_csv(url: str) -> pd.DataFrame:
    r = requests.get(url, timeout=20, headers={"Cache-Control": "no-cache"})
    r.raise_for_status()
    return pd.read_csv(StringIO(r.text))

token = st.session_state.token

published = fetch_text(u("updated_at.txt", token))
st.caption(f"Published: {published}")

df = fetch_csv(u("latest.csv", token))
st.caption(f"Rows: {len(df):,} | Cols: {len(df.columns):,}")
st.dataframe(df, use_container_width=True, hide_index=True)
