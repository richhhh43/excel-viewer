import streamlit as st
import pandas as pd
import requests
from io import StringIO

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("Excel Viewer")

OWNER = "richhhh43"
REPO = "excel-viewer"
BRANCH = "main"

CSV_PATH = "data/latest.csv"
TS_PATH  = "data/updated_at.txt"

API_COMMITS = f"https://api.github.com/repos/{OWNER}/{REPO}/commits"

def latest_sha_for_path(path: str) -> str:
    # Get the most recent commit that touched this file
    r = requests.get(
        API_COMMITS,
        params={"path": path, "sha": BRANCH, "per_page": 1},
        timeout=20,
        headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
    )
    r.raise_for_status()
    data = r.json()
    return data[0]["sha"]

def raw_at_sha(path: str, sha: str) -> str:
    return f"https://raw.githubusercontent.com/{OWNER}/{REPO}/{sha}/{path}"

@st.cache_data(ttl=30)
def fetch_text(url: str) -> str:
    r = requests.get(url, timeout=20, headers={"Cache-Control": "no-cache", "Pragma": "no-cache"})
    r.raise_for_status()
    return r.text.strip()

@st.cache_data(ttl=30)
def fetch_csv(url: str) -> pd.DataFrame:
    r = requests.get(url, timeout=20, headers={"Cache-Control": "no-cache", "Pragma": "no-cache"})
    r.raise_for_status()
    return pd.read_csv(StringIO(r.text))

if st.button("Refresh now"):
    st.cache_data.clear()
    st.rerun()

# Resolve immutable SHAs
ts_sha  = latest_sha_for_path(TS_PATH)
csv_sha = latest_sha_for_path(CSV_PATH)

published = fetch_text(raw_at_sha(TS_PATH, ts_sha))
st.caption(f"Published: {published}")

# Optional: show which commit you're reading from (useful proof)
st.caption(f"Data commit: {csv_sha[:8]}")

df = fetch_csv(raw_at_sha(CSV_PATH, csv_sha))
st.dataframe(df, use_container_width=True, hide_index=True)
