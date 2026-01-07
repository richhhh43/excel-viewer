import time
from io import StringIO

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("Excel Viewer")

# ========= EDIT THESE (your repo) =========
OWNER = "richhhh43"
REPO = "excel-viewer"
BRANCH = "main"

CSV_PATH = "data/latest.csv"
TS_PATH = "data/updated_at.txt"
# =========================================

API_COMMITS = f"https://api.github.com/repos/{OWNER}/{REPO}/commits"


def raw_at(ref: str, path: str) -> str:
    return f"https://raw.githubusercontent.com/{OWNER}/{REPO}/{ref}/{path}"


@st.cache_data(ttl=30)
def latest_sha_for_path(path: str) -> str:
    """
    Returns the most recent commit SHA that touched a file path.
    Using SHA makes the raw URL immutable -> no stale caching.
    """
    r = requests.get(
        API_COMMITS,
        params={"path": path, "sha": BRANCH, "per_page": 1},
        timeout=20,
        headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
    )
    r.raise_for_status()
    data = r.json()
    return data[0]["sha"]


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


# ----- Refresh control -----
if "refresh_token" not in st.session_state:
    st.session_state.refresh_token = str(int(time.time()))

if st.button("Refresh now"):
    st.cache_data.clear()
    st.session_state.refresh_token = str(int(time.time()))
    st.rerun()


def normalize_and_format_percents(df: pd.DataFrame, cols: list) -> pd.DataFrame:
    """
    Ensures percent columns are numeric and in 0-1 range.
    Handles both 0.6351 and 63.51 inputs.
    """
    for c in cols:
        s = pd.to_numeric(df[c], errors="coerce")
        if s.dropna().size and s.dropna().median() > 1:
            df[c] = s / 100.0
        else:
            df[c] = s
    return df


# ----- Load latest data from immutable commit SHAs -----
try:
    ts_sha = latest_sha_for_path(TS_PATH)
    csv_sha = latest_sha_for_path(CSV_PATH)

    published = fetch_text(raw_at(ts_sha, TS_PATH))
    st.caption(f"Published: {published}")
    st.caption(f"Data commit: {csv_sha[:8]}")

    df = fetch_csv(raw_at(csv_sha, CSV_PATH))

except Exception:
    # Fallback: branch raw + cache-buster token
    token = st.session_state.refresh_token
    published = fetch_text(raw_at(BRANCH, TS_PATH) + f"?v={token}")
    st.caption(f"Published: {published}")
    df = fetch_csv(raw_at(BRANCH, CSV_PATH) + f"?v={token}")


# ----- Cleanup -----
# Drop "Unnamed: x" columns from Excel CSV export
df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed", case=False, na=False)]

# ----- Percent formatting: Win% and Edge -----
# These must match your current column names exactly (case-insensitive handled below)
percent_cols = []
for c in df.columns:
    key = str(c).strip().lower().replace(" ", "")
    if key == "win%" or key == "winpct" or "win%" in key:
        percent_cols.append(c)
    if key == "edge" or key == "edge%" or "edge%" in key:
        percent_cols.append(c)

# Make unique, preserve order
seen = set()
percent_cols = [c for c in percent_cols if not (c in seen or seen.add(c))]

df = normalize_and_format_percents(df, percent_cols)

# Display with percent formatting
if percent_cols:
    styler = df.style.format({c: "{:.2%}" for c in percent_cols})
    st.dataframe(styler, use_container_width=True, hide_index=True)
else:
    st.dataframe(df, use_container_width=True, hide_index=True)
