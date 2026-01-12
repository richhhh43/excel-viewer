import time
from io import StringIO

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("Excel Viewer")

# ====== EDIT THESE (your GitHub repo info) ======
OWNER = "richhhh43"
REPO = "excel-viewer"
BRANCH = "main"
# ==============================================

CSV_PATH = "data/latest.csv"
TS_PATH = "data/updated_at.txt"
API_COMMITS = f"https://api.github.com/repos/{OWNER}/{REPO}/commits"


def raw_at(ref: str, path: str) -> str:
    return f"https://raw.githubusercontent.com/{OWNER}/{REPO}/{ref}/{path}"


@st.cache_data(ttl=30)
def latest_sha_for_path(path: str) -> str:
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


def col_exists(df: pd.DataFrame, name: str):
    if name in df.columns:
        return name
    low = {str(c).strip().lower(): c for c in df.columns}
    return low.get(name.strip().lower())


def normalize_percent_series(s: pd.Series) -> pd.Series:
    # Accept both 0.2877 and 28.77 and normalize to 0-1
    s = pd.to_numeric(s, errors="coerce")
    if s.dropna().size and s.dropna().median() > 1:
        return s / 100.0
    return s


def normalize_number_series(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


# Refresh button
if "refresh_token" not in st.session_state:
    st.session_state.refresh_token = str(int(time.time()))

if st.button("Refresh now"):
    st.cache_data.clear()
    st.session_state.refresh_token = str(int(time.time()))
    st.rerun()


# Load using immutable SHA urls (best anti-cache)
try:
    ts_sha = latest_sha_for_path(TS_PATH)
    csv_sha = latest_sha_for_path(CSV_PATH)

    published = fetch_text(raw_at(ts_sha, TS_PATH))
    st.caption(f"Published: {published}")
    st.caption(f"Data commit: {csv_sha[:8]}")

    df = fetch_csv(raw_at(csv_sha, CSV_PATH))

except Exception:
    token = st.session_state.refresh_token
    published = fetch_text(raw_at(BRANCH, TS_PATH) + f"?v={token}")
    st.caption(f"Published: {published}")
    df = fetch_csv(raw_at(BRANCH, CSV_PATH) + f"?v={token}")


# Drop Excel Unnamed columns
df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed", case=False, na=False)]

# Columns (case-insensitive)
win_col = col_exists(df, "Win%")
edge_col = col_exists(df, "Edge")
odds_col = col_exists(df, "Odds")
prop_col = col_exists(df, "Prop")

# Normalize values
if win_col:
    df[win_col] = normalize_percent_series(df[win_col])
if edge_col:
    df[edge_col] = normalize_percent_series(df[edge_col])

if odds_col:
    df[odds_col] = normalize_number_series(df[odds_col])
if prop_col:
    df[prop_col] = normalize_number_series(df[prop_col])

# Formatting:
# Win% / Edge -> percent
# Odds -> 2 decimals (3.95)
# Prop -> 1 decimal (232.5)
fmt = {}
if win_col:
    fmt[win_col] = "{:.2%}"
if edge_col:
    fmt[edge_col] = "{:.2%}"
if odds_col:
    fmt[odds_col] = "{:.2f}"
if prop_col:
    fmt[prop_col] = "{:.1f}"

styler = df.style.format(fmt)
st.dataframe(styler, width="stretch", hide_index=True)
