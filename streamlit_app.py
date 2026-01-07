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


def find_percent_cols(cols) -> list:
    """
    Detect Win% and Edge columns (case/space insensitive).
    Matches:
      - Win%, Win %, win%
      - Edge, Edge%, edge
    """
    out = []
    for c in cols:
        key = str(c).strip().lower().replace(" ", "")
        if key in ("win%", "winpct") or "win%" in key or key == "win":
            out.append(c)
        elif key == "edge" or "edge%" in key or "vedge" in key:
            out.append(c)
    return out


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


# ----- Cleanup + formatting -----
# Drop "Unnamed: x" columns that come from Excel exports
df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed", case=False, na=False)]

pct_cols = find_percent_cols(df.columns)

# Normalize percent columns to 0-1 if needed
for c in pct_cols:
    s = pd.to_numeric(df[c], errors="coerce")
    if s.dropna().size and s.dropna().median() > 1:
        df[c] = s / 100.0
    else:
        df[c] = s

# Display: format percent columns as 12.34%
if pct_cols:
    styler = df.style.format({c: "{:.2%}" for c in pct_cols})
    st.dataframe(styler, use_container_width=True, hide_index=True)
else:
    st.dataframe(df, use_container_width=True, hide_index=True)
