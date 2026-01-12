import time
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("Excel Viewer")

CSV_PATH = Path("data/latest.csv")

# Cache the CSV load (and allow manual refresh)
@st.cache_data(ttl=30)
def load_csv(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    return pd.read_csv(path)

if "refresh_token" not in st.session_state:
    st.session_state.refresh_token = str(int(time.time()))

if st.button("Refresh now"):
    st.cache_data.clear()
    st.session_state.refresh_token = str(int(time.time()))
    st.rerun()

df = load_csv(CSV_PATH)

# Show published time (file modified time)
if CSV_PATH.exists():
    st.caption(f"Published: {time.strftime('%Y-%m-%d %H:%M:%S %Z', time.localtime(CSV_PATH.stat().st_mtime))}")
else:
    st.warning(f"CSV not found: {CSV_PATH.as_posix()}")

if df.empty:
    st.stop()

# -----------------------------
# QR IMAGE COLUMN
# -----------------------------
# Your screenshot shows a text column named exactly "QR Code"
# containing payload text like: ALC|EVT:484|DT:...
TEXT_COL = "QR Code"

# If your publisher named it differently, we’ll also support common fallbacks.
fallback_cols = ["qr_text", "qr_payload", "QR_TEXT"]

qr_source_col = None
if TEXT_COL in df.columns:
    qr_source_col = TEXT_COL
else:
    for c in fallback_cols:
        if c in df.columns:
            qr_source_col = c
            break

if qr_source_col:
    # Convert payload text into an image URL that returns a QR PNG
    # (QuickChart creates the QR image on the fly)
    df["QR"] = df[qr_source_col].fillna("").astype(str).apply(
        lambda s: "https://quickchart.io/qr?size=240&text=" + quote(s)
    )

    # Hide the long payload text column so it doesn't stretch the table
    df = df.drop(columns=[qr_source_col])

# Ensure QR is the LAST column
if "QR" in df.columns:
    cols = [c for c in df.columns if c != "QR"] + ["QR"]
    df = df[cols]

# -----------------------------
# DISPLAY
# -----------------------------
column_config = {}
if "QR" in df.columns:
    column_config["QR"] = st.column_config.ImageColumn("QR", width="small")

st.dataframe(
    df,
    column_config=column_config,
    hide_index=True,
    use_container_width=True,
)

