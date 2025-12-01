import streamlit as st
import pandas as pd
from io import BytesIO
import requests

st.set_page_config(page_title="Excel Viewer", layout="wide")

# -------------------------------------------------
# CONFIG – CHOOSE ONE MODE
# -------------------------------------------------
# MODE 1: BUNDLED FILE (put the Excel file in the same folder as app.py)
EXCEL_PATH = "my_data.xlsx"

# MODE 2: ONLINE FILE (direct download URL from OneDrive / Google Drive etc.)
# If you have a direct download link, paste it here and leave EXCEL_PATH as backup.
EXCEL_URL = ""  # e.g. "https://.../my_data.xlsx"


@st.cache_data
def load_excel_from_path(path: str):
    return pd.read_excel(path, sheet_name=None)


@st.cache_data
def load_excel_from_url(url: str):
    resp = requests.get(url)
    resp.raise_for_status()
    return pd.read_excel(BytesIO(resp.content), sheet_name=None)


def load_workbook():
    # Try URL first if provided
    if EXCEL_URL.strip():
        try:
            return load_excel_from_url(EXCEL_URL.strip()), "url"
        except Exception as e:
            st.warning(f"Could not load from URL, falling back to local file. Error: {e}")

    # Fall back to local file
    return load_excel_from_path(EXCEL_PATH), "local"


def main():
    st.title("📊 Excel Viewer")

    st.write(
        "This site displays data from an Excel workbook. "
        "Use the dropdown below to switch sheets."
    )

    try:
        workbook, source = load_workbook()
    except FileNotFoundError:
        st.error(
            f"Could not find Excel file.\n\n"
            f"Expected local file: `{EXCEL_PATH}`\n\n"
            f"Upload it to the same folder as `app.py` or configure `EXCEL_URL`."
        )
        return
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return

    sheet_names = list(workbook.keys())

    if not sheet_names:
        st.warning("No sheets found in the workbook.")
        return

    col1, col2 = st.columns([2, 1])
    with col1:
        selected_sheet = st.selectbox("Select sheet", sheet_names)

    with col2:
        st.caption(f"Data source: **{source}**")

    df = workbook[selected_sheet]

    st.subheader(f"Sheet: {selected_sheet}")
    st.dataframe(df, use_container_width=True)


if __name__ == "__main__":
    main()
