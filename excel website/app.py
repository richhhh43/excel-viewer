import streamlit as st
import pandas as pd

st.set_page_config(layout="wide")
st.title("📊 Online Excel Viewer")

EXCEL_FILE = "my_data.xlsm"

@st.cache_data
def load_excel(sheet_name):
    return pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, engine="openpyxl")

# Load workbook
try:
    excel_file = pd.ExcelFile(EXCEL_FILE, engine="openpyxl")
    sheet_names = excel_file.sheet_names
except Exception as e:
    st.error(f"Error loading Excel file: {e}")
    st.stop()

# Sidebar UI
sheet = st.sidebar.selectbox("Select sheet:", sheet_names)

# Load and display sheet
try:
    df = load_excel(sheet)
    st.dataframe(df, use_container_width=True)
except Exception as e:
    st.error(f"Error loading sheet: {e}")

    

