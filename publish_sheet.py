from pathlib import Path
import pandas as pd

# ===== EDIT THESE TWO =====
WORKBOOK_PATH = r"C:\Users\rich_\ncaa 1.xlsm"
SHEET_NAME = "Edges"

REPO_PATH = Path(__file__).resolve().parent
OUT_CSV = REPO_PATH / "data" / "latest.csv"

def main():
    df = pd.read_excel(WORKBOOK_PATH, sheet_name=SHEET_NAME, engine="openpyxl")
    OUT_CSV.parent.mkdir(exist_ok=True)
    df.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")
    print(f"Wrote: {OUT_CSV}")

if __name__ == "__main__":
    main()

