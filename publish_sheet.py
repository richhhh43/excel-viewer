import subprocess
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd

# ====== EDIT THESE TWO (your local workbook) ======
WORKBOOK_PATH = r"C:\Users\rich_\ncaa 1.xlsm"
SHEET_NAME = "Edges"  # must match exactly (case-sensitive)
# ================================================

REPO_PATH = Path(__file__).resolve().parent
DATA_DIR = REPO_PATH / "data"
OUT_CSV = DATA_DIR / "latest.csv"
OUT_TS = DATA_DIR / "updated_at.txt"


def run(cmd, cwd=None):
    subprocess.run(cmd, cwd=cwd, check=True)


def git_available() -> bool:
    try:
        subprocess.run(["git", "--version"], capture_output=True, text=True, check=True)
        return True
    except Exception:
        return False


def main():
    DATA_DIR.mkdir(exist_ok=True)

    # Export sheet to CSV
    df = pd.read_excel(WORKBOOK_PATH, sheet_name=SHEET_NAME, engine="openpyxl")
    df.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")

    # Timestamp (Toronto time)
    ts = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d %H:%M:%S %Z")
    OUT_TS.write_text(ts, encoding="utf-8")

    print(f"Wrote: {OUT_CSV}")
    print(f"Published: {ts}")

    # Commit + push changes (if git works)
    if not git_available():
        print("NOTE: git not found. CSV/timestamp updated locally but NOT pushed.")
        return

    run(["git", "add", "data/latest.csv", "data/updated_at.txt"], cwd=REPO_PATH)

    # If nothing changed, don't commit
    status = subprocess.run(["git", "status", "--porcelain"], cwd=REPO_PATH, capture_output=True, text=True)
    if status.stdout.strip() == "":
        print("No changes to commit.")
        return

    msg = f"Update data {ts}"
    run(["git", "commit", "-m", msg], cwd=REPO_PATH)
    run(["git", "push", "origin", "main"], cwd=REPO_PATH)
    print("Pushed to GitHub.")


if __name__ == "__main__":
    main()
