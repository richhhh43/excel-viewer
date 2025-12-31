import subprocess
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import pandas as pd

WORKBOOK_PATH = r"C:\Users\rich_\ncaa 1.xlsm"
SHEET_NAME = "Edges"

REPO_PATH = Path(__file__).resolve().parent
DATA_DIR = REPO_PATH / "data"
OUT_CSV = DATA_DIR / "latest.csv"
OUT_TS = DATA_DIR / "updated_at.txt"

def run(cmd, cwd=None):
    subprocess.run(cmd, cwd=cwd, check=True)

def git_has_changes(repo: Path) -> bool:
    r = subprocess.run(["git", "status", "--porcelain"], cwd=repo, capture_output=True, text=True)
    return r.stdout.strip() != ""

def main():
    DATA_DIR.mkdir(exist_ok=True)

    # Export sheet to CSV
    df = pd.read_excel(WORKBOOK_PATH, sheet_name=SHEET_NAME, engine="openpyxl")
    df.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")

    # Write publish timestamp (Toronto time)
    ts = datetime.now(ZoneInfo("America/Toronto")).strftime("%Y-%m-%d %H:%M:%S %Z")
    OUT_TS.write_text(ts, encoding="utf-8")

    print(f"Wrote: {OUT_CSV}")
    print(f"Published: {ts}")

    # Auto commit + push
    run(["git", "add", "data/latest.csv", "data/updated_at.txt"], cwd=REPO_PATH)

    if not git_has_changes(REPO_PATH):
        print("No changes to push.")
        return

    run(["git", "commit", "-m", f"Update data {ts}"], cwd=REPO_PATH)
    run(["git", "push"], cwd=REPO_PATH)
    print("Pushed to GitHub.")

if __name__ == "__main__":
    main()

