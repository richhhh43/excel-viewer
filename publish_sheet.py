#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import time
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

EXCEL_PATH = r"C:\Users\rich_\ncaa 1.xlsm"
SHEET_NAME = "Edges"
OUT_CSV = Path("data/latest.csv")

MAX_ROWS = 600        # <-- hard stop so it can't hang forever
MAX_SCAN_ROWS = 20     # header scan
LAST_COL = 11          # A..K

def to_str(x):
    return "" if x is None else str(x).strip()

def find_header_row(ws):
    for r in range(1, MAX_SCAN_ROWS + 1):
        v = ws[f"A{r}"].value
        if isinstance(v, str) and v.strip().lower().replace(" ", "") in ("event#", "event"):
            return r
    return 1

def main():
    print("[publish] starting...")
    t0 = time.time()

    print("[publish] loading workbook (this can take a bit on big xlsm)...")
    wb = load_workbook(EXCEL_PATH, data_only=True, keep_vba=True, read_only=True)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f'Sheet "{SHEET_NAME}" not found. Sheets: {wb.sheetnames}')
    ws = wb[SHEET_NAME]

    header_row = find_header_row(ws)
    print(f"[publish] header_row = {header_row}")

    # Read headers A..K
    headers = []
    for c in range(1, LAST_COL + 1):
        h = ws.cell(row=header_row, column=c).value
        headers.append(to_str(h) or f"Col{c}")

    # Make headers unique (Market, Market.1, etc.)
    seen = {}
    cols = []
    for h in headers:
        if h not in seen:
            seen[h] = 0
            cols.append(h)
        else:
            seen[h] += 1
            cols.append(f"{h}.{seen[h]}")

    print(f"[publish] columns: {cols}")

    # Pull rows until blank A OR MAX_ROWS
    rows = []
    start = header_row + 1
    for r in range(start, start + MAX_ROWS):
        a = ws.cell(row=r, column=1).value
        if a in (None, ""):
            print(f"[publish] stop at row {r} (blank Event#).")
            break

        row = {}
        for c in range(1, LAST_COL + 1):
            row[cols[c-1]] = ws.cell(row=r, column=c).value
        rows.append(row)

        if (r - start + 1) % 250 == 0:
            print(f"[publish] read {r - start + 1} rows...")

    df = pd.DataFrame(rows)
    print(f"[publish] df rows = {len(df)}")

    # Identify wager column (K)
    wager_col = cols[10]  # column K name (whatever header is)
    # Identify pick column: prefer Market.1 if present, else any column starting with Market
    pick_col = "Market.1" if "Market.1" in df.columns else next((c for c in df.columns if c.lower().startswith("market")), "")

    # Build QR payload text column
    def qr_row(x):
        return (
            f"ALC|EVT:{to_str(x.get('Event #')) or to_str(x.get('Event#')) or to_str(x.get('Event'))}"
            f"|DT:{to_str(x.get('Date and Time'))}"
            f"|V:{to_str(x.get('Visitor'))}"
            f"|H:{to_str(x.get('Home'))}"
            f"|M:{to_str(x.get('Market'))}"
            f"|OD:{to_str(x.get('Odds'))}"
            f"|P:{to_str(x.get('Prop'))}"
            f"|WIN:{to_str(x.get('Win%'))}"
            f"|EDGE:{to_str(x.get('Edge'))}"
            f"|PICK:{to_str(x.get(pick_col))}"
            f"|WAGER:${to_str(x.get(wager_col))}"
        )

    df["QR Code"] = df.apply(lambda r: qr_row(r.to_dict()), axis=1)

    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUT_CSV, index=False)

    print(f"[publish] saved: {OUT_CSV.as_posix()}")
    print(f"[publish] wager column (K): {wager_col}")
    print(f"[publish] pick column: {pick_col}")
    print(f"[publish] done in {time.time()-t0:.2f}s")

if __name__ == "__main__":
    main()

