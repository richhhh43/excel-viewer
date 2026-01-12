#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import time
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

# =========================
# CONFIG
# =========================
EXCEL_PATH = r"C:\Users\rich_\ncaa 1.xlsm"
SHEET_NAME = "Edges"          # <-- IMPORTANT: capital E
OUT_CSV = Path("data/latest.csv")

# Hard caps so it never runs forever
MAX_ROWS = 300               # adjust if needed
LAST_COL = 11                 # A..K

def s(x):
    return "" if x is None else str(x).strip()

def find_header_row(ws, max_scan=25):
    # looks for "Event #" in column A
    for r in range(1, max_scan + 1):
        v = ws[f"A{r}"].value
        if isinstance(v, str):
            t = v.strip().lower().replace(" ", "")
            if t in ("event#", "event"):
                return r
    return 1

def make_unique(headers):
    seen = {}
    out = []
    for h in headers:
        base = h if h else "Col"
        if base not in seen:
            seen[base] = 0
            out.append(base)
        else:
            seen[base] += 1
            out.append(f"{base}.{seen[base]}")
    return out

def main():
    t0 = time.time()
    print("[publish] starting...")
    print(f"[publish] excel: {EXCEL_PATH}")
    print(f"[publish] sheet: {SHEET_NAME}")

    # FAST MODE
    print("[publish] loading workbook (read_only)...")
    wb = load_workbook(EXCEL_PATH, data_only=False, keep_vba=True, read_only=True)

    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f'Sheet "{SHEET_NAME}" not found. Sheets: {wb.sheetnames}')

    ws = wb[SHEET_NAME]
    header_row = find_header_row(ws)
    print(f"[publish] header_row={header_row}")

    # Read headers A..K
    headers = []
    for c in range(1, LAST_COL + 1):
        headers.append(s(ws.cell(row=header_row, column=c).value) or f"Col{c}")

    cols = make_unique(headers)
    print(f"[publish] columns: {cols}")

    # Read rows until blank in A or MAX_ROWS
    rows = []
    start = header_row + 1
    for r in range(start, start + MAX_ROWS):
        a = ws.cell(row=r, column=1).value
        if a in (None, ""):
            print(f"[publish] stop at row {r} (blank Event #)")
            break

        row = {}
        for c in range(1, LAST_COL + 1):
            row[cols[c - 1]] = ws.cell(row=r, column=c).value
        rows.append(row)

        if (r - start + 1) % 250 == 0:
            print(f"[publish] read {r - start + 1} rows...")

    df = pd.DataFrame(rows)
    print(f"[publish] total rows read: {len(df)}")

    # Column K (wager) is the 11th column we read
    wager_col = cols[10]   # K
    print(f"[publish] wager column (K): {wager_col}")

    # Pick column: use Market.1 if present, else last column that starts with "Market"
    pick_col = None
    market_cols = [c for c in df.columns if c.lower().startswith("market")]
    if "Market.1" in df.columns:
        pick_col = "Market.1"
    elif market_cols:
        pick_col = market_cols[-1]
    else:
        pick_col = ""

    print(f"[publish] pick column: {pick_col or '(none)'}")

    # Build QR payload text
    def build_qr(r):
        return (
            f"ALC|EVT:{s(r.get('Event #') or r.get('Event#') or r.get('Event'))}"
            f"|DT:{s(r.get('Date and Time'))}"
            f"|V:{s(r.get('Visitor'))}"
            f"|H:{s(r.get('Home'))}"
            f"|M:{s(r.get('Market'))}"
            f"|OD:{s(r.get('Odds'))}"
            f"|P:{s(r.get('Prop'))}"
            f"|WIN:{s(r.get('Win%'))}"
            f"|EDGE:{s(r.get('Edge'))}"
            f"|PICK:{s(r.get(pick_col)) if pick_col else ''}"
            f"|WAGER:${s(r.get(wager_col))}"
        )

    df["QR Code"] = df.apply(lambda row: build_qr(row.to_dict()), axis=1)

    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUT_CSV, index=False)

    print(f"[publish] saved: {OUT_CSV.as_posix()}")
    print(f"[publish] done in {time.time() - t0:.2f}s")

if __name__ == "__main__":
    main()

