#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import time
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

EXCEL_PATH = r"C:\Users\rich_\ncaa 1.xlsm"
SHEET_NAME = "Edges"          # case-sensitive
OUT_CSV = Path("data/latest.csv")

MAX_ROWS = 8000
LAST_COL = 10                # A..J only (no wager, no QR)

def s(x):
    return "" if x is None else str(x).strip()

def is_numberish(x) -> bool:
    if x in (None, ""):
        return False
    try:
        float(str(x).strip().replace(",", ""))
        return True
    except Exception:
        return False

def find_header_row(ws, max_scan=25):
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

def fmt_pct(x):
    """Excel stores 28.04% as 0.2804. Convert to '28.04%'."""
    if x is None or x == "":
        return ""
    txt = str(x).strip().replace("%", "")
    try:
        v = float(txt)
    except Exception:
        return s(x)
    if 0 <= v <= 1:
        v *= 100.0
    return f"{v:.2f}%"

def main():
    t0 = time.time()
    print("[publish] starting...")
    print("[publish] loading workbook (read_only)...")

    wb = load_workbook(EXCEL_PATH, data_only=False, keep_vba=True, read_only=True)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f'Sheet "{SHEET_NAME}" not found. Sheets: {wb.sheetnames}')

    ws = wb[SHEET_NAME]
    header_row = find_header_row(ws)
    print(f"[publish] header_row={header_row}")

    headers = []
    for c in range(1, LAST_COL + 1):
        headers.append(s(ws.cell(row=header_row, column=c).value) or f"Col{c}")
    cols = make_unique(headers)
    print(f"[publish] columns: {cols}")

    rows = []
    start = header_row + 1

    for i, rowvals in enumerate(
        ws.iter_rows(
            min_row=start,
            max_row=start + MAX_ROWS - 1,
            min_col=1,
            max_col=LAST_COL,
            values_only=True
        ),
        start=1
    ):
        evt = rowvals[0]
        if evt in (None, ""):
            break
        if not is_numberish(evt):
            continue

        rows.append({cols[c]: rowvals[c] for c in range(LAST_COL)})

        if i % 1000 == 0:
            print(f"[publish] scanned {i} rows... kept {len(rows)}")

    df = pd.DataFrame(rows)

    # Force Win% and Edge to display like your screenshot
    if "Win%" in df.columns:
        df["Win%"] = df["Win%"].apply(fmt_pct)
    if "Edge" in df.columns:
        df["Edge"] = df["Edge"].apply(fmt_pct)

    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUT_CSV, index=False)

    print(f"[publish] saved: {OUT_CSV.as_posix()} rows={len(df)}")
    print(f"[publish] done in {time.time()-t0:.2f}s")

if __name__ == "__main__":
    main()
