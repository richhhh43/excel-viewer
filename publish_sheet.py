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
SHEET_NAME = "Edges"          # must match exactly (case-sensitive)
OUT_CSV = Path("data/latest.csv")

MAX_ROWS = 5000               # raise/lower as you like
LAST_COL = 11                 # A..K

def s(x):
    return "" if x is None else str(x).strip()

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
    """
    Excel stores 28.04% as 0.2804 with % formatting.
    Convert fractions to a percent string.
    """
    if x is None or x == "":
        return ""
    try:
        v = float(x)
    except Exception:
        return s(x)
    if 0 <= v <= 1:
        return f"{v*100:.2f}%"
    # if it's already like 28.04, treat as percent value
    return f"{v:.2f}%"

def main():
    t0 = time.time()
    print("[publish] starting...")
    print(f"[publish] excel: {EXCEL_PATH}")
    print(f"[publish] sheet: {SHEET_NAME}")

    # FAST: read_only + iter_rows
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

    # Bulk read rows A..K
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
        if rowvals[0] in (None, ""):
            print(f"[publish] stop at row {start + i - 1} (blank Event #)")
            break

        rows.append({cols[c]: rowvals[c] for c in range(LAST_COL)})

        if i % 500 == 0:
            print(f"[publish] read {i} rows...")

    df = pd.DataFrame(rows)
    print(f"[publish] total rows read: {len(df)}")

    # Column mapping
    wager_col = cols[10]  # K
    market_cols = [c for c in df.columns if c.lower().startswith("market")]
    pick_col = "Market.1" if "Market.1" in df.columns else (market_cols[-1] if market_cols else "")

    print(f"[publish] wager column (K): {wager_col}")
    print(f"[publish] pick column: {pick_col or '(none)'}")

    # Build QR payload text (percent formatted)
    def build_qr(r):
        return (
            f"ALC|EVT:{s(r.get('Event #') or r.get('Event#') or r.get('Event'))}"
            f"|DT:{s(r.get('Date and Time'))}"
            f"|V:{s(r.get('Visitor'))}"
            f"|H:{s(r.get('Home'))}"
            f"|M:{s(r.get('Market'))}"
            f"|OD:{s(r.get('Odds'))}"
            f"|P:{s(r.get('Prop'))}"
            f"|WIN:{fmt_pct(r.get('Win%'))}"
            f"|EDGE:{fmt_pct(r.get('Edge'))}"
            f"|PICK:{s(r.get(pick_col)) if pick_col else ''}"
            f"|WAGER:{s(r.get(wager_col))}"
        )

    df["QR Code"] = df.apply(lambda row: build_qr(row.to_dict()), axis=1)

    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUT_CSV, index=False)

    print(f"[publish] saved: {OUT_CSV.as_posix()}")
    print(f"[publish] done in {time.time() - t0:.2f}s")

if __name__ == "__main__":
    main()

