#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import time
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

# =========================
# CONFIG (EDIT THESE 2)
# =========================
EXCEL_PATH = r"C:\Users\rich_\ncaa 1.xlsm"
SHEET_NAME = "edges"

OUT_CSV = Path("data/latest.csv")

# If Column K is your wager amount, set this to the header name you use in Excel.
# If you don't have a header on K, this script will name it "Wager".
WAGER_HEADER_NAME = "Wager"


def find_header_row(ws, max_scan_rows=15):
    """Find the row that contains 'Event #' (your header row)."""
    for r in range(1, max_scan_rows + 1):
        v = ws[f"A{r}"].value
        if isinstance(v, str) and v.strip().lower() in ("event #", "event#", "event"):
            return r
    return 1  # fallback


def to_str(x):
    if x is None:
        return ""
    return str(x).strip()


def main():
    t0 = time.time()

    wb = load_workbook(EXCEL_PATH, data_only=True, keep_vba=True)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f'Sheet "{SHEET_NAME}" not found. Sheets: {wb.sheetnames}')
    ws = wb[SHEET_NAME]

    header_row = find_header_row(ws)

    # Read headers A:K
    headers = []
    for col in range(1, 12):  # 1..11 => A..K
        h = ws.cell(row=header_row, column=col).value
        headers.append(to_str(h))

    # If K header is blank, force it to "Wager"
    if not headers[10]:
        headers[10] = WAGER_HEADER_NAME

    # Normalize/rename duplicates safely
    # (Excel sometimes has 2 columns both named "Market")
    seen = {}
    safe_headers = []
    for h in headers:
        base = h if h else "Col"
        if base not in seen:
            seen[base] = 0
            safe_headers.append(base)
        else:
            seen[base] += 1
            safe_headers.append(f"{base}.{seen[base]}")

    # Build rows until Event # blank
    rows = []
    r = header_row + 1
    while True:
        event = ws[f"A{r}"].value
        if event in (None, ""):
            break

        row = {}
        for col in range(1, 12):  # A..K
            row[safe_headers[col - 1]] = ws.cell(row=r, column=col).value
        rows.append(row)
        r += 1

    df = pd.DataFrame(rows)

    # ---------
    # IMPORTANT: build QR Code text using Column K (Wager)
    # ---------
    # Expected column names after safe header handling:
    # A: "Event #"
    # B: "Date and Time"
    # C: "Visitor"
    # D: "Home"
    # E: "Market"
    # F: "Odds"
    # G: "Prop"
    # H: "Win%"
    # I: "Edge"
    # J: (often "Market.1" or similar if you had 2 "Market" headers)
    # K: "Wager" (or whatever is in your header row)
    #
    # We’ll detect the "pick" column automatically as the last "Market*" column if it exists.
    wager_col = headers[10] if headers[10] else WAGER_HEADER_NAME
    if wager_col not in df.columns:
        # If it got renamed due to duplication, find by position (K)
        wager_col = safe_headers[10]

    # Find a "pick" column (the second Market column) if present
    pick_col = None
    market_like = [c for c in df.columns if c.lower().startswith("market")]
    if len(market_like) >= 2:
        pick_col = market_like[-1]  # usually "Market.1"
    elif len(market_like) == 1:
        pick_col = market_like[0]

    def build_qr(row):
        evt = to_str(row.get("Event #", ""))
        dt  = to_str(row.get("Date and Time", ""))
        vis = to_str(row.get("Visitor", ""))
        home= to_str(row.get("Home", ""))
        mkt = to_str(row.get("Market", ""))
        odds= to_str(row.get("Odds", ""))
        prop= to_str(row.get("Prop", ""))
        win = to_str(row.get("Win%", ""))
        edge= to_str(row.get("Edge", ""))
        pick= to_str(row.get(pick_col, "")) if pick_col else ""
        wager = to_str(row.get(wager_col, ""))

        # This is YOUR tracking payload (not a redeem/validation code)
        return (
            f"ALC|EVT:{evt}"
            f"|DT:{dt}"
            f"|V:{vis}"
            f"|H:{home}"
            f"|M:{mkt}"
            f"|OD:{odds}"
            f"|P:{prop}"
            f"|WIN:{win}"
            f"|EDGE:{edge}"
            f"|PICK:{pick}"
            f"|WAGER:${wager}"
        )

    df["QR Code"] = df.apply(build_qr, axis=1)

    # Make sure output folder exists
    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)

    # Save CSV
    df.to_csv(OUT_CSV, index=False)

    print(f"Saved: {OUT_CSV.as_posix()}  rows={len(df)}  in {time.time()-t0:.2f}s")
    print(f"Wager column used: {wager_col}")
    print(f"Pick column used: {pick_col}")


if __name__ == "__main__":
    main()
