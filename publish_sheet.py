def fmt_pct(x):
    if x is None or x == "":
        return ""
    txt = str(x).strip().replace("%", "")
    try:
        v = float(txt)
    except Exception:
        return str(x)
    if 0 <= v <= 1:
        v *= 100.0
    return f"{v:.2f}%"

if "Win%" in df.columns:
    df["Win%"] = df["Win%"].apply(fmt_pct)
if "Edge" in df.columns:
    df["Edge"] = df["Edge"].apply(fmt_pct)
