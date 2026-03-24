import pandas as pd

def evaluate_cbr(df, cbr_min):
    df = df.copy()
    df["cbr_standard"] = cbr_min
    df["cbr_ok"] = df["cbr_loaded"] >= cbr_min
    return df

def summarize_cbr(df):
    rows = []
    for road, g in df.groupby("road"):
        rows.append({
            "road": road,
            "n_stations": len(g),
            "n_ok": g["cbr_ok"].sum(),
            "n_fail": (~g["cbr_ok"]).sum(),
            "pct_fail": round((~g["cbr_ok"]).mean() * 100, 1),
            "cbr_min_actual": g["cbr_loaded"].min(),
            "cbr_max_actual": g["cbr_loaded"].max(),
            "cbr_mean_actual": round(g["cbr_loaded"].mean(), 1),
            "cbr_standard": g["cbr_standard"].iloc[0],
        })
    return pd.DataFrame(rows)
