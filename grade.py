import pandas as pd

def evaluate_grade(df, grade_max):
    df = df.copy()
    df["grade_standard"] = grade_max
    df["grade_ok"] = df["grade_pct"].abs() <= grade_max
    return df

def summarize_grade(df):
    rows = []
    for road, g in df.groupby("road"):
        failed = g[~g["grade_ok"]]
        worst = failed.loc[failed["grade_pct"].abs().idxmax()] if len(failed) else None
        rows.append({
            "road": road,
            "n_stations": len(g),
            "n_over_grade": len(failed),
            "worst_station": worst["station_label"] if worst is not None else "-",
            "worst_grade_pct": round(abs(worst["grade_pct"]), 2) if worst is not None else 0.0,
            "grade_standard": g["grade_standard"].iloc[0],
        })
    return pd.DataFrame(rows)
