import pandas as pd

RDS_TO_RR = [
    (0,   10,  1.0),
    (11,  20,  1.5),
    (21,  30,  2.0),
    (31,  40,  2.5),
    (41,  50,  3.0),
    (51,  60,  3.1),
    (61,  70,  3.2),
    (71,  80,  3.5),
    (81,  100, 4.0),
    (101, 999, 5.0),
]

def rds_to_rr(total_rds):
    for lo, hi, rr in RDS_TO_RR:
        if lo <= total_rds <= hi:
            return rr
    return 5.0

def score_rds(df_damage):
    df = df_damage.copy()
    df["defect_score"] = df["degree"] * df["extent"]
    total = df.groupby("road")["defect_score"].transform("sum")
    df["contribution_pct"] = round(df["defect_score"] / total * 100, 1)

    summary_rows = []
    for road, g in df.groupby("road"):
        total_rds = g["defect_score"].sum()
        rr_pct = rds_to_rr(total_rds)
        dom = g.loc[g["defect_score"].idxmax(), "defect"]
        summary_rows.append({
            "road": road,
            "total_rds": total_rds,
            "rr_pct": rr_pct,
            "rr_target": 2.0,
            "rr_ok": rr_pct <= 2.0,
            "dominant_defect": dom,
        })

    return df, pd.DataFrame(summary_rows)
