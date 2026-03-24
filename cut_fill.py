"""
cut_fill.py — Estimate cut & fill volume per 50m station
based on grade change before vs after repair (Tabel 19-20 Figo 2024).

Method:
  Δh = (grade_after - grade_before) / 100 * segment_length
  If Δh > 0 → fill (road surface needs raising)
  If Δh < 0 → cut (road surface needs lowering)
  Volume = |Δh| * road_width * segment_length

Road width = 25 m (from journal data, Tabel 6-7)
Segment length = 50 m
"""
import pandas as pd
import numpy as np

ROAD_WIDTH_M = 25.0
SEGMENT_LEN_M = 50.0

GRADE_BEFORE = {
    "MonteBawah": {
        "0-50": 7.02, "50-100": 8.61, "100-150": 4.00, "150-200": 6.40,
        "200-250": 6.77, "250-300": 7.15, "300-350": 6.21, "350-400": 7.13,
        "400-450": 7.35, "450-500": 8.00, "500-550": 8.56, "550-600": 7.76,
        "600-650": 5.59, "650-700": 6.04, "700-750": 7.94, "750-800": 9.57,
        "800-850": 9.94,
    },
    "Spanyol": {
        "0-50": -0.24, "50-100": -1.02, "100-150": 1.00, "150-200": 4.02,
        "200-250": 6.76, "250-300": 6.38, "300-350": 6.30, "350-400": 2.85,
        "400-450": 1.94, "450-500": 2.83, "500-550": 4.21, "550-600": 4.21,
        "600-650": 7.53, "650-700": 10.18,
    },
}

GRADE_AFTER = {
    "MonteBawah": {
        "0-50": 4.22, "50-100": 7.12, "100-150": 5.12, "150-200": 7.46,
        "200-250": 7.65, "250-300": 7.15, "300-350": 6.21, "350-400": 7.13,
        "400-450": 7.95, "450-500": 7.96, "500-550": 7.95, "550-600": 7.81,
        "600-650": 7.13, "650-700": 7.77, "700-750": 7.95, "750-800": 6.71,
        "800-850": 7.84,
    },
    "Spanyol": {
        "0-50": -0.24, "50-100": -1.02, "100-150": 1.00, "150-200": 4.02,
        "200-250": 6.75, "250-300": 5.75, "300-350": 6.19, "350-400": 2.71,
        "400-450": 1.89, "450-500": 3.49, "500-550": 4.49, "550-600": 5.56,
        "600-650": 7.68, "650-700": 7.92,
    },
}

def compute_cut_fill():
    rows = []
    for road in ["MonteBawah", "Spanyol"]:
        stations = sorted(GRADE_BEFORE[road].keys(),
                          key=lambda s: int(s.split("-")[0]))
        for sta in stations:
            g_before = GRADE_BEFORE[road][sta]
            g_after  = GRADE_AFTER[road].get(sta, g_before)
            delta_grade = g_after - g_before  # %

            # height change over 50m segment
            dh = (delta_grade / 100.0) * SEGMENT_LEN_M  # metres
            vol = abs(dh) * ROAD_WIDTH_M * SEGMENT_LEN_M  # m³

            work_type = "fill" if dh > 0 else ("cut" if dh < 0 else "none")

            rows.append({
                "road": road,
                "station_label": sta,
                "grade_before_pct": round(g_before, 2),
                "grade_after_pct": round(g_after, 2),
                "delta_grade_pct": round(delta_grade, 2),
                "dh_m": round(dh, 3),
                "volume_m3": round(vol, 1),
                "work_type": work_type,
            })

    df = pd.DataFrame(rows)

    summary = []
    for road, g in df.groupby("road"):
        cut = g.loc[g["work_type"] == "cut", "volume_m3"].sum()
        fill = g.loc[g["work_type"] == "fill", "volume_m3"].sum()
        summary.append({
            "road": road,
            "total_cut_m3": round(cut, 1),
            "total_fill_m3": round(fill, 1),
            "net_m3": round(fill - cut, 1),
            "n_cut_stations": (g["work_type"] == "cut").sum(),
            "n_fill_stations": (g["work_type"] == "fill").sum(),
        })

    return df, pd.DataFrame(summary)
