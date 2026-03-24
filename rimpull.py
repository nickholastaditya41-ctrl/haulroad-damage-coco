"""
rimpull.py — Hitung TR, gear, dan speed per station
====================================================
Formula:
  TR (lb) = (RR% + grade%) × GVW_ton × 20
  → pilih gear tertinggi (kecepatan maks) yang rimpull_capacity >= TR
  → speed = gear_speed[gear]

RR aktual per station: diinfer balik dari TR jurnal (Tabel 14-15).
RR setelah perbaikan: 2.0% (target perusahaan).
Grade setelah perbaikan: Tabel 19-20 jurnal Figo 2024.

Downhill (grade negatif): TR = (RR - |grade|) × GVW_ton × 20, min 50% of RR-only.
"""

import pandas as pd
import numpy as np

GVW_LB   = 374785.85
GVW_TON  = GVW_LB / 2000        # 187.39 ton
RR_AFTER = 2.0                   # % setelah perbaikan

TR_ACTUAL = {
    "MonteBawah": {
        "0-50":   34918, "50-100":  40324, "100-150": 24650, "150-200": 32810,
        "200-250": 34068, "250-300": 35360, "300-350": 46444, "350-400": 49572,
        "400-450": 50320, "450-500": 52530, "500-550": 54434, "550-600": 51714,
        "600-650": 44336, "650-700": 45886, "700-750": 52326, "750-800": 57868,
        "800-850": 44846,
    },
    "Spanyol": {
        "0-50":   10234, "50-100":   7582, "100-150": 14450, "150-200": 24718,
        "200-250": 49334, "250-300": 48042, "300-350": 38590, "350-400": 30940,
        "400-450": 23766, "450-500": 26792, "500-550": 31484, "550-600": 31484,
        "600-650": 42772, "650-700": 45662,
    },
}

GRADE_ACTUAL = {
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

GRADE_IMPROVED = {
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


def _tr(rr_pct, grade_pct):
    if grade_pct >= 0:
        return (rr_pct + grade_pct) * GVW_TON * 20
    else:
        return max((rr_pct - abs(grade_pct)) * GVW_TON * 20,
                   rr_pct * GVW_TON * 20 * 0.5)


def _rr_actual(road, sta):
    TR_j  = TR_ACTUAL[road][sta]
    grade = GRADE_ACTUAL[road][sta]
    gr_lb = grade * GVW_TON * 20
    rr_lb = TR_j - gr_lb
    return rr_lb / (GVW_TON * 20)


def _gear_speed(tr_lb, rimpull_cap, gear_speed):
    best_gear = min(rimpull_cap.keys())
    for g in sorted(rimpull_cap.keys(), reverse=True):
        if rimpull_cap[g] >= tr_lb:
            best_gear = g
            break
    return best_gear, gear_speed[best_gear]


def compute_rimpull(df_cbr, df_grade, cfg):
    rimpull_cap = {int(k): v for k, v in cfg["rimpull_gear_lb"].items()}
    gear_spd    = {int(k): v for k, v in cfg["gear_speed_kph"].items()}

    rows = []
    for road in TR_ACTUAL:
        for sta in TR_ACTUAL[road]:
            grade_before = GRADE_ACTUAL[road][sta]
            grade_after  = GRADE_IMPROVED[road][sta]

            rr_before    = _rr_actual(road, sta)
            tr_before    = _tr(rr_before, grade_before)
            gear_b, spd_b = _gear_speed(tr_before, rimpull_cap, gear_spd)

            tr_after     = _tr(RR_AFTER, grade_after)
            gear_a, spd_a = _gear_speed(tr_after, rimpull_cap, gear_spd)

            rows.append({
                "road":             road,
                "station_label":    sta,
                "grade_before_pct": round(grade_before, 2),
                "grade_after_pct":  round(grade_after,  2),
                "RR_before_pct":    round(rr_before, 2),
                "RR_after_pct":     RR_AFTER,
                "TR_before_lb":     round(tr_before, 0),
                "TR_after_lb":      round(tr_after,  0),
                "gear_before":      gear_b,
                "gear_after":       gear_a,
                "speed_before_kph": spd_b,
                "speed_after_kph":  spd_a,
                "speed_delta_kph":  round(spd_a - spd_b, 1),
            })

    df_out = pd.DataFrame(rows)

    summary = df_out.groupby("road").agg(
        speed_actual_avg=("speed_before_kph", "mean"),
        speed_improved_avg=("speed_after_kph", "mean"),
    ).round(2).reset_index()
    summary["speed_delta_avg"] = (
        summary["speed_improved_avg"] - summary["speed_actual_avg"]
    ).round(2)

    return df_out, summary
