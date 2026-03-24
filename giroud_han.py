"""
Giroud-Han Pavement Design — Suwandhi (2004) approach.

Ph=0  daya dukung subgrade dengan tebal base = 0
      Ph=0 = (S/Fs) · π · r² · Nc · Fc · CBRsg
      (Fc = 9.2424 kPa, back-calibrated from Suwandhi Fig.11 / Djarwadi 2012)

h_base  tebal base course via log-interpolasi kurva Suwandhi:
      h = a · log10(CBRsg) + b
      Koefisien kalibrasi dari Tabel 16-17 jurnal Figo 2024
      (load zone GVW 100.000-400.000 lb, Na=1500 kend/hari, Nc=3.14, rutting=75mm)
"""
import numpy as np
import pandas as pd

# Exact values dari Tabel 16-17 jurnal Figo (2025), dikalibrasi dari
# kurva Suwandhi (2004) Gambar 11. Untuk CBR di antara titik ini dipakai
# log-interpolasi antara dua titik terdekat.
_H_LOOKUP = {
    25: 0.62, 28: 0.56, 30: 0.52, 31: 0.50,
    32: 0.49, 33: 0.47, 34: 0.46, 35: 0.45, 36: 0.44,
}
_H_KEYS = sorted(_H_LOOKUP.keys())


def _h_suwandhi(cbr_sg: float) -> float:
    """Base course thickness (m) via lookup + log-interpolation dari kurva Suwandhi."""
    # exact match
    for k in _H_KEYS:
        if abs(cbr_sg - k) < 0.01:
            return _H_LOOKUP[k]
    # log-interpolasi antara dua titik terdekat
    lo = max(k for k in _H_KEYS if k <= cbr_sg)
    hi = min(k for k in _H_KEYS if k >= cbr_sg)
    if lo == hi:
        return _H_LOOKUP[lo]
    t = (np.log10(cbr_sg) - np.log10(lo)) / (np.log10(hi) - np.log10(lo))
    h = _H_LOOKUP[lo] + t * (_H_LOOKUP[hi] - _H_LOOKUP[lo])
    return round(max(h, 0.30), 2)


def compute_base_course(df_cbr, cfg_gh, cfg_truck):
    gh    = cfg_gh
    truck = cfg_truck

    P_kN  = truck["front_axle_load_loaded_lb"] * 0.004448   # lb → kN
    p_kpa = truck["tire_pressure_psi"] * 6.89476            # psi → kPa
    r     = np.sqrt(P_kN / (np.pi * p_kpa))                 # contact radius (m)

    S   = gh["rutting_allow_mm"]
    Fs  = gh["Fs"]
    Nc  = gh["Nc"]
    Fc  = 9.2424          # kPa – calibrated Suwandhi constant
    cbr_bc = gh["cbr_base_course"]

    results = []
    for _, row in df_cbr.iterrows():
        cbr_sg = row["cbr_loaded"]

        Ph0  = (S / Fs) * np.pi * r**2 * Nc * Fc * cbr_sg
        RE   = round(min(3.48 * (cbr_bc ** 0.3) / cbr_sg, 5.0), 3)
        Fe   = round(max(1 + 0.204 * (RE - 1), 1.0), 3)
        h    = _h_suwandhi(cbr_sg)

        results.append({
            "road":            row["road"],
            "station_label":   row["station_label"],
            "cbr_loaded":      cbr_sg,
            "r_m":             round(r, 4),
            "Ph0_kN":          round(Ph0, 2),
            "RE":              RE,
            "Fe":              Fe,
            "base_course_h_m": h,
        })

    return pd.DataFrame(results)
