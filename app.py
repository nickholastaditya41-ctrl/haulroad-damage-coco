"""
Road Damage Analyzer — Streamlit Dashboard
PT PPA Adaro Indonesia · Komatsu HD 785-7
=========================================================
Run:  streamlit run app.py
Deps: streamlit plotly pandas openpyxl xlsxwriter
"""

import io
import warnings
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import streamlit as st

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Road Haul Damage Analyzer · PT PPA Adaro",
    page_icon="🛣️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
#  THEME — DARK NAVY INDUSTRIAL
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600&family=Syne:wght@400;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Syne', sans-serif;
}
.stApp {
    background: #090F1A;
}
section[data-testid="stSidebar"] {
    background: #0C1829 !important;
    border-right: 1px solid #1A2E44;
}
section[data-testid="stSidebar"] * {
    color: #7AAABF !important;
}
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stFileUploader label,
section[data-testid="stSidebar"] h2, 
section[data-testid="stSidebar"] h3 {
    color: #4A7A9B !important;
    font-size: 10px !important;
    letter-spacing: 2px !important;
    text-transform: uppercase !important;
}

/* ── KPI cards ── */
.kpi-card {
    background: #0C1829;
    border: 1px solid #1A2E44;
    border-radius: 10px;
    padding: 16px 18px;
}
.kpi-label {
    font-size: 10px;
    color: #3A6A8A;
    letter-spacing: 2px;
    text-transform: uppercase;
    margin-bottom: 6px;
    font-family: 'JetBrains Mono', monospace;
}
.kpi-value {
    font-size: 26px;
    font-weight: 800;
    font-family: 'JetBrains Mono', monospace;
    line-height: 1;
}
.kpi-sub {
    font-size: 11px;
    font-family: 'JetBrains Mono', monospace;
    margin-top: 5px;
}
.kpi-red   { color: #E74C3C; }
.kpi-amber { color: #F39C12; }
.kpi-green { color: #2ECC71; }
.kpi-blue  { color: #5DADE2; }
.sub-red   { color: #7B2D2D; }
.sub-green { color: #1D6A40; }
.sub-amber { color: #7B5412; }
.sub-blue  { color: #1A4A6A; }

/* ── Section header ── */
.sec-header {
    font-size: 10px;
    color: #3A6A8A;
    letter-spacing: 3px;
    text-transform: uppercase;
    font-weight: 700;
    margin-bottom: 14px;
    font-family: 'JetBrains Mono', monospace;
    border-bottom: 1px solid #1A2E44;
    padding-bottom: 6px;
}

/* ── BI Insight box ── */
.bi-box {
    background: #0C1829;
    border: 1px solid #1A2E44;
    border-left: 3px solid #F39C12;
    border-radius: 8px;
    padding: 14px 16px;
    margin-bottom: 10px;
}
.bi-box.critical { border-left-color: #E74C3C; }
.bi-box.good     { border-left-color: #2ECC71; }
.bi-box.info     { border-left-color: #5DADE2; }
.bi-title {
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    margin-bottom: 4px;
    font-family: 'JetBrains Mono', monospace;
}
.bi-title.critical { color: #E74C3C; }
.bi-title.warn     { color: #F39C12; }
.bi-title.good     { color: #2ECC71; }
.bi-title.info     { color: #5DADE2; }
.bi-text {
    font-size: 13px;
    color: #8AADC2;
    line-height: 1.6;
}

/* ── Top title bar ── */
.top-title {
    font-size: 22px;
    font-weight: 800;
    color: #FFFFFF;
    letter-spacing: 1px;
}
.top-sub {
    font-size: 12px;
    color: #3A6A8A;
    font-family: 'JetBrains Mono', monospace;
    margin-top: 2px;
}

/* ── Plotly chart containers ── */
.chart-container {
    background: #0C1829;
    border: 1px solid #1A2E44;
    border-radius: 10px;
    padding: 16px;
}

/* ── Streamlit overrides ── */
div[data-testid="metric-container"] {
    background: #0C1829;
    border: 1px solid #1A2E44;
    border-radius: 10px;
    padding: 14px;
}
.stTabs [data-baseweb="tab-list"] {
    background: #0C1829;
    border-bottom: 1px solid #1A2E44;
    gap: 0;
}
.stTabs [data-baseweb="tab"] {
    background: transparent;
    color: #3A6A8A;
    font-size: 12px;
    font-weight: 600;
    letter-spacing: 1px;
    text-transform: uppercase;
    padding: 10px 20px;
    border-bottom: 2px solid transparent;
}
.stTabs [aria-selected="true"] {
    color: #5DADE2 !important;
    border-bottom: 2px solid #5DADE2 !important;
    background: transparent !important;
}
.stDataFrame { border: 1px solid #1A2E44 !important; border-radius: 8px; }
.stSelectbox > div > div {
    background: #0C1829 !important;
    border-color: #1A2E44 !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  PLOTLY THEME
# ─────────────────────────────────────────────
DARK_LAYOUT = dict(
    paper_bgcolor="#0C1829",
    plot_bgcolor="#090F1A",
    font=dict(color="#7AAABF", family="JetBrains Mono, monospace", size=11),
    title_font=dict(color="#FFFFFF", size=13, family="Syne, sans-serif"),
    margin=dict(t=50, b=40, l=50, r=20),
)
AXIS = dict(gridcolor="#1A2E44", linecolor="#1A2E44", tickcolor="#3A6A8A", zerolinecolor="#1A2E44")
LEGEND = dict(bgcolor="rgba(0,0,0,0)", bordercolor="#1A2E44", borderwidth=1, font=dict(size=11))
RED, AMBER, GREEN, BLUE = "#E74C3C", "#F39C12", "#2ECC71", "#5DADE2"

# ─────────────────────────────────────────────
#  DATA LOADING
# ─────────────────────────────────────────────
@st.cache_data
def parse_excel(file_bytes: bytes) -> dict:
    xl = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None)
    out = {}

    def clean(df, hrow=1):
        d = df.copy()
        d.columns = d.iloc[hrow]
        d = d.iloc[hrow+1:].reset_index(drop=True)
        d.columns = [str(c).strip() if pd.notna(c) else f"_c{i}" for i,c in enumerate(d.columns)]
        return d.dropna(how="all")

    # CBR Raw
    raw = xl["CBR_Raw"]
    cbr = clean(raw)
    cbr.columns = ["road","station","cbr_loaded","cbr_standard","status"]
    cbr["cbr_loaded"] = pd.to_numeric(cbr["cbr_loaded"], errors="coerce")
    cbr["cbr_standard"] = pd.to_numeric(cbr["cbr_standard"], errors="coerce")
    out["cbr_raw"] = cbr.dropna(subset=["cbr_loaded"])

    # CBR Summary
    cs = clean(xl["CBR_Summary"])
    cs.columns = ["road","n_stations","n_ok","n_fail","pct_fail","cbr_min","cbr_max","cbr_mean","cbr_standard"]
    for c in ["n_stations","n_ok","n_fail","cbr_min","cbr_max","cbr_mean","cbr_standard"]:
        cs[c] = pd.to_numeric(cs[c], errors="coerce")
    out["cbr_summary"] = cs.dropna(subset=["road"])

    # Grade Raw
    gr = clean(xl["Grade_Raw"])
    gr.columns = ["road","station","grade_pct","grade_standard","status"]
    gr["grade_pct"] = pd.to_numeric(gr["grade_pct"], errors="coerce")
    out["grade_raw"] = gr.dropna(subset=["grade_pct"])

    # RDS — split at row with "Summary per Road"
    rds_raw = xl["RDS_Eval"]
    # detail rows: rows 2-11 (0-indexed after header at row 1)
    def rds_detail(raw):
        rows = []
        for i, r in raw.iterrows():
            v0 = str(r.iloc[0]).strip()
            if v0 in ("MonteBawah","Spanyol"):
                rows.append({
                    "road": v0,
                    "defect": str(r.iloc[1]).strip(),
                    "degree": pd.to_numeric(r.iloc[2], errors="coerce"),
                    "extent": pd.to_numeric(r.iloc[3], errors="coerce"),
                    "defect_score": pd.to_numeric(r.iloc[4], errors="coerce"),
                    "contribution_pct": str(r.iloc[5]).replace("%",""),
                })
        df = pd.DataFrame(rows)
        df["contribution_pct"] = pd.to_numeric(df["contribution_pct"], errors="coerce")
        return df

    def rds_summary(raw):
        rows = []
        found = False
        for i, r in raw.iterrows():
            if "Summary per Road" in str(r.iloc[0]):
                found = True; continue
            if found and str(r.iloc[0]).strip() in ("MonteBawah","Spanyol"):
                rows.append({
                    "road": str(r.iloc[0]).strip(),
                    "total_rds": pd.to_numeric(r.iloc[1], errors="coerce"),
                    "rr_pct": pd.to_numeric(r.iloc[2], errors="coerce"),
                    "rr_target": pd.to_numeric(r.iloc[3], errors="coerce"),
                    "rr_ok": str(r.iloc[4]),
                    "dominant_defect": str(r.iloc[5]).strip(),
                })
        return pd.DataFrame(rows)

    out["rds_detail"]  = rds_detail(rds_raw)
    out["rds_summary"] = rds_summary(rds_raw)

    # Giroud-Han
    gh = clean(xl["Giroud_Han"])
    gh.columns = ["road","station","cbr_sg","r_contact","ph0_kN","RE","Fe","h_base_m"]
    for c in ["cbr_sg","r_contact","ph0_kN","RE","Fe","h_base_m"]:
        gh[c] = pd.to_numeric(gh[c], errors="coerce")
    out["gh"] = gh.dropna(subset=["cbr_sg"])

    # Rimpull / Speed
    rs = clean(xl["Rimpull_Speed"])
    rs.columns = ["road","station","grade_before","grade_after","rr_before","rr_after",
                  "TR_before","TR_after","gear_before","gear_after","speed_before","speed_after","speed_delta"]
    for c in ["grade_before","grade_after","rr_before","rr_after","TR_before","TR_after",
              "gear_before","gear_after","speed_before","speed_after","speed_delta"]:
        rs[c] = pd.to_numeric(rs[c], errors="coerce")
    # Drop summary rows (no station value)
    out["rimpull"] = rs.dropna(subset=["station","speed_before"]).query("road in ['MonteBawah','Spanyol']")

    def sp_summary(raw):
        rows = []
        found = False
        for i, r in raw.iterrows():
            if "Speed Summary" in str(r.iloc[0]): found = True; continue
            if found and str(r.iloc[0]).strip() in ("MonteBawah","Spanyol"):
                rows.append({
                    "road": str(r.iloc[0]).strip(),
                    "avg_before": pd.to_numeric(r.iloc[1], errors="coerce"),
                    "avg_after": pd.to_numeric(r.iloc[2], errors="coerce"),
                    "avg_delta": pd.to_numeric(r.iloc[3], errors="coerce"),
                })
        return pd.DataFrame(rows)
    out["speed_summary"] = sp_summary(xl["Rimpull_Speed"])

    # Cut & Fill
    cf = clean(xl["Cut_Fill"])
    cf.columns = ["road","station","grade_before","grade_after","delta_grade","dh_m","volume_m3","work_type"]
    for c in ["grade_before","grade_after","delta_grade","dh_m","volume_m3"]:
        cf[c] = pd.to_numeric(cf[c], errors="coerce")
    out["cut_fill"] = cf.dropna(subset=["station","volume_m3"]).query("road in ['MonteBawah','Spanyol']")

    def cf_summary(raw):
        rows = []
        found = False
        for i, r in raw.iterrows():
            if "Volume Summary" in str(r.iloc[0]): found = True; continue
            if found and str(r.iloc[0]).strip() in ("MonteBawah","Spanyol"):
                rows.append({
                    "road": str(r.iloc[0]).strip(),
                    "total_cut": pd.to_numeric(r.iloc[1], errors="coerce"),
                    "total_fill": pd.to_numeric(r.iloc[2], errors="coerce"),
                    "net_m3": pd.to_numeric(r.iloc[3], errors="coerce"),
                    "n_cut": pd.to_numeric(r.iloc[4], errors="coerce"),
                    "n_fill": pd.to_numeric(r.iloc[5], errors="coerce"),
                })
        return pd.DataFrame(rows)
    out["cf_summary"] = cf_summary(xl["Cut_Fill"])

    # Method Notes
    mn = xl["Method_Notes"]
    notes = []
    for _, r in mn.iterrows():
        k = str(r.iloc[0]).strip() if pd.notna(r.iloc[0]) else ""
        v = str(r.iloc[1]).strip() if pd.notna(r.iloc[1]) else ""
        if k and k != "nan":
            notes.append({"key": k, "value": v})
    out["notes"] = pd.DataFrame(notes)

    return out

# ─────────────────────────────────────────────
#  BI INSIGHT GENERATOR
# ─────────────────────────────────────────────
def generate_bi(data: dict) -> list[dict]:
    insights = []
    cs = data["cbr_summary"]
    rs = data["rds_summary"]
    sp = data["speed_summary"]
    cf = data.get("cf_summary", pd.DataFrame())
    gh = data["gh"]

    # CBR
    for _, row in cs.iterrows():
        insights.append({
            "phase": "CBR", "severity": "critical", "road": row["road"],
            "title": f"CBR 100% FAIL — {row['road']}",
            "text": (f"Semua {int(row['n_stations'])} station di {row['road']} di bawah standar CBR 39%. "
                     f"CBR aktual rata-rata {row['cbr_mean']:.1f}% (terendah {row['cbr_min']:.0f}%, tertinggi {row['cbr_max']:.0f}%). "
                     f"Gap rata-rata {39 - row['cbr_mean']:.1f}% di bawah standar — subgrade tidak mampu menahan beban HD 785 170 ton "
                     f"tanpa perkuatan base course.")
        })

    # CBR gap analysis
    all_cbr = data["cbr_raw"]
    worst = all_cbr.loc[all_cbr["cbr_loaded"].idxmin()]
    insights.append({
        "phase": "CBR", "severity": "warn", "road": "All",
        "title": "Titik Kritis CBR",
        "text": (f"Station terlemah: {worst['road']} sta {worst['station']} dengan CBR {worst['cbr_loaded']:.0f}% "
                 f"— hanya {worst['cbr_loaded']/39*100:.0f}% dari standar minimum. "
                 f"Potensi kerusakan parah dan penurunan daya dukung subgrade jika tanpa perbaikan segera.")
    })

    # RDS
    for _, row in rs.iterrows():
        insights.append({
            "phase": "RDS", "severity": "critical", "road": row["road"],
            "title": f"Rolling Resistance {row['rr_pct']}% — {row['road']}",
            "text": (f"RR aktual {row['rr_pct']}% melebihi target perusahaan 2.0% sebesar "
                     f"{row['rr_pct']-2.0:.1f}%. Defect dominan: {row['dominant_defect']}. "
                     f"Setiap kenaikan 1% RR pada HD 785 berbobot 374.785 lb meningkatkan traction requirement "
                     f"±{374785*0.01/2000:.0f} ton-force, mempercepat wear engine dan meningkatkan konsumsi BBM.")
        })

    # Speed improvement
    for _, row in sp.iterrows():
        pct = (row["avg_delta"] / row["avg_before"]) * 100
        insights.append({
            "phase": "Speed", "severity": "good", "road": row["road"],
            "title": f"Proyeksi Peningkatan Speed +{row['avg_delta']:.1f} km/h — {row['road']}",
            "text": (f"Setelah perbaikan, rata-rata speed HD 785 di {row['road']} naik dari "
                     f"{row['avg_before']:.1f} km/h ke {row['avg_after']:.1f} km/h (+{pct:.0f}%). "
                     f"Peningkatan ini langsung berpengaruh pada produktivitas: jika jarak hauling 850m, "
                     f"cycle time turun ±{850/row['avg_before']*60-850/row['avg_after']*60:.1f} menit per trip.")
        })

    # GH base course
    avg_h = gh["h_base_m"].mean()
    max_h = gh["h_base_m"].max()
    worst_gh = gh.loc[gh["h_base_m"].idxmax()]
    insights.append({
        "phase": "Giroud-Han", "severity": "warn", "road": "All",
        "title": "Ketebalan Base Course — Rekomendasi Giroud-Han",
        "text": (f"Tebal base course rata-rata yang dibutuhkan: {avg_h:.2f}m. "
                 f"Paling tebal di {worst_gh['road']} sta {worst_gh['station']} = {max_h:.2f}m "
                 f"(CBR subgrade {worst_gh['cbr_sg']:.0f}%). Material base course yang direkomendasikan: "
                 f"EWT Floor 100 dari Pit Wara (CBR ≥ 80%). Estimasi volume material per meter panjang jalan "
                 f"(lebar 25m): {avg_h*25:.1f} m³/m — perlu dikroscek dengan anggaran material.")
    })

    # Cut & Fill
    if not cf.empty:
        for _, row in cf.iterrows():
            net_type = "cut dominan" if row["net_m3"] < 0 else "fill dominan"
            insights.append({
                "phase": "Cut & Fill", "severity": "info", "road": row["road"],
                "title": f"Pekerjaan Tanah {row['road']} — {net_type.upper()}",
                "text": (f"Total cut: {row['total_cut']:,.0f} m³ ({int(row['n_cut'])} station). "
                         f"Total fill: {row['total_fill']:,.0f} m³ ({int(row['n_fill'])} station). "
                         f"Net volume: {abs(row['net_m3']):,.0f} m³ {'harus didatangkan dari luar' if row['net_m3'] > 0 else 'excess material bisa dibuang/dimanfaatkan'}. "
                         f"Potensi optimasi: gabungkan pekerjaan cut di awal segmen sebagai sumber material fill di segmen lain.")
            })

    # Overall BI summary
    total_stations = int(cs["n_stations"].sum())
    insights.append({
        "phase": "Summary", "severity": "info", "road": "All",
        "title": "Executive Summary — Kondisi Keseluruhan",
        "text": (f"Kedua jalan (MonteBawah & Spanyol) dalam kondisi KRITIS: {total_stations} dari {total_stations} "
                 f"station tidak memenuhi standar CBR, RR aktual {rs['rr_pct'].mean():.1f}% vs target 2%, "
                 f"dan ada {int((data['grade_raw']['grade_pct'].abs() > 8).sum())} station grade melebihi batas 8%. "
                 f"Prioritas perbaikan: (1) perkuatan base course Giroud-Han, (2) grading ulang station over-grade, "
                 f"(3) perbaikan permukaan untuk turunkan RDS. ROI estimasi: kenaikan speed rata-rata "
                 f"{sp['avg_delta'].mean():.1f} km/h setara peningkatan produktivitas hauling ±{sp['avg_delta'].mean()/sp['avg_before'].mean()*100:.0f}%.")
    })

    return insights


# ─────────────────────────────────────────────
#  CHART HELPERS
# ─────────────────────────────────────────────
def cbr_bar_chart(df_raw, road, std=39):
    sub = df_raw[df_raw["road"] == road].copy()
    colors = [RED if v < std else GREEN for v in sub["cbr_loaded"]]
    fig = go.Figure()
    fig.add_hline(y=std, line_dash="dash", line_color="#3A8AC8", line_width=1.5,
                  annotation_text=f"Standard {std}%", annotation_font_color="#3A8AC8",
                  annotation_position="top right")
    fig.add_trace(go.Bar(
        x=sub["station"], y=sub["cbr_loaded"],
        marker_color=colors, marker_line_width=0,
        hovertemplate="<b>%{x}</b><br>CBR: %{y}%<extra></extra>",
    ))
    fig.update_layout(**DARK_LAYOUT, title=f"CBR per Station — {road}",
                      xaxis_title="Station (m)", yaxis_title="CBR (%)")
    fig.update_xaxes(**AXIS, tickangle=45)
    fig.update_yaxes(**AXIS, range=[0, 50])
    return fig


def grade_bar_chart(df_raw, road, std=8):
    sub = df_raw[df_raw["road"] == road].copy()
    colors = [RED if abs(v) > std else GREEN for v in sub["grade_pct"]]
    fig = go.Figure()
    fig.add_hline(y=std, line_dash="dash", line_color="#3A8AC8", line_width=1.5,
                  annotation_text=f"Max {std}%", annotation_font_color="#3A8AC8",
                  annotation_position="top right")
    fig.add_trace(go.Bar(
        x=sub["station"], y=sub["grade_pct"].abs(),
        marker_color=colors, marker_line_width=0,
        hovertemplate="<b>%{x}</b><br>Grade: %{y:.2f}%<extra></extra>",
    ))
    fig.update_layout(**DARK_LAYOUT, title=f"Grade per Station — {road}",
                      xaxis_title="Station (m)", yaxis_title="Grade (%)")
    fig.update_xaxes(**AXIS, tickangle=45)
    fig.update_yaxes(**AXIS)
    return fig


def rds_horizontal_chart(df_detail, road):
    sub = df_detail[df_detail["road"] == road].sort_values("defect_score", ascending=True)
    palette = [RED, AMBER, "#F39C12", GREEN, BLUE]
    fig = go.Figure()
    for i, (_, row) in enumerate(sub.iterrows()):
        fig.add_trace(go.Bar(
            y=[row["defect"]], x=[row["defect_score"]],
            orientation="h",
            marker_color=palette[i % len(palette)],
            name=row["defect"],
            hovertemplate=f"<b>{row['defect']}</b><br>Score: {row['defect_score']}<br>Kontribusi: {row['contribution_pct']:.1f}%<extra></extra>",
        ))
    fig.update_layout(**DARK_LAYOUT, title=f"RDS Breakdown — {road}",
                      xaxis_title="Defect Score (Degree × Extent)",
                      showlegend=False, xaxis_range=[0, 20])
    fig.update_xaxes(**AXIS)
    fig.update_yaxes(**AXIS)
    return fig


def speed_comparison_chart(df_ri, road):
    sub = df_ri[df_ri["road"] == road].copy()
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Sebelum Perbaikan", x=sub["station"], y=sub["speed_before"],
        marker_color=RED, marker_opacity=0.85,
        hovertemplate="<b>%{x}</b><br>Speed before: %{y} km/h<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        name="Setelah Perbaikan", x=sub["station"], y=sub["speed_after"],
        marker_color=GREEN, marker_opacity=0.85,
        hovertemplate="<b>%{x}</b><br>Speed after: %{y} km/h<extra></extra>",
    ))
    fig.update_layout(**DARK_LAYOUT, title=f"Speed Comparison — {road}",
                      barmode="group", xaxis_title="Station (m)", yaxis_title="Speed (km/h)",
                      legend=LEGEND)
    fig.update_xaxes(**AXIS, tickangle=45)
    fig.update_yaxes(**AXIS)
    return fig


def cut_fill_chart(df_cf, road):
    sub = df_cf[df_cf["road"] == road].copy()
    sub["color"] = sub["work_type"].map({"CUT": RED, "FILL": GREEN, "NONE": "#3A6A8A"})
    fig = go.Figure()
    cut = sub[sub["work_type"] == "CUT"]
    fill = sub[sub["work_type"] == "FILL"]
    none_ = sub[sub["work_type"] == "NONE"]
    if not cut.empty:
        fig.add_trace(go.Bar(name="CUT", x=cut["station"], y=cut["volume_m3"],
                             marker_color=RED,
                             hovertemplate="<b>%{x}</b><br>Cut: %{y:.0f} m³<extra></extra>"))
    if not fill.empty:
        fig.add_trace(go.Bar(name="FILL", x=fill["station"], y=fill["volume_m3"],
                             marker_color=GREEN,
                             hovertemplate="<b>%{x}</b><br>Fill: %{y:.0f} m³<extra></extra>"))
    if not none_.empty:
        fig.add_trace(go.Bar(name="NONE", x=none_["station"], y=none_["volume_m3"],
                             marker_color="#3A6A8A",
                             hovertemplate="<b>%{x}</b><br>Volume: 0 m³<extra></extra>"))
    fig.update_layout(**DARK_LAYOUT, title=f"Cut & Fill Volume — {road}",
                      barmode="group", xaxis_title="Station (m)", yaxis_title="Volume (m³)",
                      legend=LEGEND)
    fig.update_xaxes(**AXIS, tickangle=45)
    fig.update_yaxes(**AXIS)
    return fig


def gh_chart(df_gh, road):
    sub = df_gh[df_gh["road"] == road].copy()
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        name="CBR Subgrade (%)", x=sub["station"], y=sub["cbr_sg"],
        marker_color="#3A8AC8", marker_opacity=0.6,
        hovertemplate="<b>%{x}</b><br>CBR: %{y}%<extra></extra>",
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        name="Base Course h (m)", x=sub["station"], y=sub["h_base_m"],
        mode="lines+markers", line=dict(color=AMBER, width=2.5),
        marker=dict(size=7, color=AMBER),
        hovertemplate="<b>%{x}</b><br>h base: %{y:.2f}m<extra></extra>",
    ), secondary_y=True)
    fig.add_hline(y=0.30, line_dash="dot", line_color=GREEN, line_width=1,
                  annotation_text="h_min 0.30m", annotation_font_color=GREEN,
                  secondary_y=True)
    fig.update_layout(**DARK_LAYOUT, title=f"Giroud-Han Base Course — {road}", legend=LEGEND)
    fig.update_xaxes(**AXIS, tickangle=45, title_text="Station (m)")
    fig.update_yaxes(**AXIS, title_text="CBR (%)", secondary_y=False)
    fig.update_yaxes(**AXIS, title_text="Base Course h (m)", secondary_y=True)
    return fig


def rr_gauge(rr_val, target=2.0, road=""):
    color = RED if rr_val > target else GREEN
    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=rr_val,
        delta={"reference": target, "increasing": {"color": RED}, "decreasing": {"color": GREEN}},
        number={"suffix": "%", "font": {"color": color, "size": 36, "family": "JetBrains Mono"}},
        gauge={
            "axis": {"range": [0, 6], "tickcolor": "#3A6A8A", "tickfont": {"size": 10}},
            "bar": {"color": color, "thickness": 0.25},
            "bgcolor": "#090F1A",
            "bordercolor": "#1A2E44",
            "steps": [
                {"range": [0, 2], "color": "#0A2A15"},
                {"range": [2, 3.5], "color": "#2A1A08"},
                {"range": [3.5, 6], "color": "#2A1010"},
            ],
            "threshold": {"line": {"color": "#3A8AC8", "width": 2}, "thickness": 0.8, "value": target},
        },
        title={"text": f"RR — {road}", "font": {"color": "#7AAABF", "size": 12}},
    ))
    fig.update_layout(paper_bgcolor="#0C1829", height=200,
                      margin=dict(t=40, b=10, l=20, r=20),
                      font=dict(color="#7AAABF"))
    return fig


# ─────────────────────────────────────────────
#  EXCEL EXPORT
# ─────────────────────────────────────────────
def export_excel(data: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr_fmt = wb.add_format({"bold": True, "bg_color": "#0C1829", "font_color": "#5DADE2",
                                  "border": 1, "border_color": "#1A2E44", "font_name": "Consolas"})
        cell_fmt = wb.add_format({"bg_color": "#090F1A", "font_color": "#8AADC2",
                                   "border": 1, "border_color": "#1A2E44", "font_name": "Consolas"})
        fail_fmt = wb.add_format({"bg_color": "#2A1010", "font_color": "#E74C3C",
                                   "border": 1, "border_color": "#1A2E44", "font_name": "Consolas"})
        ok_fmt   = wb.add_format({"bg_color": "#0A2A15", "font_color": "#2ECC71",
                                   "border": 1, "border_color": "#1A2E44", "font_name": "Consolas"})

        sheets = {
            "CBR Raw": data["cbr_raw"],
            "CBR Summary": data["cbr_summary"],
            "Grade Raw": data["grade_raw"],
            "RDS Detail": data["rds_detail"],
            "RDS Summary": data["rds_summary"],
            "Giroud-Han": data["gh"],
            "Rimpull-Speed": data["rimpull"],
            "Speed Summary": data["speed_summary"],
            "Cut Fill": data["cut_fill"],
            "CF Summary": data.get("cf_summary", pd.DataFrame()),
        }
        for sname, df in sheets.items():
            if df.empty:
                continue
            df.to_excel(writer, sheet_name=sname[:31], index=False)
            ws = writer.sheets[sname[:31]]
            for ci, col in enumerate(df.columns):
                ws.write(0, ci, col, hdr_fmt)
                ws.set_column(ci, ci, max(14, len(str(col))+2))

    return buf.getvalue()


# ─────────────────────────────────────────────
#  SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🛣️ ROAD HAUL DAMAGE ANALYZER")
    st.markdown("**PT PPA Adaro · Komatsu HD 785-7**")
    st.divider()

    st.markdown("### Upload Data")
    uploaded = st.file_uploader("Upload report.xlsx", type=["xlsx"],
                                 help="Upload hasil output pipeline (semua sheet: CBR_Raw, Grade_Raw, RDS_Eval, Giroud_Han, Rimpull_Speed, Cut_Fill)")

    st.divider()
    st.markdown("### Filter")
    road_filter = st.selectbox("Jalan", ["Semua", "MonteBawah", "Spanyol"])

    dark_toggle = st.toggle("Dark Mode", value=True)

    st.divider()
    st.markdown("### Standards")
    st.markdown("""
    <div style='font-family:monospace; font-size:11px; color:#3A6A8A; line-height:2'>
    CBR min &nbsp;&nbsp;: 39%<br>
    Grade max : 8%<br>
    RR target : 2.0%<br>
    Truck &nbsp;&nbsp;&nbsp;&nbsp;: HD 785-7<br>
    GVW loaded: 170 ton
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  LOAD DATA
# ─────────────────────────────────────────────
if uploaded is None:
    st.markdown("""
    <div style='text-align:center; padding: 80px 40px;'>
        <div style='font-size:48px; margin-bottom:16px'>🛣️</div>
        <div class='top-title'>ROAD HAUL DAMAGE ANALYZER</div>
        <div class='top-sub' style='margin-top:8px'>PT PPA Adaro Indonesia · Komatsu HD 785-7 · Pit Wara, Kalimantan Selatan</div>
        <div style='margin-top:32px; font-size:14px; color:#3A6A8A; font-family:monospace'>
            ← Upload <b style='color:#5DADE2'>report.xlsx</b> di sidebar untuk memulai
        </div>
        <div style='margin-top:12px; font-size:12px; color:#1A3A5A; font-family:monospace'>
            Full pipeline: CBR Eval · Grade Eval · RDS Scoring · Giroud-Han · Rimpull/Speed · Cut & Fill
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

data = parse_excel(uploaded.read())

# Apply road filter
def filt(df, col="road"):
    if road_filter != "Semua":
        return df[df[col] == road_filter]
    return df

roads = ["MonteBawah", "Spanyol"] if road_filter == "Semua" else [road_filter]

# ─────────────────────────────────────────────
#  TOP BAR
# ─────────────────────────────────────────────
st.markdown("""
<div style='display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:20px; padding:18px 0 12px 0; border-bottom:1px solid #1A2E44'>
  <div>
    <div class='top-title'>ROAD HAUL DAMAGE ANALYZER</div>
    <div class='top-sub'>PT PPA Adaro Indonesia · Komatsu HD 785-7 · Full Pipeline Phase 1–4</div>
  </div>
  <div style='font-family:monospace; font-size:11px; color:#3A6A8A; text-align:right'>
    Reference: Figo Febriawan et al.<br>Jurnal Teknologi Pertambangan, 2025
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  KPI ROW
# ─────────────────────────────────────────────
cs = data["cbr_summary"]
rs = data["rds_summary"]
sp = data["speed_summary"]
cf_sum = data.get("cf_summary", pd.DataFrame())
gh = data["gh"]

total_sta    = int(filt(cs)["n_stations"].sum())
total_fail   = int(filt(cs)["n_fail"].sum())
avg_rr       = filt(rs)["rr_pct"].mean()
avg_spd_b    = filt(sp)["avg_before"].mean()
avg_spd_a    = filt(sp)["avg_after"].mean()
avg_h_base   = filt(gh)["h_base_m"].mean()

k1, k2, k3, k4, k5, k6 = st.columns(6)
for col, label, val, sub, cls in [
    (k1, "Total Station", f"{total_sta}", "MonteBawah + Spanyol", "kpi-blue"),
    (k2, "CBR Fail", f"{total_fail}/{total_sta}", "100% fail — all below 39%", "kpi-red"),
    (k3, "RR Aktual", f"{avg_rr:.1f}%", f"target 2.0% | over +{avg_rr-2:.1f}%", "kpi-red"),
    (k4, "Speed Before", f"{avg_spd_b:.1f} km/h", "rata-rata loaded", "kpi-amber"),
    (k5, "Speed After ↑", f"{avg_spd_a:.1f} km/h", f"+{avg_spd_a-avg_spd_b:.1f} km/h improvement", "kpi-green"),
    (k6, "Avg h Base", f"{avg_h_base:.2f}m", "Giroud-Han required", "kpi-amber"),
]:
    col.markdown(f"""
    <div class='kpi-card'>
      <div class='kpi-label'>{label}</div>
      <div class='kpi-value {cls}'>{val}</div>
      <div class='kpi-sub sub-{"red" if "red" in cls else "green" if "green" in cls else "amber" if "amber" in cls else "blue"}'>{sub}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  MAIN TABS
# ─────────────────────────────────────────────
tab_overview, tab_cbr, tab_grade, tab_rds, tab_gh, tab_speed, tab_cf, tab_bi, tab_data, tab_notes = st.tabs([
    "📊 Overview", "🪨 CBR", "📐 Grade", "🛞 RDS Scoring",
    "🏗️ Giroud-Han", "🚛 Rimpull/Speed", "⛏️ Cut & Fill",
    "💡 BI Insights", "📋 Data Tables", "📝 Method Notes"
])

# ──────────────────────────────
#  OVERVIEW TAB
# ──────────────────────────────
with tab_overview:
    st.markdown("<div class='sec-header'>Kondisi Ringkas — Semua Jalan</div>", unsafe_allow_html=True)

    # RR Gauges
    g1, g2 = st.columns(2)
    for col, road in zip([g1, g2], roads):
        rr_row = rs[rs["road"] == road]
        if not rr_row.empty:
            col.plotly_chart(rr_gauge(rr_row["rr_pct"].values[0], road=road),
                             use_container_width=True, key=f"gauge_ov_{road}")

    # CBR + Grade side by side per road
    for road in roads:
        st.markdown(f"<div class='sec-header'>{road}</div>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        c1.plotly_chart(cbr_bar_chart(data["cbr_raw"], road), use_container_width=True, key=f"cbr_ov_{road}")
        c2.plotly_chart(grade_bar_chart(data["grade_raw"], road), use_container_width=True, key=f"grade_ov_{road}")

# ──────────────────────────────
#  CBR TAB
# ──────────────────────────────
with tab_cbr:
    st.markdown("<div class='sec-header'>CBR Evaluation — DCP per Station</div>", unsafe_allow_html=True)

    # Summary table
    cs_show = filt(cs).copy()
    cs_show.columns = ["Jalan","Stations","OK","Fail","Fail %","CBR Min","CBR Max","CBR Mean","Standard"]
    st.dataframe(cs_show, use_container_width=True, hide_index=True)
    st.markdown("<br>", unsafe_allow_html=True)

    for road in roads:
        st.plotly_chart(cbr_bar_chart(data["cbr_raw"], road), use_container_width=True, key=f"cbr_tab_{road}")

        sub = data["cbr_raw"][data["cbr_raw"]["road"] == road]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Min CBR", f"{sub['cbr_loaded'].min():.0f}%", f"{sub['cbr_loaded'].min()-39:.0f}% vs std")
        c2.metric("Max CBR", f"{sub['cbr_loaded'].max():.0f}%", f"{sub['cbr_loaded'].max()-39:.0f}% vs std")
        c3.metric("Mean CBR", f"{sub['cbr_loaded'].mean():.1f}%", f"{sub['cbr_loaded'].mean()-39:.1f}% vs std")
        c4.metric("Gap to Standard", f"{39-sub['cbr_loaded'].mean():.1f}%", "below minimum")
        st.divider()

# ──────────────────────────────
#  GRADE TAB
# ──────────────────────────────
with tab_grade:
    st.markdown("<div class='sec-header'>Grade Evaluation — Per 50m Station</div>", unsafe_allow_html=True)

    gr = filt(data["grade_raw"])
    over = gr[gr["grade_pct"].abs() > 8]
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Stations", len(gr))
    c2.metric("Over Grade (>8%)", len(over), f"{len(over)/len(gr)*100:.0f}% of total")
    c3.metric("Worst Grade", f"{gr['grade_pct'].abs().max():.2f}%",
              f"sta {gr.loc[gr['grade_pct'].abs().idxmax(), 'station']} — {gr.loc[gr['grade_pct'].abs().idxmax(), 'road']}")

    st.markdown("<br>", unsafe_allow_html=True)
    for road in roads:
        st.plotly_chart(grade_bar_chart(data["grade_raw"], road), use_container_width=True, key=f"grade_tab_{road}")

        sub_over = data["grade_raw"][(data["grade_raw"]["road"]==road) & (data["grade_raw"]["grade_pct"].abs()>8)]
        if not sub_over.empty:
            st.markdown(f"**Stations over-grade di {road}:**")
            st.dataframe(sub_over[["station","grade_pct","grade_standard","status"]],
                         use_container_width=True, hide_index=True)
        st.divider()

# ──────────────────────────────
#  RDS TAB
# ──────────────────────────────
with tab_rds:
    st.markdown("<div class='sec-header'>RDS Scoring — Qualitative Rolling Resistance (Thompson, 2011)</div>",
                unsafe_allow_html=True)

    rs_show = filt(rs).copy()
    st.dataframe(rs_show, use_container_width=True, hide_index=True)
    st.markdown("<br>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    for col, road in zip([c1, c2], roads):
        with col:
            col.plotly_chart(rds_horizontal_chart(data["rds_detail"], road), use_container_width=True, key=f"rds_tab_{road}")
            rds_row = rs[rs["road"] == road]
            if not rds_row.empty:
                r = rds_row.iloc[0]
                col.markdown(f"""
                <div class='bi-box critical'>
                  <div class='bi-title critical'>RDS={int(r['total_rds'])} → RR={r['rr_pct']}%</div>
                  <div class='bi-text'>
                  Dominant defect: <b>{r['dominant_defect']}</b><br>
                  Target RR ≤ {r['rr_target']}% — saat ini over +{r['rr_pct']-r['rr_target']:.1f}%
                  </div>
                </div>
                """, unsafe_allow_html=True)

    st.markdown("<br>")
    st.markdown("<div class='sec-header'>Tabel Detail RDS per Defect</div>", unsafe_allow_html=True)
    det_show = filt(data["rds_detail"])
    st.dataframe(det_show, use_container_width=True, hide_index=True)

# ──────────────────────────────
#  GIROUD-HAN TAB
# ──────────────────────────────
with tab_gh:
    st.markdown("<div class='sec-header'>Giroud-Han Pavement Design — Base Course Thickness</div>",
                unsafe_allow_html=True)

    gh_f = filt(gh)
    c1, c2, c3 = st.columns(3)
    c1.metric("Avg h Base Course", f"{gh_f['h_base_m'].mean():.3f}m")
    c2.metric("Max h Base Course", f"{gh_f['h_base_m'].max():.2f}m",
              f"sta {gh_f.loc[gh_f['h_base_m'].idxmax(),'station']}")
    c3.metric("Min h Base Course", f"{gh_f['h_base_m'].min():.2f}m")

    st.markdown("<br>", unsafe_allow_html=True)
    for road in roads:
        st.plotly_chart(gh_chart(gh, road), use_container_width=True, key=f"gh_tab_{road}")

    st.markdown("<div class='sec-header'>Data Giroud-Han Lengkap</div>", unsafe_allow_html=True)
    st.dataframe(gh_f[["road","station","cbr_sg","r_contact","ph0_kN","RE","Fe","h_base_m"]],
                 use_container_width=True, hide_index=True)

# ──────────────────────────────
#  SPEED TAB
# ──────────────────────────────
with tab_speed:
    st.markdown("<div class='sec-header'>Rimpull & Speed Analysis — Before vs After Repair</div>",
                unsafe_allow_html=True)

    sp_f = filt(sp)
    for _, row in sp_f.iterrows():
        c1, c2, c3 = st.columns(3)
        c1.metric(f"Avg Speed Aktual — {row['road']}", f"{row['avg_before']:.2f} km/h")
        c2.metric(f"Avg Speed Improved — {row['road']}", f"{row['avg_after']:.2f} km/h",
                  f"+{row['avg_delta']:.2f} km/h")
        c3.metric("Improvement", f"+{row['avg_after']-row['avg_before']:.1f} km/h",
                  f"+{(row['avg_after']-row['avg_before'])/row['avg_before']*100:.0f}%")

    st.markdown("<br>", unsafe_allow_html=True)
    for road in roads:
        st.plotly_chart(speed_comparison_chart(data["rimpull"], road), use_container_width=True, key=f"speed_tab_{road}")

    st.markdown("<div class='sec-header'>Data Rimpull Detail</div>", unsafe_allow_html=True)
    ri_f = filt(data["rimpull"])
    st.dataframe(ri_f[["road","station","grade_before","grade_after","rr_before","rr_after",
                         "TR_before","TR_after","gear_before","gear_after","speed_before","speed_after","speed_delta"]],
                 use_container_width=True, hide_index=True)

# ──────────────────────────────
#  CUT & FILL TAB
# ──────────────────────────────
with tab_cf:
    st.markdown("<div class='sec-header'>Cut & Fill Volume Estimation — Grade Repair</div>",
                unsafe_allow_html=True)

    if not cf_sum.empty:
        cf_f = filt(cf_sum)
        for _, row in cf_f.iterrows():
            c1, c2, c3, c4 = st.columns(4)
            c1.metric(f"Total Cut — {row['road']}", f"{row['total_cut']:,.0f} m³",
                      f"{int(row['n_cut'])} stations")
            c2.metric(f"Total Fill — {row['road']}", f"{row['total_fill']:,.0f} m³",
                      f"{int(row['n_fill'])} stations")
            c3.metric("Net Volume", f"{abs(row['net_m3']):,.0f} m³",
                      "excess cut" if row["net_m3"] < 0 else "fill needed")
            c4.metric("Net Type", "CUT DOM." if row["net_m3"] < 0 else "FILL DOM.")

    st.markdown("<br>", unsafe_allow_html=True)
    for road in roads:
        st.plotly_chart(cut_fill_chart(data["cut_fill"], road), use_container_width=True, key=f"cf_tab_{road}")

    st.markdown("<div class='sec-header'>Data Cut & Fill Detail</div>", unsafe_allow_html=True)
    cf_det = filt(data["cut_fill"])
    st.dataframe(cf_det, use_container_width=True, hide_index=True)

# ──────────────────────────────
#  BI INSIGHTS TAB
# ──────────────────────────────
with tab_bi:
    st.markdown("<div class='sec-header'>Business Intelligence Insights — Analisis & Rekomendasi</div>",
                unsafe_allow_html=True)

    insights = generate_bi(data)

    # Filter by road
    if road_filter != "Semua":
        insights = [i for i in insights if i["road"] in [road_filter, "All"]]

    # Phase filter
    phases = list(dict.fromkeys(i["phase"] for i in insights))
    sel_phase = st.multiselect("Filter Phase", phases, default=phases, key="bi_phase")
    insights_show = [i for i in insights if i["phase"] in sel_phase]

    sev_colors = {"critical": "critical", "warn": "warn", "good": "good", "info": "info"}
    for ins in insights_show:
        sev = ins["severity"]
        cls = sev_colors.get(sev, "info")
        box_cls = "critical" if sev == "critical" else ("good" if sev == "good" else "")
        st.markdown(f"""
        <div class='bi-box {box_cls}' style='margin-bottom:12px'>
          <div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:6px'>
            <div class='bi-title {cls}'>{ins["title"]}</div>
            <div style='display:flex; gap:6px'>
              <span style='font-size:10px; padding:2px 8px; border-radius:4px; background:#1A2E44; color:#3A6A8A; font-family:monospace'>{ins["phase"]}</span>
              <span style='font-size:10px; padding:2px 8px; border-radius:4px; background:#1A2E44; color:#3A6A8A; font-family:monospace'>{ins["road"]}</span>
            </div>
          </div>
          <div class='bi-text'>{ins["text"]}</div>
        </div>
        """, unsafe_allow_html=True)

# ──────────────────────────────
#  DATA TABLES TAB
# ──────────────────────────────
with tab_data:
    st.markdown("<div class='sec-header'>Semua Data Tables — Raw & Summary</div>", unsafe_allow_html=True)

    t1, t2, t3, t4, t5 = st.tabs(["CBR","Grade","RDS","Giroud-Han","Rimpull & CF"])

    with t1:
        st.dataframe(filt(data["cbr_raw"]), use_container_width=True, hide_index=True)
    with t2:
        st.dataframe(filt(data["grade_raw"]), use_container_width=True, hide_index=True)
    with t3:
        col1, col2 = st.columns(2)
        col1.dataframe(filt(data["rds_detail"]), use_container_width=True, hide_index=True)
        col2.dataframe(filt(data["rds_summary"]), use_container_width=True, hide_index=True)
    with t4:
        st.dataframe(filt(data["gh"]), use_container_width=True, hide_index=True)
    with t5:
        st.subheader("Rimpull / Speed")
        st.dataframe(filt(data["rimpull"]), use_container_width=True, hide_index=True)
        st.subheader("Cut & Fill")
        st.dataframe(filt(data["cut_fill"]), use_container_width=True, hide_index=True)

    # Export
    st.markdown("<br>", unsafe_allow_html=True)
    excel_bytes = export_excel(data)
    st.download_button(
        label="⬇ Export Semua Data ke Excel",
        data=excel_bytes,
        file_name="road_damage_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ──────────────────────────────
#  METHOD NOTES TAB
# ──────────────────────────────
with tab_notes:
    st.markdown("<div class='sec-header'>Method Notes & Assumptions</div>", unsafe_allow_html=True)

    notes = data["notes"]
    phases_map = {
        "PHASE 1 — CBR Evaluation": "🪨 Phase 1 — CBR",
        "PHASE 1 — Grade Evaluation": "📐 Phase 1 — Grade",
        "PHASE 2 — RDS Scoring": "🛞 Phase 2 — RDS",
        "PHASE 3 — Pavement Design": "🏗️ Phase 3 — Giroud-Han",
        "PHASE 4 — Rimpull & Speed": "🚛 Phase 4 — Rimpull",
        "PHASE 4 — Cut & Fill": "⛏️ Phase 4 — Cut & Fill",
    }
    current = None
    for _, row in notes.iterrows():
        k = row["key"]
        v = row["value"]
        if k in phases_map:
            current = phases_map[k]
            st.markdown(f"<div class='sec-header' style='margin-top:18px'>{current}</div>",
                        unsafe_allow_html=True)
        elif k not in ("nan","NaN") and k:
            col1, col2 = st.columns([2, 5])
            col1.markdown(f"<div style='font-family:monospace; font-size:12px; color:#3A6A8A; padding:4px 0'>{k}</div>",
                          unsafe_allow_html=True)
            col2.markdown(f"<div style='font-family:monospace; font-size:12px; color:#7AAABF; padding:4px 0'>{v if v and v!='nan' else '—'}</div>",
                          unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  FOOTER
# ─────────────────────────────────────────────
st.markdown("""
<div style='text-align:center; padding:24px; margin-top:32px;
  border-top:1px solid #1A2E44; font-family:monospace; font-size:11px; color:#1A3A5A'>
  ROAD HAUL DAMAGE ANALYZER · DEVOLEPT by NICKHOLAST ADITYA · PT PPA Adaro Indonesia · Pit Wara, Tabalong, Kalimantan Selatan<br>
  Ref: Figo Febriawan et al., Jurnal Teknologi Pertambangan Vol.10 No.2, Jan 2025 |
  Method: DCP · Thompson RDS · Giroud-Han (Suwandhi, 2004) · Komatsu HD 785-7 Rimpull Chart
</div>
""", unsafe_allow_html=True)