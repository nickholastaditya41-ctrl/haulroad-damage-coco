import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from pathlib import Path

OUT = Path(__file__).parent.parent / "outputs/charts"
COLORS = ["#E74C3C","#E67E22","#F39C12","#27AE60","#2980B9"]

def plot_rds_breakdown(df_detail, df_summary):
    OUT.mkdir(parents=True, exist_ok=True)
    roads = df_detail["road"].unique()
    fig, axes = plt.subplots(1, len(roads), figsize=(12, 5))
    if len(roads) == 1:
        axes = [axes]

    for ax, road in zip(axes, roads):
        sub = df_detail[df_detail["road"] == road].sort_values("defect_score", ascending=False)
        bars = ax.barh(sub["defect"], sub["defect_score"],
                       color=COLORS[:len(sub)], edgecolor="white", height=0.6)
        for bar, pct in zip(bars, sub["contribution_pct"]):
            ax.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height()/2,
                    f"{pct:.1f}%", va="center", fontsize=9)
        rr = df_summary.loc[df_summary["road"]==road, "rr_pct"].values[0]
        rds = df_summary.loc[df_summary["road"]==road, "total_rds"].values[0]
        ax.set_title(f"{road}\nRDS={rds}  →  RR={rr}%", fontsize=11, fontweight="bold")
        ax.set_xlabel("Defect Score (Degree × Extent)", fontsize=9)
        ax.axvline(x=2.0, color="red", linestyle="--", alpha=0.0)  # placeholder
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.set_xlim(0, 20)

    fig.suptitle("RDS Breakdown — Qualitative Rolling Resistance Assessment", fontsize=13, fontweight="bold", y=1.01)
    plt.tight_layout()
    path = OUT / "rds_breakdown.png"
    plt.savefig(path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"[chart_rds] Chart saved: {path}")
