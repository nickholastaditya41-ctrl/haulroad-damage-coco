"""
main.py — Project 2: Road Damage Analyzer + Pavement Design
============================================================
Usage:
  python main.py           # Phase 1 only
  python main.py --phase 2 # + RDS scoring
  python main.py --phase 3 # + Giroud-Han
  python main.py --phase 4 # Full pipeline (+ Rimpull + Cut&Fill)
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "modules"))

from loader import load_segments, load_config, load_damage, load_grade_raw
from cbr_eval import evaluate_cbr, summarize_cbr
from grade import evaluate_grade, summarize_grade
from reporter import (write_report, plot_cbr_profile, plot_grade_profile,
                       plot_speed_comparison, plot_cut_fill_chart)

SEP = "═" * 60


def phase1(cfg, verbose=True):
    print(f"\n{SEP}\n  PHASE 1 — CBR & Grade Evaluation\n{SEP}")

    df_seg   = load_segments()
    df_grade_raw = load_grade_raw()

    df_cbr   = evaluate_cbr(df_seg, cfg["standards"]["cbr_min_pct"])
    df_grade = evaluate_grade(df_grade_raw, cfg["standards"]["grade_max_pct"])

    cbr_sum   = summarize_cbr(df_cbr)
    grade_sum = summarize_grade(df_grade)

    if verbose:
        print("\n── CBR Summary ─────────────────────────────")
        print(cbr_sum.to_string(index=False))
        print("\n── Grade Summary ───────────────────────────")
        print(grade_sum.to_string(index=False))

    write_report(df_cbr, df_grade, cbr_sum, grade_sum)
    plot_cbr_profile(df_cbr, cfg["standards"]["cbr_min_pct"])
    plot_grade_profile(df_grade, cfg["standards"]["grade_max_pct"])
    print("\n✓ Phase 1 complete.")
    return df_cbr, df_grade, cbr_sum, grade_sum


def phase2(cfg, df_cbr, df_grade, cbr_sum, grade_sum, verbose=True):
    print(f"\n{SEP}\n  PHASE 2 — RDS Scoring (Qualitative RR Assessment)\n{SEP}")

    from rds_scoring import score_rds
    from chart_rds import plot_rds_breakdown

    df_damage = load_damage()
    df_detail, df_summary = score_rds(df_damage)

    if verbose:
        print("\n── RDS Detail per Defect ────────────────────")
        print(df_detail.to_string(index=False))
        print("\n── RDS Summary per Jalan ────────────────────")
        print(df_summary.to_string(index=False))

    write_report(df_cbr, df_grade, cbr_sum, grade_sum,
                 df_rds_detail=df_detail, df_rds_summary=df_summary)
    plot_rds_breakdown(df_detail, df_summary)
    print("\n✓ Phase 2 complete.")
    return df_detail, df_summary


def phase3(cfg, df_cbr, df_grade, cbr_sum, grade_sum,
           df_rds_detail=None, df_rds_summary=None, verbose=True):
    print(f"\n{SEP}\n  PHASE 3 — Giroud-Han Pavement Design\n{SEP}")

    from giroud_han import compute_base_course

    df_gh = compute_base_course(df_cbr, cfg["giroud_han"], cfg["truck"])

    if verbose:
        print("\n── Giroud-Han Results (Base Course Thickness) ───────────")
        print(df_gh[["road","station_label","cbr_loaded",
                      "Ph0_kN","base_course_h_m"]].to_string(index=False))

    write_report(df_cbr, df_grade, cbr_sum, grade_sum,
                 df_rds_detail=df_rds_detail, df_rds_summary=df_rds_summary,
                 df_gh=df_gh)
    print("\n✓ Phase 3 complete.")
    return df_gh


def phase4(cfg, df_cbr, df_grade, cbr_sum, grade_sum,
           df_rds_detail, df_rds_summary, df_gh, verbose=True):
    print(f"\n{SEP}\n  PHASE 4 — Rimpull/Speed + Cut&Fill (Final Polish)\n{SEP}")

    from rimpull import compute_rimpull
    from cut_fill import compute_cut_fill

    df_rimpull, df_speed_sum = compute_rimpull(df_cbr, df_grade, cfg)
    df_cf, df_cf_sum = compute_cut_fill()

    if verbose:
        print("\n── Speed Summary ────────────────────────────")
        print(df_speed_sum.to_string(index=False))
        print("\n── Cut & Fill Summary ───────────────────────")
        print(df_cf_sum.to_string(index=False))

    write_report(df_cbr, df_grade, cbr_sum, grade_sum,
                 df_rds_detail=df_rds_detail, df_rds_summary=df_rds_summary,
                 df_gh=df_gh,
                 df_rimpull=df_rimpull, df_speed_summary=df_speed_sum,
                 df_cf=df_cf, df_cf_summary=df_cf_sum)

    plot_speed_comparison(df_rimpull, df_speed_sum)
    plot_cut_fill_chart(df_cf)

    print("\n✓ Phase 4 complete.")
    return df_rimpull, df_speed_sum, df_cf, df_cf_sum


def main():
    parser = argparse.ArgumentParser(description="Road Damage Analyzer — PT PPA Adaro")
    parser.add_argument("--phase", type=int, default=1,
                        help="1=CBR+Grade, 2=+RDS, 3=+GiroudHan, 4=full")
    args = parser.parse_args()

    cfg = load_config()
    df_cbr, df_grade, cbr_sum, grade_sum = phase1(cfg)

    df_rds_detail = df_rds_summary = df_gh = None

    if args.phase >= 2:
        df_rds_detail, df_rds_summary = phase2(cfg, df_cbr, df_grade, cbr_sum, grade_sum)

    if args.phase >= 3:
        df_gh = phase3(cfg, df_cbr, df_grade, cbr_sum, grade_sum,
                       df_rds_detail, df_rds_summary)

    if args.phase >= 4:
        phase4(cfg, df_cbr, df_grade, cbr_sum, grade_sum,
               df_rds_detail, df_rds_summary, df_gh)

    print(f"\n{SEP}")
    print(f"  Pipeline Phase {args.phase} selesai.")
    print(f"  Output: outputs/report.xlsx + outputs/charts/")
    print(f"{SEP}\n")


if __name__ == "__main__":
    main()
