# Decoding Sunrisers Hyderabad's IPL 2026 Campaign

## A Ball-by-Ball Performance Analysis

**AICTE | IBM SkillsBuild Data Analytics with AI Internship 2026**

---

## Author

**Lanka Saimanikanta Sharma**

---

## Description

This project analyses Sunrisers Hyderabad's (SRH) performance during the IPL 2026
season using five CSV datasets covering ball-by-ball deliveries, batting and
bowling aggregates, fielding statistics, and venue metadata.

The analysis is based strictly on the observed IPL 2026 data available in the
provided datasets. The project focuses on descriptive and comparative analytics
and does not perform win/loss modelling, outcome prediction, or multi-season
extrapolation.

The primary dataset is the ball-by-ball deliveries file. The supporting datasets
provide additional context for batting, bowling, fielding, and venue analysis.

The project follows a data analytics workflow:

**Data → Cleaning → Analysis → Visualisation → Insights → Observations**

---

## Dataset Overview

| File | Rows | Columns | Key Fields | SRH Relevance |
|---|---:|---:|---|---|
| `deliveries (1).csv` | 17,477 | 19 | match_no, date, stage, venue, batting_team, bowling_team, innings, over, striker, bowler, runs_of_bat, extras, wide, legbyes, byes, noballs, wicket_type, player_dismissed, fielder | Primary dataset containing SRH delivery-level data |
| `batting_stats.csv` | 15 | 17 | batsman, team, runs, average, strike_rate, impact, hundreds, fifties, fours, sixes | 3 SRH players in the top-15 batting summary |
| `bowling.csv` | 15 | 10 | Player, Team, Wickets, Economy, Average, Overs | 2 SRH players in the top-15 bowling summary |
| `fielding_stats.csv` | 15 | 6 | player, team, matches, catches, cpm | 2 SRH players in the top-15 fielding summary |
| `venues.csv` | 12 | 5 | venue_stadium, city, state, capacity, home_team | Contains SRH home-ground metadata |

### SRH Coverage

The delivery-level dataset contains SRH records across **15 matches**.

The analysis uses the deliveries dataset as the primary source for:

- Batting performance
- Bowling performance
- Phase-wise analysis
- Over-by-over analysis
- Wicket patterns
- Venue analysis
- Player-level delivery statistics
- Stage-level comparisons

The aggregate datasets are used as supporting sources where applicable.

---

## Dataset Source

**Source URL:**

> **TO BE VERIFIED BEFORE SUBMISSION**
>
> The original dataset source URL was not present in the provided workspace
> files. The verified original source URL should be added here before the
> project is submitted or published.
>
> Do not replace this placeholder with an unverified URL.

---

## Project File Structure

```text
srh_ipl2026_analysis/
│
├── srh_analysis.ipynb
├── report_generator.py
├── findings.json
├── README.md
├── requirements.txt
│
├── visuals/
│   ├── V-01_srh_batsmen_runs.png
│   ├── V-02_srh_vs_league_strike_rate.png
│   ├── V-03_run_type_composition.png
│   ├── V-04_extras_breakdown.png
│   ├── V-05_srh_bowlers_wickets.png
│   ├── V-06_economy_comparison.png
│   ├── V-07_dot_ball_pct.png
│   ├── V-08_extras_per_over.png
│   ├── V-09_srh_catch_leaders.png
│   ├── V-10_cpm_comparison.png
│   ├── V-11_batting_rr_by_phase.png
│   ├── V-12_bowling_economy_by_phase.png
│   ├── V-13_ovo_batting_rr.png
│   ├── V-14_ovo_bowling_economy.png
│   ├── V-15_wickets_by_phase.png
│   ├── V-16_srh_wickets_taken_type.png
│   ├── V-17_srh_wickets_lost_type.png
│   ├── V-18_srh_bowlers_ranked.png
│   ├── V-19_bowler_wicket_heatmap.png
│   ├── V-20_venue_batting_rr.png
│   ├── V-21_venue_bowling_economy.png
│   ├── V-22_home_vs_away.png
│   ├── V-23_metrics_by_stage.png
│   ├── V-24_avg_vs_sr_scatter.png
│   ├── V-25_bowler_combined.png
│   ├── V-26_dual_contribution_radar.png
│   └── V-27_squad_summary.png
│
└── report/
    ├── srh_report.html
    ├── srh_report.pdf
    └── srh_report.docx
