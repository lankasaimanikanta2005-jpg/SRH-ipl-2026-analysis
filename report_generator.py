"""
report_generator.py
Generates srh_report.pdf (fpdf2) and srh_report.docx (python-docx) for:

    "Decoding Sunrisers Hyderabad's IPL 2026 Campaign:
     A Ball-by-Ball Performance Analysis"
    Author: Lanka Saimanikanta Sharma
    Programme: AICTE | IBM SkillsBuild Data Analytics with AI Internship 2026

PREREQUISITE
------------
Run srh_analysis.ipynb completely (Kernel > Restart and Run All).
Section 14 of the notebook writes findings.json to this directory.
This script loads findings.json at runtime and embeds the ACTUAL computed
metric values in every narrative sentence.

If findings.json is absent this script exits immediately with a clear error.
It NEVER falls back to placeholder, representative, or fabricated values.

Usage (from srh_ipl2026_analysis/):
    python report_generator.py

Requirements: fpdf2>=2.7.0  python-docx>=1.0.0
"""

import json
import sys
from pathlib import Path

from fpdf import FPDF
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ── Paths ──────────────────────────────────────────────────────────────────
VISUALS_DIR   = Path("visuals")
REPORT_DIR    = Path("report")
FINDINGS_PATH = Path("findings.json")
PDF_PATH      = REPORT_DIR / "srh_report.pdf"
DOCX_PATH     = REPORT_DIR / "srh_report.docx"

# ── Fixed metadata ─────────────────────────────────────────────────────────
AUTHOR      = "Lanka Saimanikanta Sharma"
TITLE       = ("Decoding Sunrisers Hyderabad's IPL 2026 Campaign: "
               "A Ball-by-Ball Performance Analysis")
TITLE_SHORT = "SRH IPL 2026 Analysis"
ATTRIBUTION = "AICTE | IBM SkillsBuild Data Analytics with AI Internship 2026"
SEASON      = "IPL 2026 (March 28 \u2013 May 9, 2026)"

SOURCE_URL_PLACEHOLDER = (
    "[Requires user verification before submission \u2014 the original dataset "
    "source URL was not found in any file within this workspace. Please confirm "
    "the source and add the verified URL here before submission.]"
)

# ── Dataset overview (static, from verified file inspection) ───────────────
DATASET_TABLE = [
    ("deliveries (1).csv", "17,477", "19",
     "Ball-by-ball delivery records", "Primary \u2014 all SRH match data"),
    ("batting_stats.csv",  "15",     "17",
     "Top-15 IPL 2026 batting aggregates", "3 SRH players included"),
    ("bowling.csv",        "15",     "10",
     "Top-15 IPL 2026 bowling aggregates",
     "2 SRH players; Best Figures excluded"),
    ("fielding_stats.csv", "15",     "6",
     "Top-15 IPL 2026 catch leaders", "2 SRH players included"),
    ("venues.csv",         "12",     "5",
     "IPL 2026 stadium metadata", "1 SRH home ground"),
]

DATA_QUALITY_NOTES = [
    ("Best Figures corruption", "bowling.csv",
     "Values like '4/32' read as dates (e.g. 'Apr-32'). Column excluded entirely."),
    ("Team name mismatch", "bowling.csv",
     "Full names vs abbreviations in other files. Resolved via TEAM_MAP."),
    ("Venue name mismatch", "deliveries vs venues.csv",
     "Long strings vs canonical names. 12/13 venues mapped; Raipur = Unmapped."),
    ("Extras repeat ball label", "deliveries",
     "Wide/no-ball rows share the over.N label of interrupted delivery. Validated before use."),
    ("Dismissal columns sparse", "deliveries",
     "~16,600 empty values in wicket_type/player_dismissed/fielder — expected cricket structure."),
    ("hitwicket spelling", "deliveries",
     "Two spellings normalised to 'hit wicket'."),
    ("Dataset source URL", "All files",
     "No source URL found in workspace. Placeholder retained — user must verify."),
]

LIMITATIONS = [
    "Single season (IPL 2026) only \u2014 results cannot be generalised to other seasons.",
    "Batting, bowling, and fielding summary files cover only the top-15 players in each category.",
    "No match result column exists in any dataset \u2014 win/loss analysis is not supported.",
    "No toss data \u2014 toss-related analysis is not supported.",
    ("Fielding statistics are limited to catches; stumpings, run-outs, "
     "and direct hits are not captured."),
    "No player role field exists in any dataset \u2014 players cannot be classified by role.",
    ("Shaheed Veer Narayan Singh International Stadium (Raipur) is absent from "
     "venues.csv \u2014 no venue metadata available for that ground."),
    "The Best Figures column in bowling.csv is corrupted and excluded from all calculations.",
]

# ── 18 curated charts for PDF and DOCX ────────────────────────────────────
CHARTS_FOR_REPORT = [
    ("V-01_srh_batsmen_runs.png",
     "V-01: SRH Batsmen by Runs Scored (IPL 2026, \u226510 legal balls)"),
    ("V-02_srh_vs_league_strike_rate.png",
     "V-02: SRH Top-3 vs League Top-15 Average Strike Rate (batting_stats.csv)"),
    ("V-03_run_type_composition.png",
     "V-03: Run Type Composition per SRH Batsman (\u226510 legal balls, deliveries)"),
    ("V-05_srh_bowlers_wickets.png",
     "V-05: SRH Bowlers by Wickets Taken (IPL 2026, \u22656 legal deliveries)"),
    ("V-06_economy_comparison.png",
     "V-06: Economy Rate \u2014 Top-15 Bowlers (SRH highlighted; Best Figures excluded)"),
    ("V-07_dot_ball_pct.png",
     "V-07: SRH Bowler Dot Ball Percentage (deliveries, legal balls only)"),
    ("V-09_srh_catch_leaders.png",
     "V-09: Catch Leaders \u2014 Top-15 (SRH highlighted, fielding_stats.csv)"),
    ("V-10_cpm_comparison.png",
     "V-10: Catches per Match \u2014 Top-15 (SRH highlighted)"),
    ("V-11_batting_rr_by_phase.png",
     "V-11: Batting Run Rate by Phase \u2014 SRH vs. Tournament Average"),
    ("V-12_bowling_economy_by_phase.png",
     "V-12: Bowling Economy by Phase \u2014 SRH vs. Tournament Average"),
    ("V-13_ovo_batting_rr.png",
     "V-13: SRH Batting Run Rate Over-by-Over (all 15 matches aggregated)"),
    ("V-16_srh_wickets_taken_type.png",
     "V-16: SRH Wickets Taken by Dismissal Type"),
    ("V-17_srh_wickets_lost_type.png",
     "V-17: SRH Wickets Lost by Dismissal Type"),
    ("V-18_srh_bowlers_ranked.png",
     "V-18: SRH Bowlers Ranked by Wickets Taken (deliveries)"),
    ("V-20_venue_batting_rr.png",
     "V-20: SRH Batting Run Rate by Venue (home venue = orange)"),
    ("V-23_metrics_by_stage.png",
     "V-23: SRH Batting and Bowling Metrics by Stage"),
    ("V-26_dual_contribution_radar.png",
     "V-26: Dual-Contribution Player Profile (normalised metrics, deliveries)"),
    ("V-27_squad_summary.png",
     "V-27: SRH Squad Performance Summary (all 5 datasets)"),
]  # exactly 18 entries

# All 27 chart filenames (all present in visuals/ and HTML export)
ALL_CHART_FILES = [
    "V-01_srh_batsmen_runs.png",
    "V-02_srh_vs_league_strike_rate.png",
    "V-03_run_type_composition.png",
    "V-04_extras_breakdown.png",
    "V-05_srh_bowlers_wickets.png",
    "V-06_economy_comparison.png",
    "V-07_dot_ball_pct.png",
    "V-08_extras_per_over.png",
    "V-09_srh_catch_leaders.png",
    "V-10_cpm_comparison.png",
    "V-11_batting_rr_by_phase.png",
    "V-12_bowling_economy_by_phase.png",
    "V-13_ovo_batting_rr.png",
    "V-14_ovo_bowling_economy.png",
    "V-15_wickets_by_phase.png",
    "V-16_srh_wickets_taken_type.png",
    "V-17_srh_wickets_lost_type.png",
    "V-18_srh_bowlers_ranked.png",
    "V-19_bowler_wicket_heatmap.png",
    "V-20_venue_batting_rr.png",
    "V-21_venue_bowling_economy.png",
    "V-22_home_vs_away.png",
    "V-23_metrics_by_stage.png",
    "V-24_avg_vs_sr_scatter.png",
    "V-25_bowler_combined.png",
    "V-26_dual_contribution_radar.png",
    "V-27_squad_summary.png",
]  # 27 entries

# ── Required keys — every F["key"] used in narrative is listed here ────────
# All keys are written by cell-s14-findings in srh_analysis.ipynb.
REQUIRED_KEYS = {
    "srh_bat_rr", "srh_bat_legal", "srh_bat_total", "srh_bat_runs",
    "srh_bat_extras", "srh_fours", "srh_sixes", "srh_bat_matches",
    "srh_bowl_econ", "srh_bowl_legal", "srh_bowl_total", "srh_wkts_taken_n",
    "highest_bat_phase", "highest_bat_phase_rr", "highest_bat_phase_n",
    "lowest_bowl_phase", "lowest_bowl_phase_econ", "lowest_bowl_phase_n",
    "bat_pp_rr", "bat_pp_n", "bat_mid_rr", "bat_mid_n",
    "bat_death_rr", "bat_death_n",
    "bowl_pp_econ", "bowl_pp_n", "bowl_mid_econ", "bowl_mid_n",
    "bowl_death_econ", "bowl_death_n",
    "abhi_sr", "league_sr_avg",
    "caught_taken_pct", "caught_taken_n", "wkts_taken_total",
    "caught_lost_pct", "caught_lost_n", "wkts_lost_total",
    "highest_bat_venue", "highest_bat_venue_rr",
    "highest_bat_venue_ndel", "highest_bat_venue_nmat",
    "lowest_bowl_venue", "lowest_bowl_venue_econ", "lowest_bowl_venue_ndel",
    "f10_srh_names", "f10_srh_cpm", "league_cpm_avg", "srh_cpm_avg",
    "dual_players", "dual_players_n",
    "srh_dot_pct", "tour_dot_pct",
    "home_rr", "away_rr", "home_legal", "away_legal",
    "srh_impact_avg", "league_impact_avg",
}

# Bowling economy definition note — used verbatim in narratives
BOWL_ECON_DEF = (
    "runs off the bat, wides, and no-ball runs "
    "(byes and leg-byes excluded, as they are charged to the fielding side)"
)
BOWL_DENOM_DEF = (
    "legal deliveries (wides and no-balls excluded) divided by 6"
)


# ═══════════════════════════════════════════════════════════════════════════
#  FINDINGS LOADER
# ═══════════════════════════════════════════════════════════════════════════

def load_findings():
    """
    Load findings.json produced by srh_analysis.ipynb Section 14.
    Exits with a clear, actionable error if:
      - findings.json is absent (notebook not yet run)
      - any required key is missing
      - any required numeric key is None
    Never returns placeholder, representative, or fabricated values.
    """
    if not FINDINGS_PATH.exists():
        sys.exit(
            "\n"
            "ERROR: findings.json not found in this directory.\n"
            "This file is written by srh_analysis.ipynb when Section 14 executes.\n"
            "\n"
            "Steps to fix:\n"
            "  1. Open srh_analysis.ipynb in Jupyter.\n"
            "  2. Kernel > Restart and Run All.\n"
            "  3. Confirm 'findings.json written' appears in the Section 14 output.\n"
            "  4. Re-run:  python report_generator.py\n"
            "\n"
            "This script will NOT generate reports without actual computed values.\n"
        )

    with open(FINDINGS_PATH, encoding="utf-8") as fh:
        F = json.load(fh)

    missing = REQUIRED_KEYS - set(F.keys())
    if missing:
        sys.exit(
            f"\nERROR: findings.json is missing {len(missing)} required key(s):\n"
            + "\n".join(f"  - {k}" for k in sorted(missing))
            + "\n\nRe-run the notebook completely to regenerate findings.json.\n"
        )

    # Validate numeric keys that are formatted with .2f / .1f
    numeric_keys = {
        "srh_bat_rr", "srh_bowl_econ",
        "highest_bat_phase_rr", "lowest_bowl_phase_econ",
        "bat_pp_rr", "bat_mid_rr", "bat_death_rr",
        "bowl_pp_econ", "bowl_mid_econ", "bowl_death_econ",
        "caught_taken_pct", "caught_lost_pct",
        "highest_bat_venue_rr", "lowest_bowl_venue_econ",
        "srh_dot_pct", "tour_dot_pct",
        "home_rr", "away_rr",
        "league_cpm_avg", "srh_cpm_avg",
        "srh_impact_avg", "league_impact_avg",
        "league_sr_avg",
    }
    bad = [k for k in numeric_keys if F.get(k) is None]
    if bad:
        sys.exit(
            f"\nERROR: The following numeric keys are None in findings.json:\n"
            + "\n".join(f"  - {k}" for k in sorted(bad))
            + "\nRe-run the notebook and verify the Section 14 output.\n"
        )

    print(f"findings.json loaded successfully: {len(F)} keys.")
    return F


def check_charts():
    """Warn about missing report charts; do not abort."""
    missing = [f for f, _ in CHARTS_FOR_REPORT
               if not (VISUALS_DIR / f).exists()]
    if missing:
        print(f"WARNING: {len(missing)} of 18 report chart(s) not found in visuals/:")
        for m in sorted(missing):
            print(f"  {m}")
        print("Run the notebook first. Missing charts appear as error notes in reports.")
    else:
        print("All 18 report charts confirmed present in visuals/.")
    return missing


# ═══════════════════════════════════════════════════════════════════════════
#  NARRATIVE BUILDERS  — every sentence uses F["key"], no hardcoded metrics
# ═══════════════════════════════════════════════════════════════════════════

def _dual_str(F):
    d = F["dual_players"]
    return ", ".join(d) if d else "none identified in the deliveries dataset"


def build_executive_summary(F):
    return (
        f"This report presents a data analytics study of Sunrisers Hyderabad's (SRH) "
        f"performance during the Indian Premier League (IPL) 2026 season "
        f"(March 28 \u2013 May 9, 2026). Five CSV datasets are analysed: a ball-by-ball "
        f"deliveries file (17,477 rows across 74 matches and all 10 IPL teams), batting "
        f"and bowling aggregate summaries (top-15 players each), fielding catch statistics "
        f"(top-15 catch leaders), and stadium metadata. SRH participated in "
        f"{F['srh_bat_matches']} matches spanning the League stage, Playoffs, and Final.\n\n"
        f"SRH's batting sub-dataset contains {F['srh_bat_legal']:,} legal deliveries "
        f"and their bowling sub-dataset contains {F['srh_bowl_legal']:,} legal deliveries. "
        f"Batting run rate is computed as (runs off bat + all extras received) / "
        f"(legal deliveries / 6). "
        f"Bowling economy is computed as ({BOWL_ECON_DEF}) / ({BOWL_DENOM_DEF}). "
        f"This definition matches standard cricket scoring convention.\n\n"
        f"Key analytical areas covered: batting, bowling, fielding, phase breakdown "
        f"(Powerplay / Middle / Death), wicket pattern analysis, venue analysis, "
        f"stage comparison (League vs Playoffs/Final), player profiles, and "
        f"cross-dataset integration. "
        f"{F['dual_players_n']} player(s) have delivery records in both the SRH batting "
        f"and SRH bowling subsets ({_dual_str(F)}). No role field exists in any dataset; "
        f"this is a data-presence observation only.\n\n"
        f"No win/loss modelling, outcome predictions, or multi-season claims are made. "
        f"The corrupted Best Figures column from bowling.csv is excluded from all analysis."
    )


def build_batting_narrative(F):
    abhi_line = (
        f"Abhishek Sharma recorded a strike rate of {F['abhi_sr']:.1f} "
        f"(league top-15 average SR: {F['league_sr_avg']:.1f}, batting_stats.csv)."
        if F["abhi_sr"] is not None
        else
        "Abhishek Sharma's name was not matched in batting_stats.csv; "
        "refer to the notebook output for the exact value."
    )
    return (
        f"SRH accumulated {F['srh_bat_total']:,} total runs across "
        f"{F['srh_bat_matches']} matches ({F['srh_bat_runs']:,} runs off the bat, "
        f"{F['srh_bat_extras']:,} extras received). "
        f"The batting run rate \u2014 computed as total runs divided by "
        f"legal deliveries / 6 \u2014 was {F['srh_bat_rr']:.2f} "
        f"(N\u202f=\u202f{F['srh_bat_legal']:,} legal deliveries). "
        f"SRH struck {F['srh_fours']} fours and {F['srh_sixes']} sixes. "
        f"The dot ball percentage was {F['srh_dot_pct']:.1f}% of legal deliveries faced "
        f"(tournament-wide average: {F['tour_dot_pct']:.1f}%).\n\n"
        f"{abhi_line} "
        f"SRH top-3 batsmen average impact score: {F['srh_impact_avg']:.2f} "
        f"(league top-15 average: {F['league_impact_avg']:.2f}; "
        f"n\u202f=\u202f3 SRH / n\u202f=\u202f15 total, batting_stats.csv)."
    )


def build_bowling_narrative(F):
    return (
        f"Bowling runs conceded are defined as {BOWL_ECON_DEF}. "
        f"The economy rate denominator is {BOWL_DENOM_DEF}. "
        f"This follows standard cricket scoring convention.\n\n"
        f"SRH conceded {F['srh_bowl_total']:,} bowling runs across "
        f"{F['srh_bat_matches']} matches. "
        f"The bowling economy rate was {F['srh_bowl_econ']:.2f} "
        f"(N\u202f=\u202f{F['srh_bowl_legal']:,} legal deliveries). "
        f"SRH took {F['srh_wkts_taken_n']} wickets while bowling.\n\n"
        f"The Best Figures column from bowling.csv is completely excluded from all "
        f"calculations, tables, and visualisations due to a documented date-parse "
        f"corruption (source values like '4/32' were interpreted as dates). "
        f"The bowling.csv file lists 2 SRH bowlers among the top-15 IPL 2026 "
        f"bowling aggregates; their economy rates are compared against the "
        f"league top-15 average in the chart below."
    )


def build_fielding_narrative(F):
    names = F["f10_srh_names"]
    cpms  = F["f10_srh_cpm"]
    pairs = ", ".join(f"{n} (cpm\u202f{c:.3f})" for n, c in zip(names, cpms))
    return (
        f"The fielding_stats.csv file records catches for the top-15 catch leaders "
        f"in IPL 2026. SRH has 2 players in this list: {pairs}. "
        f"The league top-15 average catches-per-match (cpm) is {F['league_cpm_avg']:.3f}; "
        f"the average cpm for the 2 SRH players listed is {F['srh_cpm_avg']:.3f}. "
        f"This comparison is bounded to the top-15 catch leaders only.\n\n"
        f"DATA LIMITATION: fielding_stats.csv records catches only. "
        f"Stumpings, run-outs, and direct hits are not captured. "
        f"The fielder column in deliveries is populated only for certain dismissal "
        f"types (caught, stumped, and some run-outs). Fielding analysis is therefore "
        f"limited to catch statistics and fielder-involvement counts from wicket deliveries."
    )


def build_phase_narrative(F):
    return (
        f"Each delivery is classified into one of three phases based on over number: "
        f"Powerplay (overs 0\u20135), Middle (overs 6\u201315), Death (overs 16\u201319). "
        f"All economy rates use {BOWL_ECON_DEF} as the numerator "
        f"and {BOWL_DENOM_DEF} as the denominator. "
        f"Sample sizes (N\u202f= legal deliveries) are stated for every metric.\n\n"
        f"SRH batting run rates by phase: "
        f"Powerplay {F['bat_pp_rr']:.2f} (N\u202f=\u202f{F['bat_pp_n']:,}), "
        f"Middle {F['bat_mid_rr']:.2f} (N\u202f=\u202f{F['bat_mid_n']:,}), "
        f"Death {F['bat_death_rr']:.2f} (N\u202f=\u202f{F['bat_death_n']:,}). "
        f"The phase with the highest observed SRH batting run rate was the "
        f"{F['highest_bat_phase']} phase "
        f"({F['highest_bat_phase_rr']:.2f}, N\u202f=\u202f{F['highest_bat_phase_n']:,} "
        f"legal deliveries).\n\n"
        f"SRH bowling economy rates by phase: "
        f"Powerplay {F['bowl_pp_econ']:.2f} (N\u202f=\u202f{F['bowl_pp_n']:,}), "
        f"Middle {F['bowl_mid_econ']:.2f} (N\u202f=\u202f{F['bowl_mid_n']:,}), "
        f"Death {F['bowl_death_econ']:.2f} (N\u202f=\u202f{F['bowl_death_n']:,}). "
        f"The phase with the lowest observed SRH bowling economy was the "
        f"{F['lowest_bowl_phase']} phase "
        f"({F['lowest_bowl_phase_econ']:.2f}, N\u202f=\u202f{F['lowest_bowl_phase_n']:,} "
        f"legal deliveries). "
        f"All phase comparisons are benchmarked against tournament-wide averages "
        f"computed from the full deliveries dataset."
    )


def build_wickets_narrative(F):
    return (
        f"SRH took {F['wkts_taken_total']} wickets while bowling. "
        f"Of these, {F['caught_taken_n']} were caught dismissals "
        f"({F['caught_taken_pct']:.1f}% of total wickets taken, deliveries dataset). "
        f"Full dismissal-type distributions, per-bowler wicket crosstabs, and "
        f"phase-wise wicket counts are shown in the charts below.\n\n"
        f"SRH lost {F['wkts_lost_total']} wickets while batting. "
        f"Of these, {F['caught_lost_n']} were caught dismissals "
        f"({F['caught_lost_pct']:.1f}% of total wickets lost, deliveries dataset). "
        f"Per-batsman dismissal records are presented in the notebook."
    )


def build_venue_narrative(F):
    return (
        f"SRH played across multiple venues during IPL 2026. "
        f"The venue with the highest observed SRH batting run rate was "
        f"{F['highest_bat_venue']} "
        f"({F['highest_bat_venue_rr']:.2f}, N\u202f=\u202f{F['highest_bat_venue_ndel']:,} "
        f"legal deliveries, {F['highest_bat_venue_nmat']} match(es)). "
        f"The venue with the lowest observed SRH bowling economy "
        f"(using {BOWL_ECON_DEF}) was "
        f"{F['lowest_bowl_venue']} "
        f"({F['lowest_bowl_venue_econ']:.2f}, "
        f"N\u202f=\u202f{F['lowest_bowl_venue_ndel']:,} legal deliveries).\n\n"
        f"Home venue (Rajiv Gandhi International Stadium, Hyderabad): "
        f"SRH batting run rate {F['home_rr']:.2f} "
        f"(N\u202f=\u202f{F['home_legal']:,} legal deliveries). "
        f"Away venues: SRH batting run rate {F['away_rr']:.2f} "
        f"(N\u202f=\u202f{F['away_legal']:,} legal deliveries). "
        f"Shaheed Veer Narayan Singh International Stadium (Raipur) is absent from "
        f"venues.csv and is labelled 'Unmapped'. "
        f"Venue capacity is included as descriptive context only \u2014 "
        f"it is not an analytical variable."
    )


def build_stage_narrative(F):
    return (
        "CAVEAT: The League stage dataset contains substantially more deliveries than "
        "Playoffs and Final combined. All stage comparisons are directional observations "
        "based on the available data and are not statistically conclusive findings. "
        "Delivery counts are stated alongside every stage metric.\n\n"
        f"Bowling economy in this section uses {BOWL_ECON_DEF} "
        f"as the numerator and {BOWL_DENOM_DEF} as the denominator. "
        "Stage metrics include: match count, delivery count, batting run rate, "
        "bowling economy, boundary percentage, dot ball percentage, and wickets. "
        "A Phase \u00d7 Stage pivot table showing legal delivery counts per cell "
        "is presented in the notebook."
    )


def build_player_narrative(F):
    return (
        f"{F['dual_players_n']} player(s) have delivery records in both the SRH batting "
        f"and SRH bowling subsets of the deliveries dataset: {_dual_str(F)}. "
        f"No role field exists in any of the five datasets; this is a data-presence "
        f"observation only. These players are referred to as contributing in both "
        f"batting and bowling, not as 'all-rounders'.\n\n"
        f"Batting profiles cover SRH batsmen with \u226520 legal balls faced. "
        f"Bowling profiles cover bowlers with \u226512 legal deliveries. "
        f"Player bowling economy values use {BOWL_ECON_DEF} "
        f"as the numerator and {BOWL_DENOM_DEF} as the denominator. "
        f"For 3 SRH batsmen in batting_stats.csv and 2 SRH bowlers in bowling.csv, "
        f"delivery-derived metrics are shown alongside summary-file metrics. "
        f"The Best Figures column is excluded from all bowler displays."
    )


def build_cross_dataset_narrative(F):
    return (
        "Four explicit cross-dataset joins demonstrate what each integration enables "
        "beyond what any single dataset could provide alone:\n\n"
        "Integration 1 \u2014 deliveries.striker == batting_stats.batsman: "
        "delivery-level granularity alongside season-aggregate validation for 3 SRH batsmen.\n"
        "Integration 2 \u2014 deliveries.venue_clean == venues.venue_stadium: "
        "city, state, and capacity alongside SRH delivery metrics per venue.\n"
        "Integration 3 \u2014 deliveries.bowler == bowling.Player (after TEAM_MAP): "
        "delivery-derived economy (runs off bat + wides + no-ball runs, byes/leg-byes "
        "excluded) cross-validated against summary economy; Best Figures excluded.\n"
        "Integration 4 \u2014 fielding_stats.player == deliveries.fielder: "
        "catch count (fielding_stats) and dismissal-involvement count (deliveries) "
        "in a single table."
    )


def build_findings_narrative(F):
    abhi_line = (
        f"F5.  Abhishek Sharma strike rate: {F['abhi_sr']:.1f} "
        f"(league top-15 average SR: {F['league_sr_avg']:.1f}, batting_stats.csv).\n"
        if F["abhi_sr"] is not None
        else
        "F5.  Abhishek Sharma SR: not matched in batting_stats.csv "
        "(refer to notebook output).\n"
    )
    dual = _dual_str(F)
    return (
        f"The following 12 metric findings are derived from the executed notebook. "
        f"Every value cites its source dataset and sample size. "
        f"Bowling economy throughout uses {BOWL_ECON_DEF} "
        f"as the numerator and {BOWL_DENOM_DEF} as the denominator. "
        f"All findings are based on IPL 2026 data only (March\u2013May 2026).\n\n"
        f"F1.  SRH overall batting run rate: {F['srh_bat_rr']:.2f} "
        f"(N\u202f=\u202f{F['srh_bat_legal']:,} legal deliveries, deliveries dataset).\n"
        f"F2.  SRH overall bowling economy: {F['srh_bowl_econ']:.2f} "
        f"(N\u202f=\u202f{F['srh_bowl_legal']:,} legal deliveries, deliveries dataset).\n"
        f"F3.  Highest observed SRH batting run rate by phase: {F['highest_bat_phase']} "
        f"({F['highest_bat_phase_rr']:.2f}, N\u202f=\u202f{F['highest_bat_phase_n']:,} "
        f"legal deliveries).\n"
        f"F4.  Lowest observed SRH bowling economy by phase: {F['lowest_bowl_phase']} "
        f"({F['lowest_bowl_phase_econ']:.2f}, N\u202f=\u202f{F['lowest_bowl_phase_n']:,} "
        f"legal deliveries).\n"
        + abhi_line
        + f"F6.  SRH wickets taken as catches: {F['caught_taken_n']} of "
        f"{F['wkts_taken_total']} ({F['caught_taken_pct']:.1f}%, deliveries).\n"
        f"F7.  SRH wickets lost as catches: {F['caught_lost_n']} of "
        f"{F['wkts_lost_total']} ({F['caught_lost_pct']:.1f}%, deliveries).\n"
        f"F8.  Venue with highest observed SRH batting run rate: "
        f"{F['highest_bat_venue']} ({F['highest_bat_venue_rr']:.2f}, "
        f"N\u202f=\u202f{F['highest_bat_venue_ndel']:,} legal deliveries, "
        f"{F['highest_bat_venue_nmat']} match(es), deliveries + venues).\n"
        f"F9.  Venue with lowest observed SRH bowling economy: "
        f"{F['lowest_bowl_venue']} ({F['lowest_bowl_venue_econ']:.2f}, "
        f"N\u202f=\u202f{F['lowest_bowl_venue_ndel']:,} legal deliveries).\n"
        f"F10. SRH fielders catches-per-match vs league top-15 average: "
        f"SRH top-2 avg cpm\u202f=\u202f{F['srh_cpm_avg']:.3f}, "
        f"league top-15 avg cpm\u202f=\u202f{F['league_cpm_avg']:.3f} "
        f"(fielding_stats.csv).\n"
        f"F11. Players with delivery records in both SRH batting and bowling "
        f"subsets: n\u202f=\u202f{F['dual_players_n']} ({dual}, deliveries dataset).\n"
        f"F12. SRH dot ball percentage while batting: {F['srh_dot_pct']:.1f}% "
        f"(tournament average: {F['tour_dot_pct']:.1f}%, deliveries dataset).\n\n"
        f"Batting, bowling, and fielding summary files contain top-15 players only; "
        f"SRH players not in those lists are represented solely through delivery-level data."
    )


def build_conclusion(F):
    dual = _dual_str(F)
    return (
        f"This project applied a structured data analytics workflow to five IPL 2026 "
        f"datasets to characterise Sunrisers Hyderabad's campaign across "
        f"{F['srh_bat_matches']} matches. The ball-by-ball deliveries dataset (17,477 rows) "
        f"served as the primary source, enabling phase-level, over-by-over, per-player, "
        f"and per-venue breakdowns. Supporting files provided aggregate context for "
        f"top-15 batsmen, bowlers, and fielders.\n\n"
        f"Key observed metrics (all from IPL 2026 data): "
        f"SRH batting run rate {F['srh_bat_rr']:.2f} "
        f"(N\u202f=\u202f{F['srh_bat_legal']:,} legal deliveries); "
        f"bowling economy {F['srh_bowl_econ']:.2f} "
        f"(N\u202f=\u202f{F['srh_bowl_legal']:,} legal deliveries; "
        f"economy = {BOWL_ECON_DEF}); "
        f"{F['caught_taken_pct']:.1f}% of SRH wickets taken were caught "
        f"({F['caught_taken_n']} of {F['wkts_taken_total']}); "
        f"{F['dual_players_n']} player(s) had delivery records in both batting and "
        f"bowling subsets ({dual}).\n\n"
        f"All analyses are descriptive and bounded by available data. No win/loss "
        f"outcomes, predictions, or multi-season extrapolations are made. Data quality "
        f"issues \u2014 including the Best Figures corruption, venue name mismatches, "
        f"and team name format differences \u2014 were documented and resolved "
        f"transparently before analysis."
    )


# ═══════════════════════════════════════════════════════════════════════════
#  PDF BUILDER
# ═══════════════════════════════════════════════════════════════════════════

class SRHPdf(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 8)
        self.set_text_color(80, 80, 80)
        self.cell(0, 6, _pdf_safe_text(TITLE_SHORT), align="L", new_x="LMARGIN", new_y="NEXT")
        self.set_text_color(0, 0, 0)

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "I", 7)
        self.set_text_color(100, 100, 100)
        self.cell(0, 5, _pdf_safe_text(ATTRIBUTION), align="C", new_x="LMARGIN", new_y="NEXT")
        self.cell(0, 5, f"Page {self.page_no()}", align="C")
        self.set_text_color(0, 0, 0)


def _pdf_safe_text(text):
    """Convert report text to a Helvetica-safe string."""
    replacements = {
        "\u2013": "-", "\u2014": "-", "\u2011": "-",
        "\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"',
        "\u2022": "-", "\u202f": " ", "\u00a0": " ",
        "\u00d7": "x", "\u2265": ">=", "\u2264": "<=",
    }
    value = str(text)
    for old, new in replacements.items():
        value = value.replace(old, new)
    return value.encode("latin-1", "replace").decode("latin-1")


def _pdf_multicell(pdf, height, text, align="L"):
    """Safe fpdf2 text writer: normalizes Unicode and allows character wrapping."""
    pdf.multi_cell(
        0, height, _pdf_safe_text(text), align=align,
        wrapmode="CHAR", new_x="LMARGIN", new_y="NEXT"
    )


def _ph1(pdf, text):
    pdf.set_font("Helvetica", "B", 14)
    pdf.ln(4)
    _pdf_multicell(pdf, 7, text)
    pdf.ln(2)


def _pbody(pdf, text):
    pdf.set_font("Helvetica", "", 10)
    _pdf_multicell(pdf, 5.5, text)
    pdf.ln(2)


def _pcaption(pdf, text):
    pdf.set_font("Helvetica", "I", 8)
    pdf.set_text_color(80, 80, 80)
    _pdf_multicell(pdf, 5, text, align="C")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(3)


def _pchart(pdf, filename, caption):
    img = VISUALS_DIR / filename
    if not img.exists():
        _pbody(pdf, f"[Chart not generated: {filename} — run notebook first.]")
        return
    pw = pdf.w - pdf.l_margin - pdf.r_margin
    try:
        pdf.image(str(img), x=pdf.l_margin, w=pw)
    except Exception as e:
        _pbody(pdf, f"[Could not embed {filename}: {e}]")
        return
    _pcaption(pdf, caption)


def _ptrow(pdf, cells, widths, bold=False):
    pdf.set_font("Helvetica", "B" if bold else "", 8)
    pdf.set_fill_color(220, 220, 220) if bold else pdf.set_fill_color(255, 255, 255)
    for text, w in zip(cells, widths):
        pdf.cell(w, 6, _pdf_safe_text(str(text)[:42]), border=1, fill=bold,
                 new_x="RIGHT", new_y="TOP")
    pdf.ln()


def build_pdf(F):
    REPORT_DIR.mkdir(exist_ok=True)
    pdf = SRHPdf(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.set_margins(20, 15, 20)

    # Cover
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 18)
    pdf.ln(28)
    _pdf_multicell(pdf, 10, "Decoding Sunrisers Hyderabad's IPL 2026 Campaign", align="C")
    pdf.set_font("Helvetica", "B", 13)
    pdf.ln(2)
    _pdf_multicell(pdf, 8, "A Ball-by-Ball Performance Analysis", align="C")
    pdf.ln(14)
    pdf.set_font("Helvetica", "", 11)
    for line in [f"Author: {AUTHOR}", ATTRIBUTION, SEASON]:
        _pdf_multicell(pdf, 7, line, align="C")
        pdf.ln(1)

    # TOC
    pdf.add_page()
    _ph1(pdf, "Table of Contents")
    toc = [
        "1.  Executive Summary",
        "2.  Dataset Overview",
        "3.  Data Quality Notes",
        "4.  Batting Analysis",
        "5.  Bowling Analysis",
        "6.  Fielding Analysis",
        "7.  Phase Analysis (Powerplay / Middle / Death Overs)",
        "8.  Wicket Pattern Analysis",
        "9.  Venue Analysis",
        "10. League Stage vs. Playoffs Analysis",
        "11. Player Performance Analysis",
        "12. Cross-Dataset Integration",
        "13. Key Findings",
        "14. Limitations and Scope",
        "15. Conclusion",
        "16. References and Dataset Attribution",
    ]
    pdf.set_font("Helvetica", "", 10)
    for entry in toc:
        _pdf_multicell(pdf, 6, entry)
    pdf.ln(3)
    pdf.set_font("Helvetica", "I", 8)
    _pdf_multicell(
        pdf, 5,
        "Page numbers are not shown as they cannot be determined reliably "
        "during programmatic document generation."
    )

    # 1
    pdf.add_page()
    _ph1(pdf, "1. Executive Summary")
    _pbody(pdf, build_executive_summary(F))

    # 2
    pdf.add_page()
    _ph1(pdf, "2. Dataset Overview")
    cols = [46, 14, 12, 54, 44]
    _ptrow(pdf, ["File", "Rows", "Cols", "Description", "SRH Relevance"], cols, bold=True)
    for row in DATASET_TABLE:
        _ptrow(pdf, row, cols)
    pdf.ln(4)
    _pbody(pdf, f"Dataset Source: {SOURCE_URL_PLACEHOLDER}")

    # 3
    pdf.add_page()
    _ph1(pdf, "3. Data Quality Notes")
    qcols = [50, 35, 85]
    _ptrow(pdf, ["Issue", "File", "Resolution"], qcols, bold=True)
    for issue, file_, res in DATA_QUALITY_NOTES:
        _ptrow(pdf, [issue, file_, res], qcols)
    pdf.ln(3)
    _pbody(pdf,
        "Bowling economy convention: runs off the bat + wides + no-ball runs "
        "(byes and leg-byes excluded); denominator = legal deliveries / 6. "
        "The over field uses a decimal encoding validated before use: "
        "integer part = over number (0-19), decimal digit = ball label (1-6).")

    # 4
    pdf.add_page()
    _ph1(pdf, "4. Batting Analysis")
    _pbody(pdf, build_batting_narrative(F))
    for fname, cap in CHARTS_FOR_REPORT[0:3]:
        _pchart(pdf, fname, cap)

    # 5
    pdf.add_page()
    _ph1(pdf, "5. Bowling Analysis")
    _pbody(pdf, build_bowling_narrative(F))
    for fname, cap in CHARTS_FOR_REPORT[3:6]:
        _pchart(pdf, fname, cap)

    # 6
    pdf.add_page()
    _ph1(pdf, "6. Fielding Analysis")
    _pbody(pdf, build_fielding_narrative(F))
    for fname, cap in CHARTS_FOR_REPORT[6:8]:
        _pchart(pdf, fname, cap)

    # 7
    pdf.add_page()
    _ph1(pdf, "7. Phase Analysis (Powerplay / Middle / Death Overs)")
    _pbody(pdf, build_phase_narrative(F))
    for fname, cap in CHARTS_FOR_REPORT[8:11]:
        _pchart(pdf, fname, cap)

    # 8
    pdf.add_page()
    _ph1(pdf, "8. Wicket Pattern Analysis")
    _pbody(pdf, build_wickets_narrative(F))
    for fname, cap in CHARTS_FOR_REPORT[11:14]:
        _pchart(pdf, fname, cap)

    # 9
    pdf.add_page()
    _ph1(pdf, "9. Venue Analysis")
    _pbody(pdf, build_venue_narrative(F))
    _pchart(pdf, *CHARTS_FOR_REPORT[14])

    # 10
    pdf.add_page()
    _ph1(pdf, "10. League Stage vs. Playoffs Analysis")
    _pbody(pdf, build_stage_narrative(F))
    _pchart(pdf, *CHARTS_FOR_REPORT[15])

    # 11
    pdf.add_page()
    _ph1(pdf, "11. Player Performance Analysis")
    _pbody(pdf, build_player_narrative(F))
    for fname, cap in CHARTS_FOR_REPORT[16:18]:
        _pchart(pdf, fname, cap)

    # 12
    pdf.add_page()
    _ph1(pdf, "12. Cross-Dataset Integration")
    _pbody(pdf, build_cross_dataset_narrative(F))
    icols = [32, 66, 72]
    _ptrow(pdf, ["Integration", "Join Key", "Enables"], icols, bold=True)
    for row in [
        ("Integration 1", "deliveries.striker == batting_stats.batsman",
         "Delivery stats + season aggregates for 3 SRH batsmen"),
        ("Integration 2", "deliveries.venue_clean == venues.venue_stadium",
         "City, state, capacity with delivery metrics"),
        ("Integration 3", "deliveries.bowler == bowling.Player (after TEAM_MAP)",
         "Delivery economy vs summary economy (Best Figures excluded)"),
        ("Integration 4", "fielding_stats.player == deliveries.fielder",
         "Catch count + dismissal involvement in one table"),
    ]:
        _ptrow(pdf, row, icols)

    # 13
    pdf.add_page()
    _ph1(pdf, "13. Key Findings")
    _pbody(pdf, build_findings_narrative(F))

    # 14
    pdf.add_page()
    _ph1(pdf, "14. Limitations and Scope")
    pdf.set_font("Helvetica", "", 10)
    for lim in LIMITATIONS:
        _pdf_multicell(pdf, 5.5, f"- {lim}")
        pdf.ln(1)

    # 15
    pdf.add_page()
    _ph1(pdf, "15. Conclusion")
    _pbody(pdf, build_conclusion(F))

    # 16
    pdf.add_page()
    _ph1(pdf, "16. References and Dataset Attribution")
    _pbody(pdf, f"Dataset Source: {SOURCE_URL_PLACEHOLDER}")
    _pbody(pdf, f"Programme:      {ATTRIBUTION}")
    _pbody(pdf, f"Author:         {AUTHOR}")
    _pbody(pdf,
        "Tools: Python 3, pandas, numpy, matplotlib, seaborn, "
        "fpdf2, python-docx, jupyter, nbconvert")

    pdf.output(str(PDF_PATH))
    print(f"PDF  -> {PDF_PATH.resolve()}")


# ═══════════════════════════════════════════════════════════════════════════
#  DOCX BUILDER
# ═══════════════════════════════════════════════════════════════════════════

def _cfg_styles(doc):
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10)
    h1 = doc.styles["Heading 1"]
    h1.font.name, h1.font.size, h1.font.bold = "Calibri", Pt(14), True
    h1.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    h2 = doc.styles["Heading 2"]
    h2.font.name, h2.font.size, h2.font.bold = "Calibri", Pt(12), True


def _set_margins(doc):
    s = doc.sections[0]
    s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(1.0)


def _add_hf(doc):
    section = doc.sections[0]
    section.header.paragraphs[0].text = TITLE_SHORT
    fp = section.footer.paragraphs[0]
    fp.text = ATTRIBUTION
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if fp.runs:
        fp.runs[0].font.size = Pt(8)


def _dbody(doc, text):
    p = doc.add_paragraph(text)
    p.style = doc.styles["Normal"]
    return p


def _dchart(doc, filename, caption):
    img = VISUALS_DIR / filename
    if not img.exists():
        _dbody(doc, f"[Chart not generated: {filename} — run notebook first.]")
        return
    try:
        doc.add_picture(str(img), width=Inches(6.0))
    except Exception as e:
        _dbody(doc, f"[Could not embed {filename}: {e}]")
        return
    cap_p = doc.add_paragraph(caption)
    cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if cap_p.runs:
        cap_p.runs[0].italic    = True
        cap_p.runs[0].font.size = Pt(9)


def _dtable(doc, headers, rows):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = h
        for run in c.paragraphs[0].runs:
            run.bold = True
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            t.rows[ri + 1].cells[ci].text = str(val)[:60]
    doc.add_paragraph()


def _add_toc(doc):
    doc.add_heading("Table of Contents", level=1)
    for entry in [
        "1.  Executive Summary",
        "2.  Dataset Overview",
        "3.  Data Quality Notes",
        "4.  Batting Analysis",
        "5.  Bowling Analysis",
        "6.  Fielding Analysis",
        "7.  Phase Analysis (Powerplay / Middle / Death Overs)",
        "8.  Wicket Pattern Analysis",
        "9.  Venue Analysis",
        "10. League Stage vs. Playoffs Analysis",
        "11. Player Performance Analysis",
        "12. Cross-Dataset Integration",
        "13. Key Findings",
        "14. Limitations and Scope",
        "15. Conclusion",
        "16. References and Dataset Attribution",
    ]:
        p = doc.add_paragraph(entry)
        p.style = doc.styles["Normal"]
        p.paragraph_format.left_indent = Inches(0.25)
    note = doc.add_paragraph(
        "Page numbers are not shown in this Table of Contents as they cannot be "
        "determined reliably during programmatic document generation. "
        "Use the document headings to navigate."
    )
    note.style = doc.styles["Normal"]
    if note.runs:
        note.runs[0].italic    = True
        note.runs[0].font.size = Pt(9)
    doc.add_page_break()


def build_docx(F):
    REPORT_DIR.mkdir(exist_ok=True)
    doc = Document()
    _cfg_styles(doc)
    _set_margins(doc)

    # Cover
    doc.add_paragraph()
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = tp.add_run(
        "Decoding Sunrisers Hyderabad's IPL 2026 Campaign\n"
        "A Ball-by-Ball Performance Analysis"
    )
    r.bold, r.font.size = True, Pt(16)
    doc.add_paragraph()
    for line in [f"Author: {AUTHOR}", ATTRIBUTION, SEASON]:
        p = doc.add_paragraph(line)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p.runs:
            p.runs[0].font.size = Pt(11)
    doc.add_page_break()

    _add_toc(doc)
    _add_hf(doc)

    # 1
    doc.add_heading("1. Executive Summary", level=1)
    _dbody(doc, build_executive_summary(F))
    doc.add_page_break()

    # 2
    doc.add_heading("2. Dataset Overview", level=1)
    _dtable(doc,
        ["File", "Rows", "Cols", "Description", "SRH Relevance"],
        DATASET_TABLE)
    _dbody(doc, f"Dataset Source: {SOURCE_URL_PLACEHOLDER}")
    doc.add_page_break()

    # 3
    doc.add_heading("3. Data Quality Notes", level=1)
    _dtable(doc, ["Issue", "File", "Resolution"], DATA_QUALITY_NOTES)
    _dbody(doc,
        "Bowling economy convention: runs off the bat + wides + no-ball runs "
        "(byes and leg-byes excluded); denominator = legal deliveries / 6.")
    doc.add_page_break()

    # 4
    doc.add_heading("4. Batting Analysis", level=1)
    _dbody(doc, build_batting_narrative(F))
    for fname, cap in CHARTS_FOR_REPORT[0:3]:
        _dchart(doc, fname, cap)
    doc.add_page_break()

    # 5
    doc.add_heading("5. Bowling Analysis", level=1)
    _dbody(doc, build_bowling_narrative(F))
    for fname, cap in CHARTS_FOR_REPORT[3:6]:
        _dchart(doc, fname, cap)
    doc.add_page_break()

    # 6
    doc.add_heading("6. Fielding Analysis", level=1)
    _dbody(doc, build_fielding_narrative(F))
    for fname, cap in CHARTS_FOR_REPORT[6:8]:
        _dchart(doc, fname, cap)
    doc.add_page_break()

    # 7
    doc.add_heading("7. Phase Analysis (Powerplay / Middle / Death Overs)", level=1)
    _dbody(doc, build_phase_narrative(F))
    for fname, cap in CHARTS_FOR_REPORT[8:11]:
        _dchart(doc, fname, cap)
    doc.add_page_break()

    # 8
    doc.add_heading("8. Wicket Pattern Analysis", level=1)
    _dbody(doc, build_wickets_narrative(F))
    for fname, cap in CHARTS_FOR_REPORT[11:14]:
        _dchart(doc, fname, cap)
    doc.add_page_break()

    # 9
    doc.add_heading("9. Venue Analysis", level=1)
    _dbody(doc, build_venue_narrative(F))
    _dchart(doc, *CHARTS_FOR_REPORT[14])
    doc.add_page_break()

    # 10
    doc.add_heading("10. League Stage vs. Playoffs Analysis", level=1)
    _dbody(doc, build_stage_narrative(F))
    _dchart(doc, *CHARTS_FOR_REPORT[15])
    doc.add_page_break()

    # 11
    doc.add_heading("11. Player Performance Analysis", level=1)
    _dbody(doc, build_player_narrative(F))
    for fname, cap in CHARTS_FOR_REPORT[16:18]:
        _dchart(doc, fname, cap)
    doc.add_page_break()

    # 12
    doc.add_heading("12. Cross-Dataset Integration", level=1)
    _dbody(doc, build_cross_dataset_narrative(F))
    _dtable(doc,
        ["Integration", "Join Key", "Enables"],
        [
            ("Integration 1", "deliveries.striker == batting_stats.batsman",
             "Delivery stats + season aggregates for 3 SRH batsmen"),
            ("Integration 2", "deliveries.venue_clean == venues.venue_stadium",
             "City, state, capacity with delivery metrics"),
            ("Integration 3", "deliveries.bowler == bowling.Player (after TEAM_MAP)",
             "Delivery economy vs summary economy (Best Figures excluded)"),
            ("Integration 4", "fielding_stats.player == deliveries.fielder",
             "Catch count + dismissal involvement in one table"),
        ])
    doc.add_page_break()

    # 13
    doc.add_heading("13. Key Findings", level=1)
    _dbody(doc, build_findings_narrative(F))
    doc.add_page_break()

    # 14
    doc.add_heading("14. Limitations and Scope", level=1)
    for lim in LIMITATIONS:
        p = doc.add_paragraph(f"\u2022 {lim}")
        p.style = doc.styles["Normal"]
    doc.add_page_break()

    # 15
    doc.add_heading("15. Conclusion", level=1)
    _dbody(doc, build_conclusion(F))
    doc.add_page_break()

    # 16
    doc.add_heading("16. References and Dataset Attribution", level=1)
    _dbody(doc, f"Dataset Source: {SOURCE_URL_PLACEHOLDER}")
    _dbody(doc, f"Programme:      {ATTRIBUTION}")
    _dbody(doc, f"Author:         {AUTHOR}")
    _dbody(doc,
        "Tools: Python 3, pandas, numpy, matplotlib, seaborn, "
        "fpdf2, python-docx, jupyter, nbconvert")

    doc.save(str(DOCX_PATH))
    print(f"DOCX -> {DOCX_PATH.resolve()}")


# ═══════════════════════════════════════════════════════════════════════════
#  Entry point
# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("=" * 60)
    print("SRH IPL 2026 Report Generator")
    print("=" * 60)

    # Step 1 — load and validate findings (exits on any problem)
    F = load_findings()

    # Step 2 — warn about missing charts (non-fatal)
    check_charts()

    # Step 3 — generate both reports
    print("\nGenerating PDF ...")
    build_pdf(F)

    print("Generating DOCX ...")
    build_docx(F)

    print("\nDone.")
    print(f"  PDF  : {PDF_PATH.resolve()}")
    print(f"  DOCX : {DOCX_PATH.resolve()}")
