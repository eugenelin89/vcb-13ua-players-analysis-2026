from __future__ import annotations

import os
os.environ.setdefault("MPLCONFIGDIR", str((__import__("pathlib").Path(".") / ".mplconfig").resolve()))

from pathlib import Path
import random
from typing import Dict, List, Tuple

import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import FancyBboxPatch
import numpy as np
import pandas as pd

from analysis_pipeline import (
    WORKBOOK_PATH,
    build_draft_board,
    build_hidden_value_table,
    build_master_players,
    build_rankings,
    inspect_workbook,
    team_strength,
)


OUTPUT_PDF = Path("vcb_13u_draft_infographic_report.pdf")
OUTPUT_TEAMS_CSV = Path("suggested_balanced_7_teams.csv")
NUM_TEAMS = 7
ROSTER_SIZE = 11
SEED = 13

BG = "#f7f3ea"
INK = "#152238"
ACCENT = "#c95c32"
ACCENT_2 = "#2c7a7b"
ACCENT_3 = "#d9a441"
MUTED = "#6b7280"
PANEL = "#fffdf8"
GRID = "#d8d2c2"


def style_matplotlib() -> None:
    plt.rcParams.update(
        {
            "figure.facecolor": BG,
            "axes.facecolor": PANEL,
            "axes.edgecolor": GRID,
            "axes.labelcolor": INK,
            "xtick.color": INK,
            "ytick.color": INK,
            "text.color": INK,
            "font.size": 10,
            "axes.titleweight": "bold",
            "axes.titlesize": 14,
        }
    )


def draft_objective(team_frames: List[pd.DataFrame]) -> float:
    summaries = [team_strength(team) for team in team_frames]
    df = pd.DataFrame(summaries)
    if df.empty:
        return float("inf")
    return float(
        df["team_strength_score"].std(ddof=0)
        + 0.55 * df["pitching_strength"].std(ddof=0)
        + 0.35 * df["overall_strength"].std(ddof=0)
        + 0.15 * df["average_balance"].std(ddof=0)
        + 0.20 * df["top_end_talent"].std(ddof=0)
    )


def initial_balanced_assignment(master_players: pd.DataFrame) -> Dict[int, List[str]]:
    ordered = master_players.sort_values(["overall_rank", "player_name"]).reset_index(drop=True)
    teams = {team: [] for team in range(1, NUM_TEAMS + 1)}
    for start in range(0, len(ordered), NUM_TEAMS):
        block = ordered.iloc[start : start + NUM_TEAMS]
        team_order = list(range(1, NUM_TEAMS + 1))
        round_index = start // NUM_TEAMS
        if round_index % 2 == 1:
            team_order.reverse()
        for (_, row), team in zip(block.iterrows(), team_order):
            teams[team].append(row["player_name"])
    return teams


def teams_to_frames(teams: Dict[int, List[str]], master_players: pd.DataFrame) -> List[pd.DataFrame]:
    return [
        master_players[master_players["player_name"].isin(players)].copy()
        for _, players in sorted(teams.items())
    ]


def optimize_balanced_teams(master_players: pd.DataFrame, iterations: int = 5000) -> Tuple[Dict[int, List[str]], pd.DataFrame]:
    rng = random.Random(SEED)
    teams = initial_balanced_assignment(master_players)
    frames = teams_to_frames(teams, master_players)
    best_obj = draft_objective(frames)
    best_teams = {k: v[:] for k, v in teams.items()}

    for _ in range(iterations):
        t1, t2 = rng.sample(list(teams.keys()), 2)
        if not teams[t1] or not teams[t2]:
            continue
        p1 = rng.choice(teams[t1])
        p2 = rng.choice(teams[t2])
        teams[t1].remove(p1)
        teams[t2].remove(p2)
        teams[t1].append(p2)
        teams[t2].append(p1)
        candidate_frames = teams_to_frames(teams, master_players)
        candidate_obj = draft_objective(candidate_frames)
        if candidate_obj <= best_obj:
            best_obj = candidate_obj
            best_teams = {k: v[:] for k, v in teams.items()}
        else:
            teams[t1].remove(p2)
            teams[t2].remove(p1)
            teams[t1].append(p1)
            teams[t2].append(p2)

    rows = []
    for team_num, players in sorted(best_teams.items()):
        team_df = master_players[master_players["player_name"].isin(players)].copy()
        summary = team_strength(team_df)
        summary["team"] = team_num
        summary["players"] = ", ".join(team_df.sort_values("overall_rank")["player_name"])
        summary["avg_overall_rank"] = float(team_df["overall_rank"].mean())
        summary["top_pitcher"] = (
            team_df.sort_values("pitcher_rank")["player_name"].iloc[0]
            if team_df["pitcher_rank"].notna().any()
            else ""
        )
        rows.append(summary)
    summary_df = pd.DataFrame(rows).sort_values("team").reset_index(drop=True)
    return best_teams, summary_df


def team_roster_rows(teams: Dict[int, List[str]], master_players: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for team_num, players in sorted(teams.items()):
        team_df = master_players[master_players["player_name"].isin(players)].sort_values("overall_rank")
        for _, row in team_df.iterrows():
            rows.append(
                {
                    "team": team_num,
                    "player_name": row["player_name"],
                    "overall_rank": row["overall_rank"],
                    "pitcher_rank": row["pitcher_rank"],
                    "tier": row["tier"],
                    "overall_composite": row["overall_composite"],
                    "pitching_score": row["pitching_score"],
                    "balance_score": row["balance_score"],
                    "top_3_strengths": row["top_3_strengths"],
                    "risk_flag": row["risk_flag"],
                }
            )
    return pd.DataFrame(rows)


def add_header(fig, title: str, subtitle: str = "") -> None:
    fig.text(0.05, 0.955, title, fontsize=24, fontweight="bold", color=INK)
    if subtitle:
        fig.text(0.05, 0.92, subtitle, fontsize=10.5, color=MUTED)


def add_metric_card(fig, x: float, y: float, w: float, h: float, title: str, value: str, accent: str) -> None:
    ax = fig.add_axes([x, y, w, h])
    ax.axis("off")
    patch = FancyBboxPatch(
        (0, 0),
        1,
        1,
        boxstyle="round,pad=0.02,rounding_size=0.08",
        facecolor=PANEL,
        edgecolor=accent,
        linewidth=2.0,
    )
    ax.add_patch(patch)
    ax.text(0.06, 0.72, title, fontsize=10, color=MUTED, transform=ax.transAxes)
    ax.text(0.06, 0.28, value, fontsize=22, fontweight="bold", color=INK, transform=ax.transAxes)


def cover_page(pdf: PdfPages, inspection_summary: str, master_players: pd.DataFrame, team_summary: pd.DataFrame) -> None:
    fig = plt.figure(figsize=(11, 8.5))
    fig.patch.set_facecolor(BG)
    add_header(fig, "VCB 13U Draft Intelligence Report", "Corrected player model, draft board, and a suggested 7-team balanced build")

    add_metric_card(fig, 0.05, 0.74, 0.2, 0.12, "Players", str(len(master_players)), ACCENT)
    add_metric_card(fig, 0.27, 0.74, 0.2, 0.12, "Pitchers Evaluated", str(int(master_players["has_pitching_data"].sum())), ACCENT_2)
    add_metric_card(fig, 0.49, 0.74, 0.2, 0.12, "Suggested Teams", str(NUM_TEAMS), ACCENT_3)
    spread = team_summary["team_strength_score"].max() - team_summary["team_strength_score"].min()
    add_metric_card(fig, 0.71, 0.74, 0.24, 0.12, "Balanced Team Spread", f"{spread:.1f}", ACCENT)

    ax_left = fig.add_axes([0.05, 0.12, 0.43, 0.54])
    ax_left.axis("off")
    patch = FancyBboxPatch((0, 0), 1, 1, boxstyle="round,pad=0.02,rounding_size=0.04", facecolor=PANEL, edgecolor=GRID)
    ax_left.add_patch(patch)
    ax_left.text(0.04, 0.93, "What Changed", fontsize=15, fontweight="bold")
    bullets = [
        "Workbook parsing was corrected after confirming that higher 1-to-5 and 1-to-3 scores mean better grades.",
        "The board blends raw athletic/hitting data with coach-entered ranking sheets while keeping pitcher value separate.",
        "The proposed 7-team build optimizes team strength spread, pitching distribution, and top-end talent balance.",
        "Use this report as a draft-room guide, not as a substitute for live coach judgment or positional needs.",
    ]
    y = 0.82
    for bullet in bullets:
        ax_left.text(0.06, y, f"• {bullet}", fontsize=11, va="top", wrap=True)
        y -= 0.16

    ax_right = fig.add_axes([0.53, 0.12, 0.42, 0.54])
    ax_right.axis("off")
    patch = FancyBboxPatch((0, 0), 1, 1, boxstyle="round,pad=0.02,rounding_size=0.04", facecolor=PANEL, edgecolor=GRID)
    ax_right.add_patch(patch)
    ax_right.text(0.04, 0.93, "Workbook Reality Check", fontsize=15, fontweight="bold")
    summary_lines = inspection_summary.splitlines()
    y = 0.84
    for line in summary_lines[:6]:
        ax_right.text(0.05, y, line, fontsize=10.5, va="top", wrap=True)
        y -= 0.12

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def overview_page(pdf: PdfPages, master_players: pd.DataFrame, rankings: Dict[str, pd.DataFrame], hidden: pd.DataFrame) -> None:
    fig = plt.figure(figsize=(11, 8.5))
    add_header(fig, "Player Pool Overview", "Top-end talent, pitching depth, and value pockets")

    ax1 = fig.add_axes([0.06, 0.55, 0.4, 0.30])
    top10 = rankings["overall"].head(10).sort_values("overall_composite")
    ax1.barh(top10["player_name"], top10["overall_composite"], color=ACCENT)
    ax1.set_title("Top 10 Overall")
    ax1.grid(axis="x", color=GRID, alpha=0.6)

    ax2 = fig.add_axes([0.54, 0.55, 0.4, 0.30])
    top_pitch = rankings["pitchers"].head(10).sort_values("pitching_score")
    ax2.barh(top_pitch["player_name"], top_pitch["pitching_score"], color=ACCENT_2)
    ax2.set_title("Top 10 Pitchers")
    ax2.grid(axis="x", color=GRID, alpha=0.6)

    ax3 = fig.add_axes([0.06, 0.10, 0.4, 0.30])
    tier_counts = master_players["tier"].value_counts().sort_index()
    ax3.bar(tier_counts.index, tier_counts.values, color=[ACCENT, ACCENT_3, ACCENT_2, MUTED])
    ax3.set_title("Tier Distribution")
    ax3.tick_params(axis="x", rotation=20)
    ax3.grid(axis="y", color=GRID, alpha=0.6)

    ax4 = fig.add_axes([0.54, 0.10, 0.4, 0.30])
    ax4.axis("off")
    patch = FancyBboxPatch((0, 0), 1, 1, boxstyle="round,pad=0.02,rounding_size=0.04", facecolor=PANEL, edgecolor=GRID)
    ax4.add_patch(patch)
    ax4.text(0.04, 0.92, "Hidden Value / Role Fits", fontsize=14, fontweight="bold")
    y = 0.82
    for _, row in hidden.head(8).iterrows():
        ax4.text(
            0.05,
            y,
            f"{row['player_name']} | tier {row['tier'].replace('Tier ', '')} | strengths: {row['top_3_strengths']}",
            fontsize=10.2,
            va="top",
            wrap=True,
        )
        y -= 0.10

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def distributions_page(pdf: PdfPages, master_players: pd.DataFrame) -> None:
    fig = plt.figure(figsize=(11, 8.5))
    add_header(fig, "Score Shapes And Player Archetypes", "Where the separation is strongest")

    ax1 = fig.add_axes([0.06, 0.55, 0.4, 0.28])
    ax1.hist(master_players["overall_composite"], bins=14, color=ACCENT, alpha=0.85)
    ax1.set_title("Overall Composite Distribution")
    ax1.grid(axis="y", color=GRID, alpha=0.5)

    ax2 = fig.add_axes([0.54, 0.55, 0.4, 0.28])
    ax2.hist(master_players["pitching_score"].dropna(), bins=12, color=ACCENT_2, alpha=0.85)
    ax2.set_title("Pitching Score Distribution")
    ax2.grid(axis="y", color=GRID, alpha=0.5)

    ax3 = fig.add_axes([0.06, 0.10, 0.4, 0.30])
    scatter = ax3.scatter(
        master_players["position_player_score"],
        master_players["pitching_score"].fillna(0),
        c=master_players["overall_rank"],
        cmap="viridis_r",
        s=60,
        alpha=0.8,
        edgecolors="white",
        linewidths=0.5,
    )
    ax3.set_title("Position Value vs Pitching Value")
    ax3.set_xlabel("Position Player Score")
    ax3.set_ylabel("Pitching Score")
    fig.colorbar(scatter, ax=ax3, label="Overall Rank")

    ax4 = fig.add_axes([0.54, 0.10, 0.4, 0.30])
    corr_cols = [
        "athleticism_score",
        "hitting_objective_score",
        "fielding_score",
        "throwing_score",
        "pitching_score",
        "overall_composite",
    ]
    corr = master_players[corr_cols].corr()
    im = ax4.imshow(corr, cmap="coolwarm", vmin=-1, vmax=1)
    ax4.set_xticks(range(len(corr_cols)))
    ax4.set_xticklabels(corr_cols, rotation=40, ha="right")
    ax4.set_yticks(range(len(corr_cols)))
    ax4.set_yticklabels(corr_cols)
    ax4.set_title("Category Correlation")
    fig.colorbar(im, ax=ax4)

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def player_spotlights_page(pdf: PdfPages, master_players: pd.DataFrame, hidden: pd.DataFrame) -> None:
    fig = plt.figure(figsize=(11, 8.5))
    add_header(fig, "Detailed Player Analysis", "Leaders, balanced profiles, and specialist upside")

    leaders = [
        ("Top Overall", master_players.nsmallest(6, "overall_rank")[["player_name", "overall_rank", "top_3_strengths"]]),
        ("Top Pitchers", master_players[master_players["has_pitching_data"]].nsmallest(6, "pitcher_rank")[["player_name", "pitcher_rank", "top_3_strengths"]]),
        ("Most Balanced", master_players.sort_values(["balance_score", "overall_rank"], ascending=[False, True]).head(6)[["player_name", "balance_score", "top_3_strengths"]]),
        ("Value Targets", hidden.head(6)[["player_name", "value_delta", "top_3_strengths"]]),
    ]
    positions = [(0.05, 0.53), (0.53, 0.53), (0.05, 0.10), (0.53, 0.10)]
    for (title, table), (x, y) in zip(leaders, positions):
        ax = fig.add_axes([x, y, 0.42, 0.32])
        ax.axis("off")
        patch = FancyBboxPatch((0, 0), 1, 1, boxstyle="round,pad=0.02,rounding_size=0.04", facecolor=PANEL, edgecolor=GRID)
        ax.add_patch(patch)
        ax.text(0.04, 0.92, title, fontsize=14, fontweight="bold")
        yy = 0.80
        for _, row in table.iterrows():
            pieces = [str(v) for v in row.tolist()]
            ax.text(0.05, yy, " | ".join(pieces), fontsize=10.2, va="top", wrap=True)
            yy -= 0.12

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def teams_summary_page(pdf: PdfPages, team_summary: pd.DataFrame) -> None:
    fig = plt.figure(figsize=(11, 8.5))
    add_header(fig, "Suggested 7-Team Balanced Build", "All 77 players allocated into seven 11-player rosters")

    ax1 = fig.add_axes([0.06, 0.53, 0.42, 0.30])
    ax1.bar(team_summary["team"].astype(str), team_summary["team_strength_score"], color=ACCENT)
    ax1.set_title("Overall Team Strength")
    ax1.grid(axis="y", color=GRID, alpha=0.6)

    ax2 = fig.add_axes([0.54, 0.53, 0.40, 0.30])
    ax2.bar(team_summary["team"].astype(str), team_summary["pitching_strength"], color=ACCENT_2)
    ax2.set_title("Pitching Strength")
    ax2.grid(axis="y", color=GRID, alpha=0.6)

    ax3 = fig.add_axes([0.06, 0.10, 0.88, 0.28])
    ax3.axis("off")
    patch = FancyBboxPatch((0, 0), 1, 1, boxstyle="round,pad=0.02,rounding_size=0.04", facecolor=PANEL, edgecolor=GRID)
    ax3.add_patch(patch)
    ax3.text(0.02, 0.90, "Recommended Team Snapshot", fontsize=14, fontweight="bold")
    yy = 0.78
    for _, row in team_summary.iterrows():
        ax3.text(
            0.03,
            yy,
            f"Team {int(row['team'])}: strength {row['team_strength_score']:.1f}, pitching {row['pitching_strength']:.1f}, avg rank {row['avg_overall_rank']:.1f}, top arm {row['top_pitcher']}",
            fontsize=10.4,
            va="top",
            wrap=True,
        )
        yy -= 0.11

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def team_roster_page(pdf: PdfPages, team_num: int, team_df: pd.DataFrame) -> None:
    fig = plt.figure(figsize=(11, 8.5))
    summary = team_strength(team_df)
    add_header(
        fig,
        f"Team {team_num} Roster",
        f"Strength {summary['team_strength_score']:.1f} | Pitching {summary['pitching_strength']:.1f} | Avg balance {summary['average_balance']:.1f}",
    )

    ax = fig.add_axes([0.05, 0.08, 0.9, 0.80])
    ax.axis("off")
    patch = FancyBboxPatch((0, 0), 1, 1, boxstyle="round,pad=0.02,rounding_size=0.03", facecolor=PANEL, edgecolor=GRID)
    ax.add_patch(patch)
    ax.text(0.02, 0.96, "Player", fontweight="bold", fontsize=11)
    ax.text(0.28, 0.96, "Overall", fontweight="bold", fontsize=11)
    ax.text(0.37, 0.96, "Pitch", fontweight="bold", fontsize=11)
    ax.text(0.45, 0.96, "Tier", fontweight="bold", fontsize=11)
    ax.text(0.63, 0.96, "Top Strengths", fontweight="bold", fontsize=11)
    ax.text(0.87, 0.96, "Flag", fontweight="bold", fontsize=11)

    yy = 0.90
    for _, row in team_df.sort_values("overall_rank").iterrows():
        ax.text(0.02, yy, row["player_name"], fontsize=10)
        ax.text(0.29, yy, f"{int(row['overall_rank'])}", fontsize=10)
        ax.text(0.38, yy, "-" if pd.isna(row["pitcher_rank"]) else f"{int(row['pitcher_rank'])}", fontsize=10)
        ax.text(0.45, yy, row["tier"].replace("Tier ", "T"), fontsize=10)
        ax.text(0.63, yy, row["top_3_strengths"][:32], fontsize=10)
        ax.text(0.87, yy, row["risk_flag"][:18], fontsize=10)
        yy -= 0.073

    pdf.savefig(fig, bbox_inches="tight")
    plt.close(fig)


def generate_report() -> Dict[str, object]:
    style_matplotlib()
    inspection, inspection_summary = inspect_workbook(WORKBOOK_PATH)
    master_players = build_master_players(WORKBOOK_PATH)
    rankings = build_rankings(master_players)
    hidden = build_hidden_value_table(master_players)
    draft_board = build_draft_board(master_players)
    teams, team_summary = optimize_balanced_teams(master_players)
    rosters = team_roster_rows(teams, master_players)
    rosters.to_csv(OUTPUT_TEAMS_CSV, index=False)

    with PdfPages(OUTPUT_PDF) as pdf:
        cover_page(pdf, inspection_summary, master_players, team_summary)
        overview_page(pdf, master_players, rankings, hidden)
        distributions_page(pdf, master_players)
        player_spotlights_page(pdf, master_players, hidden)
        teams_summary_page(pdf, team_summary)
        for team_num in range(1, NUM_TEAMS + 1):
            team_df = master_players[master_players["player_name"].isin(teams[team_num])].copy()
            team_roster_page(pdf, team_num, team_df)

    return {
        "pdf": OUTPUT_PDF,
        "teams_csv": OUTPUT_TEAMS_CSV,
        "team_summary": team_summary,
        "teams": teams,
        "draft_board": draft_board,
    }


def main() -> None:
    generate_report()


if __name__ == "__main__":
    main()
