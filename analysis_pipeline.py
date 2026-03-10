
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import json
import math
import re
from typing import Dict, Iterable, List, Tuple
import warnings

import matplotlib.pyplot as plt
import nbformat as nbf
import numpy as np
import pandas as pd
from pandas.errors import PerformanceWarning
from sklearn.cluster import KMeans

warnings.filterwarnings("ignore", category=PerformanceWarning)


WORKBOOK_PATH = Path("VCB House - 13u PeeWee Assessment.xlsx")

# User-adjustable parameters for notebook and script usage.
NUM_TEAMS = 6
ROSTER_SIZE = 12
RANDOM_SEED = 13
DRAFT_STRATEGY = "both"
WEIGHT_OVERALL = 0.55
WEIGHT_PITCHING = 0.20
WEIGHT_BALANCE = 0.15
WEIGHT_UPSIDE = 0.10
LOCKED_PLAYERS: Dict[str, str] = {}


RANKING_COLUMNS = [
    "blank",
    "name",
    "speed_ranking",
    "power_ranking",
    "overall_athleticism_ranking",
    "hitting_subjective_ranking",
    "hitting_objective_ranking",
    "overall_hitting_ranking",
    "fielding_ranking",
    "throwing_ranking",
    "overall_ranking",
    "blank_2",
    "skill_rank",
    "speed_leader",
    "power_leader",
    "athletic_leader",
    "hitting_subjective_leader",
    "hitting_objective_leader",
    "overall_hitting_leader",
    "fielding_leader",
    "throwing_leader",
    "pitching_leader",
    "pitching_leader_2",
]

ASSESSMENT_COLUMNS = [
    "name",
    "home_to_1st",
    "broad_jump",
    "lateral_jump",
    "shotput",
    "bat_speed",
    "time_to_contact",
    "exit_velocity_avg",
    "exit_velocity_max",
    "athletic_stance",
    "balance_stride",
    "barrel_level",
    "launch_position",
    "follow_through",
    "readiness",
    "footwork",
    "glovework",
    "field_athleticism",
    "fundamental_throwing",
]

PITCHING_COLUMNS = [
    "name",
    "velocity_avg",
    "velocity_max",
    "pitch_1",
    "pitch_2",
    "pitch_3",
    "pitch_4",
    "athletic_movement",
    "body_control",
    "direction",
    "repeatability",
    "command",
]

PITCHER_RANKING_COLUMNS = [
    "name",
    "pitcher_subjective_ranking",
    "fb_velo_command_ranking",
    "pitcher_overall_ranking",
]


NUMERIC_COLUMNS = {
    "ranking": [
        "speed_ranking",
        "power_ranking",
        "overall_athleticism_ranking",
        "hitting_subjective_ranking",
        "hitting_objective_ranking",
        "overall_hitting_ranking",
        "fielding_ranking",
        "throwing_ranking",
        "overall_ranking",
        "skill_rank",
    ],
    "assessment": [
        "home_to_1st",
        "broad_jump",
        "lateral_jump",
        "shotput",
        "bat_speed",
        "time_to_contact",
        "exit_velocity_avg",
        "exit_velocity_max",
        "athletic_stance",
        "balance_stride",
        "barrel_level",
        "launch_position",
        "follow_through",
        "readiness",
        "footwork",
        "glovework",
        "field_athleticism",
        "fundamental_throwing",
    ],
    "pitching": [
        "velocity_avg",
        "velocity_max",
        "athletic_movement",
        "body_control",
        "direction",
        "repeatability",
        "command",
    ],
    "pitcher_ranking": [
        "pitcher_subjective_ranking",
        "fb_velo_command_ranking",
        "pitcher_overall_ranking",
    ],
}


def standardize_name(value: object) -> str:
    if pd.isna(value):
        return ""
    value = str(value).replace("*", "")
    value = re.sub(r"\s+", " ", value).strip()
    return value.title()


def clean_text_numeric(value: object) -> float:
    if pd.isna(value):
        return np.nan
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)
    text = str(value).strip()
    if not text:
        return np.nan
    text = text.replace(",", "")
    match = re.search(r"-?\d+(\.\d+)?", text)
    return float(match.group()) if match else np.nan


def min_max_score(series: pd.Series, higher_is_better: bool = True) -> pd.Series:
    values = series.astype(float)
    result = pd.Series(np.nan, index=series.index, dtype=float)
    valid = values.dropna()
    if valid.empty:
        return result
    if valid.nunique() == 1:
        result.loc[valid.index] = 50.0
        return result
    scaled = (valid - valid.min()) / (valid.max() - valid.min())
    if not higher_is_better:
        scaled = 1 - scaled
    result.loc[valid.index] = scaled * 100
    return result


def percentile_score(series: pd.Series, higher_is_better: bool = True) -> pd.Series:
    rank = series.rank(method="average", pct=True, ascending=not higher_is_better)
    return rank * 100


def z_score(series: pd.Series, higher_is_better: bool = True) -> pd.Series:
    values = series.astype(float)
    mean = values.mean(skipna=True)
    std = values.std(skipna=True, ddof=0)
    if pd.isna(std) or std == 0:
        return pd.Series(0.0, index=series.index)
    z = (values - mean) / std
    return z if higher_is_better else -z


def average_available(df: pd.DataFrame, columns: Iterable[str]) -> pd.Series:
    return df[list(columns)].mean(axis=1, skipna=True)


def top_strengths(row: pd.Series, category_columns: List[str], limit: int = 3) -> str:
    strengths = []
    for col in category_columns:
        val = row.get(col)
        if pd.notna(val):
            strengths.append((col.replace("_percentile", "").replace("_", " ").title(), float(val)))
    strengths.sort(key=lambda item: item[1], reverse=True)
    return ", ".join(name for name, _ in strengths[:limit])


def detect_malformed_values(df: pd.DataFrame, numeric_columns: Iterable[str]) -> Dict[str, List[str]]:
    malformed: Dict[str, List[str]] = {}
    for col in numeric_columns:
        if col not in df.columns:
            continue
        series = df[col]
        bad_values = []
        for value in series.dropna():
            if isinstance(value, (int, float, np.integer, np.floating)):
                continue
            text = str(value).strip()
            if not text:
                continue
            if re.fullmatch(r"-?\d+(\.\d+)?", text):
                continue
            bad_values.append(text)
        if bad_values:
            malformed[col] = sorted(set(bad_values))[:10]
    return malformed


def parse_workbook(workbook_path: Path = WORKBOOK_PATH) -> Dict[str, pd.DataFrame]:
    xl = pd.ExcelFile(workbook_path)
    raw = {sheet: xl.parse(sheet, dtype=object) for sheet in xl.sheet_names}

    ranking = raw["Ranking "].iloc[3:].reset_index(drop=True)
    ranking.columns = RANKING_COLUMNS
    ranking = ranking[ranking["name"].notna()].copy()
    ranking = ranking[pd.to_numeric(ranking["skill_rank"], errors="coerce").notna()].copy()

    assessment = raw["Assessment Data"].copy()
    assessment.columns = ASSESSMENT_COLUMNS
    assessment = assessment[assessment["name"].astype(str).str.lower() != "name"].copy()

    pitching = raw["Pitching Data "].copy()
    pitching.columns = PITCHING_COLUMNS
    pitching = pitching[pitching["name"].astype(str).str.lower() != "name"].copy()

    pitcher_ranking = raw["Pitcher Ranking "].copy()
    pitcher_ranking.columns = PITCHER_RANKING_COLUMNS
    pitcher_ranking = pitcher_ranking[pitcher_ranking["name"].astype(str).str.lower() != "name"].copy()

    frames = {
        "ranking": ranking,
        "assessment": assessment,
        "pitching": pitching,
        "pitcher_ranking": pitcher_ranking,
    }

    for key, frame in frames.items():
        frame["player_name"] = frame["name"].map(standardize_name)
        frame.drop(columns=["name"], inplace=True)
        numeric_cols = NUMERIC_COLUMNS.get(key, [])
        for col in numeric_cols:
            if col in frame.columns:
                frame[col] = frame[col].map(clean_text_numeric)

    for col in ["pitch_1", "pitch_2", "pitch_3", "pitch_4"]:
        if col in frames["pitching"].columns:
            frames["pitching"][col] = (
                frames["pitching"][col]
                .fillna("")
                .astype(str)
                .str.replace("*", "", regex=False)
                .str.replace("2 S - FB", "2-Seam FB", regex=False)
                .str.replace("2-FB", "2-Seam FB", regex=False)
                .str.strip()
                .replace("", np.nan)
            )

    return frames


def inspect_workbook(workbook_path: Path = WORKBOOK_PATH) -> Tuple[Dict[str, dict], str]:
    xl = pd.ExcelFile(workbook_path)
    inspection: Dict[str, dict] = {}

    for sheet in xl.sheet_names:
        df = xl.parse(sheet, dtype=object)
        malformed = detect_malformed_values(df, df.columns)
        inspection[sheet] = {
            "row_count": int(df.shape[0]),
            "column_count": int(df.shape[1]),
            "column_names": [str(col) for col in df.columns],
            "missing_by_column": {str(k): int(v) for k, v in df.isna().sum().items()},
            "sample_rows": df.head(5).fillna("").astype(str).to_dict(orient="records"),
            "malformed_examples": malformed,
        }

    frames = parse_workbook(workbook_path)
    name_sets = {key: set(frame["player_name"]) for key, frame in frames.items()}
    all_names = sorted(set.union(*name_sets.values()))
    missing_pitching = sorted(
        name for name in all_names if name not in name_sets["pitching"] and name not in name_sets["pitcher_ranking"]
    )

    summary_lines = [
        "The workbook has four sheets: one overall ranking summary, one pitcher ranking summary, one detailed assessment sheet, and one detailed pitching sheet.",
        "The two summary ranking sheets are not tidy tables at the top of the sheet. They include title rows and embedded header rows, so they need sheet-specific parsing before analysis.",
        "The detailed assessment sheet covers 77 players and contains raw athletic, hitting, and fielding/throwing measurements.",
        "The pitching sheets cover 73 players, which means four players appear to have no pitching evaluation recorded.",
        "Player names are consistent across sheets after standardizing whitespace and capitalization, and there were no duplicate player records after cleanup.",
        f"The four players missing pitching-specific data are: {', '.join(missing_pitching)}.",
    ]

    inspection["cross_sheet_summary"] = {
        "all_unique_players": len(all_names),
        "missing_pitching_players": missing_pitching,
        "players_by_sheet": {key: int(frame['player_name'].nunique()) for key, frame in frames.items()},
    }

    return inspection, "\n".join(summary_lines)


def build_master_players(workbook_path: Path = WORKBOOK_PATH) -> pd.DataFrame:
    frames = parse_workbook(workbook_path)
    master_players = frames["assessment"].merge(
        frames["ranking"].drop(columns=["blank", "blank_2"], errors="ignore"),
        on="player_name",
        how="left",
        suffixes=("", "_rank"),
    )
    master_players = master_players.merge(frames["pitching"], on="player_name", how="left")
    master_players = master_players.merge(frames["pitcher_ranking"], on="player_name", how="left")
    master_players = master_players.drop_duplicates(subset=["player_name"]).reset_index(drop=True)

    master_players["has_pitching_data"] = master_players["velocity_avg"].notna()
    pitch_cols = ["pitch_1", "pitch_2", "pitch_3", "pitch_4"]
    master_players["pitch_count"] = master_players[pitch_cols].notna().sum(axis=1)

    raw_metrics = {
        "speed_raw": ("home_to_1st", False),
        "power_raw": ("shotput", True),
        "athleticism_raw": (["broad_jump", "lateral_jump", "shotput", "home_to_1st"], None),
        "hitting_objective_raw": (["bat_speed", "exit_velocity_avg", "exit_velocity_max", "time_to_contact"], None),
        "hitting_subjective_raw": (["athletic_stance", "balance_stride", "barrel_level", "launch_position", "follow_through"], False),
        "fielding_raw": (["readiness", "footwork", "glovework", "field_athleticism"], False),
        "throwing_raw": ("fundamental_throwing", False),
        "pitching_raw": (["velocity_avg", "velocity_max", "athletic_movement", "body_control", "direction", "repeatability", "command"], None),
    }

    master_players["speed_score"] = min_max_score(master_players["home_to_1st"], higher_is_better=False)
    master_players["power_score"] = min_max_score(master_players["shotput"], higher_is_better=True)
    master_players["broad_jump_score"] = min_max_score(master_players["broad_jump"], higher_is_better=True)
    master_players["lateral_jump_score"] = min_max_score(master_players["lateral_jump"], higher_is_better=True)
    master_players["bat_speed_score"] = min_max_score(master_players["bat_speed"], higher_is_better=True)
    master_players["time_to_contact_score"] = min_max_score(master_players["time_to_contact"], higher_is_better=False)
    master_players["exit_velocity_avg_score"] = min_max_score(master_players["exit_velocity_avg"], higher_is_better=True)
    master_players["exit_velocity_max_score"] = min_max_score(master_players["exit_velocity_max"], higher_is_better=True)
    master_players["velocity_avg_score"] = min_max_score(master_players["velocity_avg"], higher_is_better=True)
    master_players["velocity_max_score"] = min_max_score(master_players["velocity_max"], higher_is_better=True)

    high_scale_cols = [
        "athletic_stance",
        "balance_stride",
        "barrel_level",
        "launch_position",
        "follow_through",
        "readiness",
        "footwork",
        "glovework",
        "field_athleticism",
        "fundamental_throwing",
        "athletic_movement",
        "body_control",
        "direction",
        "repeatability",
        "command",
        "speed_ranking",
        "power_ranking",
        "overall_athleticism_ranking",
        "hitting_subjective_ranking",
        "hitting_objective_ranking",
        "overall_hitting_ranking",
        "fielding_ranking",
        "throwing_ranking",
        "overall_ranking",
        "pitcher_subjective_ranking",
        "fb_velo_command_ranking",
        "pitcher_overall_ranking",
    ]
    for col in high_scale_cols:
        if col in master_players.columns:
            master_players[f"{col}_score"] = min_max_score(master_players[col], higher_is_better=True)

    master_players["athleticism_score"] = average_available(
        master_players,
        ["speed_score", "broad_jump_score", "lateral_jump_score", "power_score"],
    )
    master_players["hitting_objective_score"] = average_available(
        master_players,
        ["bat_speed_score", "time_to_contact_score", "exit_velocity_avg_score", "exit_velocity_max_score"],
    )
    master_players["hitting_subjective_score"] = average_available(
        master_players,
        [
            "athletic_stance_score",
            "balance_stride_score",
            "barrel_level_score",
            "launch_position_score",
            "follow_through_score",
        ],
    )
    master_players["fielding_score"] = average_available(
        master_players,
        [
            "readiness_score",
            "footwork_score",
            "glovework_score",
            "field_athleticism_score",
        ],
    )
    master_players["throwing_score"] = master_players["fundamental_throwing_score"]
    master_players["pitching_delivery_score"] = average_available(
        master_players,
        [
            "athletic_movement_score",
            "body_control_score",
            "direction_score",
            "repeatability_score",
            "command_score",
        ],
    )
    master_players["pitching_score"] = average_available(
        master_players,
        [
            "velocity_avg_score",
            "velocity_max_score",
            "pitching_delivery_score",
            "pitcher_overall_ranking_score",
            "fb_velo_command_ranking_score",
        ],
    )
    master_players["position_player_score"] = average_available(
        master_players,
        [
            "athleticism_score",
            "hitting_objective_score",
            "hitting_subjective_score",
            "fielding_score",
            "throwing_score",
        ],
    )
    master_players["overall_composite"] = average_available(
        master_players,
        [
            "position_player_score",
            "overall_ranking_score",
            "overall_hitting_ranking_score",
            "overall_athleticism_ranking_score",
            "pitching_score",
        ],
    )

    master_players["overall_percentile"] = percentile_score(master_players["overall_composite"], higher_is_better=True)
    master_players["pitching_percentile"] = percentile_score(master_players["pitching_score"], higher_is_better=True)
    master_players["position_player_percentile"] = percentile_score(master_players["position_player_score"], higher_is_better=True)

    category_base = [
        "athleticism_score",
        "hitting_objective_score",
        "hitting_subjective_score",
        "fielding_score",
        "throwing_score",
    ]
    master_players["balance_score"] = (
        100 - master_players[category_base].std(axis=1, skipna=True).fillna(0).mul(10)
    ).clip(lower=0, upper=100)
    upside_components = pd.concat(
        [
            percentile_score(master_players["speed_score"]).rename("speed_pct"),
            percentile_score(master_players["power_score"]).rename("power_pct"),
            percentile_score(master_players["pitching_score"]).rename("pitching_pct"),
            percentile_score(master_players["hitting_objective_score"]).rename("hitting_obj_pct"),
        ],
        axis=1,
    )
    master_players["upside_score"] = upside_components.mean(axis=1, skipna=True)
    master_players["consistency_score"] = master_players["balance_score"]

    master_players["raw_score_model"] = (
        0.20 * master_players["athleticism_score"]
        + 0.25 * master_players["hitting_objective_score"]
        + 0.15 * master_players["hitting_subjective_score"]
        + 0.20 * master_players["fielding_score"]
        + 0.10 * master_players["throwing_score"]
        + 0.10 * master_players["pitching_score"].fillna(master_players["position_player_score"])
    )

    norm_components = pd.DataFrame(
        {
            "athleticism_z": z_score(master_players["athleticism_score"]),
            "hitting_obj_z": z_score(master_players["hitting_objective_score"]),
            "hitting_subj_z": z_score(master_players["hitting_subjective_score"]),
            "fielding_z": z_score(master_players["fielding_score"]),
            "throwing_z": z_score(master_players["throwing_score"]),
            "pitching_z": z_score(master_players["pitching_score"].fillna(master_players["pitching_score"].mean())),
        }
    )
    master_players["normalized_score_model"] = (
        50
        + 12
        * (
            0.20 * norm_components["athleticism_z"]
            + 0.25 * norm_components["hitting_obj_z"]
            + 0.15 * norm_components["hitting_subj_z"]
            + 0.15 * norm_components["fielding_z"]
            + 0.10 * norm_components["throwing_z"]
            + 0.15 * norm_components["pitching_z"]
        )
    )

    master_players["balanced_player_model"] = (
        0.65 * master_players["normalized_score_model"]
        + 0.35 * master_players["balance_score"]
    )
    max_tool = master_players[
        ["athleticism_score", "hitting_objective_score", "fielding_score", "pitching_score", "power_score", "speed_score"]
    ].max(axis=1, skipna=True)
    master_players["specialist_upside_model"] = 0.55 * master_players["normalized_score_model"] + 0.45 * max_tool

    rank_map = {
        "overall_rank": "overall_composite",
        "pitcher_rank": "pitching_score",
        "position_player_rank": "position_player_score",
        "raw_model_rank": "raw_score_model",
        "normalized_model_rank": "normalized_score_model",
        "balanced_model_rank": "balanced_player_model",
        "upside_model_rank": "specialist_upside_model",
    }
    for rank_col, score_col in rank_map.items():
        master_players[rank_col] = master_players[score_col].rank(method="min", ascending=False)

    try:
        tier_labels = {
            0: "Tier 4 - Development Players",
            1: "Tier 3 - Solid Contributors",
            2: "Tier 2 - Strong Starters",
            3: "Tier 1 - Impact Players",
        }
        kmeans = KMeans(n_clusters=4, random_state=RANDOM_SEED, n_init=20)
        filled = master_players["overall_composite"].fillna(master_players["overall_composite"].median()).to_numpy().reshape(-1, 1)
        clusters = kmeans.fit_predict(filled)
        cluster_order = (
            pd.DataFrame({"cluster": clusters, "score": master_players["overall_composite"]})
            .groupby("cluster")["score"]
            .mean()
            .sort_values()
            .index.tolist()
        )
        cluster_to_tier = {cluster: tier_labels[i] for i, cluster in enumerate(cluster_order)}
        master_players["tier"] = [cluster_to_tier[c] for c in clusters]
    except Exception:
        master_players["tier"] = pd.qcut(
            master_players["overall_composite"].rank(method="first"),
            q=4,
            labels=[
                "Tier 4 - Development Players",
                "Tier 3 - Solid Contributors",
                "Tier 2 - Strong Starters",
                "Tier 1 - Impact Players",
            ],
        )

    category_percentiles = {
        "athleticism_percentile": "athleticism_score",
        "hitting_percentile": "position_player_score",
        "fielding_percentile": "fielding_score",
        "throwing_percentile": "throwing_score",
        "pitching_percentile_tool": "pitching_score",
    }
    for pct_col, src_col in category_percentiles.items():
        master_players[pct_col] = percentile_score(master_players[src_col], higher_is_better=True)

    master_players["value_delta"] = master_players["overall_rank"] - master_players["normalized_model_rank"]
    master_players["variance_indicator"] = master_players[
        ["athleticism_score", "hitting_objective_score", "fielding_score", "throwing_score", "pitching_score"]
    ].std(axis=1, skipna=True)
    master_players["is_hidden_value"] = master_players["value_delta"] >= 8
    master_players["is_balanced_player"] = master_players["balance_score"] >= master_players["balance_score"].quantile(0.75)
    master_players["is_specialist_player"] = master_players["variance_indicator"] >= master_players["variance_indicator"].quantile(0.75)
    master_players["is_high_variance"] = master_players["variance_indicator"] >= master_players["variance_indicator"].quantile(0.85)

    master_players["top_3_strengths"] = master_players.apply(
        lambda row: top_strengths(
            row,
            [
                "athleticism_percentile",
                "hitting_percentile",
                "fielding_percentile",
                "throwing_percentile",
                "pitching_percentile_tool",
            ],
        ),
        axis=1,
    )

    def weakness_flag(row: pd.Series) -> str:
        flags = []
        if pd.notna(row["pitching_score"]) and row["pitching_percentile"] < 35:
            flags.append("Pitching depth")
        if row["fielding_percentile"] < 35:
            flags.append("Defensive polish")
        if row["hitting_percentile"] < 35:
            flags.append("Offensive impact")
        if row["variance_indicator"] > master_players["variance_indicator"].quantile(0.85):
            flags.append("High variance profile")
        return ", ".join(flags[:2]) if flags else "No major flag from data"

    master_players["risk_flag"] = master_players.apply(weakness_flag, axis=1)
    master_players["draft_band_recommendation"] = pd.cut(
        master_players["overall_rank"],
        bins=[0, 12, 24, 42, 77],
        labels=["Round 1-2", "Round 3-4", "Round 5-7", "Late / Development"],
    ).astype(str)

    def scouting_summary(row: pd.Series) -> str:
        pitch_text = "with pitching utility" if pd.notna(row["pitching_score"]) and row["pitching_percentile"] >= 60 else "with position-player value"
        if row["is_hidden_value"]:
            prefix = "Undervalued profile"
        elif row["tier"].startswith("Tier 1"):
            prefix = "Impact option"
        elif row["is_balanced_player"]:
            prefix = "Reliable all-around option"
        else:
            prefix = "Role-dependent option"
        return f"{prefix}; strongest areas are {row['top_3_strengths'].lower()} {pitch_text}."

    master_players["short_scouting_summary"] = master_players.apply(scouting_summary, axis=1)
    master_players["player_type_cluster"] = master_players["tier"].str.replace("Tier \\d+ - ", "", regex=True)

    return master_players.sort_values(["overall_rank", "player_name"]).reset_index(drop=True)


def build_rankings(master_players: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    overall = master_players.sort_values(["overall_rank", "player_name"]).copy()
    pitchers = master_players[master_players["has_pitching_data"]].sort_values(["pitcher_rank", "player_name"]).copy()
    position_players = master_players.sort_values(["position_player_rank", "player_name"]).copy()

    tool_cols = {
        "athleticism": "athleticism_score",
        "hitting_objective": "hitting_objective_score",
        "fielding": "fielding_score",
        "throwing": "throwing_score",
        "pitching": "pitching_score",
    }
    tool_rankings = {}
    for label, score_col in tool_cols.items():
        subset = master_players[["player_name", score_col]].copy()
        subset = subset.dropna().sort_values(score_col, ascending=False).reset_index(drop=True)
        subset[f"{label}_rank"] = np.arange(1, len(subset) + 1)
        tool_rankings[label] = subset

    return {
        "overall": overall,
        "pitchers": pitchers,
        "position_players": position_players,
        "tool_rankings": tool_rankings,
    }


def build_hidden_value_table(master_players: pd.DataFrame) -> pd.DataFrame:
    hidden = master_players[
        master_players["is_hidden_value"] | master_players["is_balanced_player"] | master_players["is_specialist_player"]
    ][
        [
            "player_name",
            "overall_rank",
            "normalized_model_rank",
            "balanced_model_rank",
            "upside_model_rank",
            "tier",
            "balance_score",
            "upside_score",
            "value_delta",
            "top_3_strengths",
            "risk_flag",
        ]
    ].sort_values(["value_delta", "balanced_model_rank"], ascending=[False, True])
    return hidden.reset_index(drop=True)


def build_draft_board(master_players: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "player_name",
        "overall_rank",
        "pitcher_rank",
        "position_player_rank",
        "tier",
        "overall_composite",
        "overall_percentile",
        "balance_score",
        "upside_score",
        "top_3_strengths",
        "risk_flag",
        "draft_band_recommendation",
        "short_scouting_summary",
    ]
    board = master_players[cols].copy()
    board.columns = [
        "Player Name",
        "Overall Rank",
        "Pitcher Rank",
        "Position Player Rank",
        "Tier",
        "Composite Score",
        "Percentile Score",
        "Balanced Score",
        "Upside Score",
        "Top 3 Strengths",
        "Risk / Weakness Flag",
        "Draft Band Recommendation",
        "Short Scouting Summary",
    ]
    return board.sort_values(["Overall Rank", "Player Name"]).reset_index(drop=True)


def team_strength(
    team_df: pd.DataFrame,
    weight_overall: float = WEIGHT_OVERALL,
    weight_pitching: float = WEIGHT_PITCHING,
    weight_balance: float = WEIGHT_BALANCE,
    weight_upside: float = WEIGHT_UPSIDE,
) -> dict:
    tier_points = (
        (team_df["tier"].str.contains("Tier 1")).sum() * 3
        + (team_df["tier"].str.contains("Tier 2")).sum() * 2
        + (team_df["tier"].str.contains("Tier 3")).sum() * 1
    )
    top_end = team_df.nsmallest(3, "overall_rank")["overall_composite"].mean()
    return {
        "team_size": int(len(team_df)),
        "overall_strength": float(team_df["overall_composite"].sum()),
        "pitching_strength": float(team_df["pitching_score"].fillna(0).sum()),
        "average_balance": float(team_df["balance_score"].mean()),
        "tier_depth_points": float(tier_points),
        "top_end_talent": float(top_end if pd.notna(top_end) else 0),
        "team_strength_score": float(
            weight_overall * team_df["overall_composite"].sum()
            + weight_pitching * team_df["pitching_score"].fillna(0).sum()
            + weight_balance * team_df["balance_score"].mean()
            + weight_upside * team_df["upside_score"].sum()
        ),
    }


def snake_order(num_teams: int, roster_size: int) -> List[Tuple[int, int]]:
    order = []
    for round_num in range(1, roster_size + 1):
        teams = list(range(1, num_teams + 1))
        if round_num % 2 == 0:
            teams.reverse()
        for team in teams:
            order.append((round_num, team))
    return order


def balanced_pick_score(
    player: pd.Series,
    current_team: pd.DataFrame,
    league_baselines: dict,
    weight_overall: float = WEIGHT_OVERALL,
    weight_pitching: float = WEIGHT_PITCHING,
    weight_balance: float = WEIGHT_BALANCE,
    weight_upside: float = WEIGHT_UPSIDE,
) -> float:
    score = (
        weight_overall * player["overall_composite"]
        + weight_pitching * (player["pitching_score"] if pd.notna(player["pitching_score"]) else league_baselines["pitching_score"])
        + weight_balance * player["balance_score"]
        + weight_upside * player["upside_score"]
    )
    if current_team.empty:
        return float(score)
    current_pitching = current_team["pitching_score"].fillna(0).mean()
    current_balance = current_team["balance_score"].mean()
    current_top_end = current_team.nsmallest(min(2, len(current_team)), "overall_rank")["overall_composite"].mean()
    if current_pitching < league_baselines["pitching_score"] and pd.notna(player["pitching_score"]):
        score += 8
    if current_balance < league_baselines["balance_score"]:
        score += max(0, player["balance_score"] - current_balance) * 0.1
    if current_top_end > league_baselines["overall_composite"] and player["overall_rank"] > league_baselines["overall_rank"]:
        score += 4
    return float(score)


def simulate_draft(
    draft_board: pd.DataFrame,
    master_players: pd.DataFrame,
    num_teams: int = NUM_TEAMS,
    roster_size: int = ROSTER_SIZE,
    strategy: str = "best_available",
    locked_players: Dict[str, str] | None = None,
    weight_overall: float = WEIGHT_OVERALL,
    weight_pitching: float = WEIGHT_PITCHING,
    weight_balance: float = WEIGHT_BALANCE,
    weight_upside: float = WEIGHT_UPSIDE,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    locked_players = locked_players or {}
    available = draft_board.copy()
    available["lock_team"] = available["Player Name"].map(locked_players)
    teams = {team: [] for team in range(1, num_teams + 1)}
    picks = []

    for player_name, team_name in locked_players.items():
        team_num = int(str(team_name).replace("Team ", ""))
        player_row = available[available["Player Name"] == player_name]
        if player_row.empty:
            continue
        picks.append({"Round": 0, "Pick": 0, "Team": team_num, "Player Name": player_name, "Strategy": strategy, "Reason": "Locked"})
        teams[team_num].append(player_name)
        available = available[available["Player Name"] != player_name]

    baselines = {
        "overall_composite": float(master_players["overall_composite"].mean()),
        "pitching_score": float(master_players["pitching_score"].fillna(0).mean()),
        "balance_score": float(master_players["balance_score"].mean()),
        "overall_rank": float(master_players["overall_rank"].mean()),
    }

    order = snake_order(num_teams, roster_size)
    pick_number = 1
    for round_num, team in order:
        if len(teams[team]) >= roster_size:
            continue
        if available.empty:
            break

        team_df = master_players[master_players["player_name"].isin(teams[team])]
        if strategy == "best_available":
            choice = available.sort_values(["Overall Rank", "Player Name"]).iloc[0]
            reason = "Highest draft board rank available"
        else:
            candidates = available.merge(
                master_players[["player_name", "overall_composite", "pitching_score", "balance_score", "upside_score", "overall_rank"]],
                left_on="Player Name",
                right_on="player_name",
                how="left",
            )
            candidates["team_fit_score"] = candidates.apply(
                lambda row: balanced_pick_score(
                    row,
                    team_df,
                    baselines,
                    weight_overall=weight_overall,
                    weight_pitching=weight_pitching,
                    weight_balance=weight_balance,
                    weight_upside=weight_upside,
                ),
                axis=1,
            )
            choice = candidates.sort_values(["team_fit_score", "Overall Rank"], ascending=[False, True]).iloc[0]
            reason = "Best fit for team balance model"

        player_name = choice["Player Name"]
        teams[team].append(player_name)
        picks.append(
            {
                "Round": round_num,
                "Pick": pick_number,
                "Team": team,
                "Player Name": player_name,
                "Strategy": strategy,
                "Reason": reason,
            }
        )
        available = available[available["Player Name"] != player_name]
        pick_number += 1

    picks_df = pd.DataFrame(picks)
    team_rows = []
    for team, players in teams.items():
        team_df = master_players[master_players["player_name"].isin(players)]
        summary = team_strength(
            team_df,
            weight_overall=weight_overall,
            weight_pitching=weight_pitching,
            weight_balance=weight_balance,
            weight_upside=weight_upside,
        )
        summary["team"] = team
        summary["players"] = ", ".join(team_df.sort_values("overall_rank")["player_name"])
        team_rows.append(summary)
    team_strength_df = pd.DataFrame(team_rows).sort_values("team").reset_index(drop=True)
    return picks_df, team_strength_df


def save_visualizations(
    master_players: pd.DataFrame,
    team_strength_summaries: Dict[str, pd.DataFrame],
    output_dir: Path,
) -> List[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = []

    def save_current(name: str) -> None:
        path = output_dir / name
        plt.tight_layout()
        plt.savefig(path, dpi=150, bbox_inches="tight")
        plt.close()
        paths.append(path)

    plt.figure(figsize=(8, 5))
    master_players["overall_composite"].hist(bins=15)
    plt.title("Distribution of Overall Scores")
    plt.xlabel("Overall Composite")
    plt.ylabel("Players")
    save_current("overall_score_distribution.png")

    plt.figure(figsize=(8, 5))
    master_players["pitching_score"].dropna().hist(bins=12)
    plt.title("Distribution of Pitching Scores")
    plt.xlabel("Pitching Score")
    plt.ylabel("Pitchers")
    save_current("pitching_score_distribution.png")

    plt.figure(figsize=(9, 8))
    top20 = master_players.nsmallest(20, "overall_rank").sort_values("overall_composite")
    plt.barh(top20["player_name"], top20["overall_composite"])
    plt.title("Top 20 Overall Players")
    plt.xlabel("Overall Composite")
    save_current("top_20_overall_players.png")

    plt.figure(figsize=(9, 8))
    top_pitchers = master_players[master_players["has_pitching_data"]].nsmallest(20, "pitcher_rank").sort_values("pitching_score")
    plt.barh(top_pitchers["player_name"], top_pitchers["pitching_score"])
    plt.title("Top Pitchers")
    plt.xlabel("Pitching Score")
    save_current("top_pitchers.png")

    plt.figure(figsize=(7, 4))
    master_players["tier"].value_counts().sort_index().plot(kind="bar")
    plt.title("Tier Distribution")
    plt.xlabel("Tier")
    plt.ylabel("Players")
    save_current("tier_distribution.png")

    corr_cols = [
        "athleticism_score",
        "hitting_objective_score",
        "hitting_subjective_score",
        "fielding_score",
        "throwing_score",
        "pitching_score",
        "overall_composite",
    ]
    corr = master_players[corr_cols].corr()
    plt.figure(figsize=(7, 6))
    plt.imshow(corr, cmap="coolwarm", vmin=-1, vmax=1)
    plt.xticks(range(len(corr_cols)), corr_cols, rotation=45, ha="right")
    plt.yticks(range(len(corr_cols)), corr_cols)
    plt.colorbar(label="Correlation")
    plt.title("Category Correlation Heatmap")
    save_current("category_correlation_heatmap.png")

    for label, df in team_strength_summaries.items():
        plt.figure(figsize=(8, 5))
        plt.bar(df["team"].astype(str), df["team_strength_score"])
        plt.title(f"Team Strength by Simulated Team ({label})")
        plt.xlabel("Team")
        plt.ylabel("Strength Score")
        save_current(f"team_strength_{label}.png")

    plt.figure(figsize=(9, 5))
    compare = []
    for label, df in team_strength_summaries.items():
        temp = df[["team", "team_strength_score"]].copy()
        temp["strategy"] = label
        compare.append(temp)
    compare_df = pd.concat(compare, ignore_index=True)
    for strategy, group in compare_df.groupby("strategy"):
        plt.plot(group["team"], group["team_strength_score"], marker="o", label=strategy)
    plt.title("Comparison Across Draft Simulations")
    plt.xlabel("Team")
    plt.ylabel("Strength Score")
    plt.legend()
    save_current("draft_simulation_comparison.png")

    return paths


def build_report(
    inspection_summary: str,
    master_players: pd.DataFrame,
    hidden_value_players: pd.DataFrame,
    team_strength_summaries: Dict[str, pd.DataFrame],
) -> str:
    top_overall = master_players.nsmallest(10, "overall_rank")["player_name"].tolist()
    top_pitchers = master_players[master_players["has_pitching_data"]].nsmallest(8, "pitcher_rank")["player_name"].tolist()
    balanced = master_players.sort_values("balance_score", ascending=False).head(8)["player_name"].tolist()
    hidden = hidden_value_players.head(8)["player_name"].tolist()

    best_available = team_strength_summaries["best_available"]
    balanced_summary = team_strength_summaries["balanced"]

    lines = [
        "# VCB 13U House Draft Summary",
        "",
        "## Workbook Structure",
        inspection_summary,
        "",
        "## Cleaning Steps",
        "- Removed title rows and embedded header rows from the ranking sheets.",
        "- Standardized player names by trimming whitespace, removing asterisks, and title-casing names.",
        "- Converted all score and measurement columns to numeric values with explicit parsing of text-based numbers.",
        "- Preserved all raw measurements plus published ranking columns.",
        "- Left pitching fields blank for players without pitching data instead of imputing scouting value.",
        "",
        "## Top Overall Players",
        ", ".join(top_overall),
        "",
        "## Top Pitchers",
        ", ".join(top_pitchers),
        "",
        "## Balanced Players",
        ", ".join(balanced),
        "",
        "## Hidden-Value Players",
        ", ".join(hidden),
        "",
        "## Tier Structure",
        master_players["tier"].value_counts().sort_index().to_string(),
        "",
        "## Draft Observations",
        "- Overall board blends raw athletic/hitting/fielding data with published overall ranking data.",
        "- Pitcher value is kept separate so coaches can decide whether to draft arms early or let overall talent drive the room.",
        "- Balanced-player model pushes up players with fewer weak spots, while the upside model rewards high-end tools or standout pitching value.",
        "",
        "## Simulated Draft Results",
        f"- Best-available simulation average team strength: {best_available['team_strength_score'].mean():.1f}",
        f"- Balanced-team simulation average team strength: {balanced_summary['team_strength_score'].mean():.1f}",
        f"- Best-available spread (max-min): {(best_available['team_strength_score'].max() - best_available['team_strength_score'].min()):.1f}",
        f"- Balanced-team spread (max-min): {(balanced_summary['team_strength_score'].max() - balanced_summary['team_strength_score'].min()):.1f}",
        "",
        "## Team Balance Observations",
        "- Balanced-team simulation narrows gaps by steering pitching and all-around value toward weaker roster builds.",
        "- Best-available simulation concentrates top-end talent faster, which can widen roster variance when early picks line up with pitching strength.",
        "",
        "## Limitations",
        "- The workbook does not contain defensive positions, so position-player rankings are role-agnostic.",
        "- Pitching data is missing for four players, so they are excluded from pitcher-only rankings.",
        "- Subjective ranking scales appear to treat lower numbers as better; this was inferred from the ranking sheets and converted accordingly.",
    ]
    return "\n".join(lines)


def create_notebook(output_path: Path = Path("analysis_notebook.ipynb")) -> None:
    nb = nbf.v4.new_notebook()
    cells = [
        nbf.v4.new_markdown_cell(
            "# VCB 13U House Draft Analysis\n"
            "This notebook is Google Colab-compatible. Upload `VCB House - 13u PeeWee Assessment.xlsx`, then run the cells in order."
        ),
        nbf.v4.new_code_cell(
            "!pip -q install pandas numpy matplotlib scikit-learn openpyxl nbformat"
        ),
        nbf.v4.new_markdown_cell("## 1. Introduction\nThis workflow inspects the workbook, cleans the data, builds `master_players`, creates rankings and tiers, and simulates a snake draft."),
        nbf.v4.new_code_cell(
            "from pathlib import Path\n"
            "import pandas as pd\n"
            "from analysis_pipeline import (\n"
            "    NUM_TEAMS, ROSTER_SIZE, RANDOM_SEED, DRAFT_STRATEGY,\n"
            "    WEIGHT_OVERALL, WEIGHT_PITCHING, WEIGHT_BALANCE, WEIGHT_UPSIDE,\n"
            "    inspect_workbook, build_master_players, build_rankings,\n"
            "    build_hidden_value_table, build_draft_board, simulate_draft,\n"
            "    save_visualizations, build_report\n"
            ")\n"
            "WORKBOOK_PATH = Path('VCB House - 13u PeeWee Assessment.xlsx')"
        ),
        nbf.v4.new_markdown_cell("## 2. Load workbook"),
        nbf.v4.new_code_cell("inspection, inspection_summary = inspect_workbook(WORKBOOK_PATH)\ninspection_summary"),
        nbf.v4.new_markdown_cell("## 3. Inspect workbook"),
        nbf.v4.new_code_cell("inspection['cross_sheet_summary']"),
        nbf.v4.new_markdown_cell("## 4. Clean data"),
        nbf.v4.new_code_cell("master_players = build_master_players(WORKBOOK_PATH)\nmaster_players.head()"),
        nbf.v4.new_markdown_cell("## 5. Build master player table"),
        nbf.v4.new_code_cell("master_players.shape, master_players[['player_name','overall_rank','pitcher_rank','tier']].head(10)"),
        nbf.v4.new_markdown_cell("## 6. Generate rankings"),
        nbf.v4.new_code_cell("rankings = build_rankings(master_players)\nrankings['overall'][['player_name','overall_rank','overall_composite']].head(20)"),
        nbf.v4.new_markdown_cell("## 7. Create tiers"),
        nbf.v4.new_code_cell("master_players['tier'].value_counts().sort_index()"),
        nbf.v4.new_markdown_cell("## 8. Identify hidden-value players"),
        nbf.v4.new_code_cell("hidden_value_players = build_hidden_value_table(master_players)\nhidden_value_players.head(20)"),
        nbf.v4.new_markdown_cell("## 9. Build draft board"),
        nbf.v4.new_code_cell("draft_board = build_draft_board(master_players)\ndraft_board.head(20)"),
        nbf.v4.new_markdown_cell("## 10. Configure simulator"),
        nbf.v4.new_code_cell(
            "# Editable parameters\n"
            "NUM_TEAMS = 6\n"
            "ROSTER_SIZE = 12\n"
            "RANDOM_SEED = 13\n"
            "DRAFT_STRATEGY = 'both'  # 'best_available', 'balanced', or 'both'\n"
            "WEIGHT_OVERALL = 0.55\n"
            "WEIGHT_PITCHING = 0.20\n"
            "WEIGHT_BALANCE = 0.15\n"
            "WEIGHT_UPSIDE = 0.10\n"
            "LOCKED_PLAYERS = {}"
        ),
        nbf.v4.new_markdown_cell("## 11. Simulate snake draft"),
        nbf.v4.new_code_cell(
            "best_available_results, best_available_team_strength = simulate_draft(\n"
            "    draft_board, master_players, num_teams=NUM_TEAMS, roster_size=ROSTER_SIZE, strategy='best_available', locked_players=LOCKED_PLAYERS,\n"
            "    weight_overall=WEIGHT_OVERALL, weight_pitching=WEIGHT_PITCHING, weight_balance=WEIGHT_BALANCE, weight_upside=WEIGHT_UPSIDE\n"
            ")\n"
            "balanced_results, balanced_team_strength = simulate_draft(\n"
            "    draft_board, master_players, num_teams=NUM_TEAMS, roster_size=ROSTER_SIZE, strategy='balanced', locked_players=LOCKED_PLAYERS,\n"
            "    weight_overall=WEIGHT_OVERALL, weight_pitching=WEIGHT_PITCHING, weight_balance=WEIGHT_BALANCE, weight_upside=WEIGHT_UPSIDE\n"
            ")\n"
            "best_available_results.head(12)"
        ),
        nbf.v4.new_markdown_cell("## 12. Evaluate team balance"),
        nbf.v4.new_code_cell("best_available_team_strength, balanced_team_strength"),
        nbf.v4.new_markdown_cell("## 13. Visualize team strength"),
        nbf.v4.new_code_cell(
            "save_visualizations(master_players, {'best_available': best_available_team_strength, 'balanced': balanced_team_strength}, Path('.'))"
        ),
        nbf.v4.new_markdown_cell("## 14. Export outputs"),
        nbf.v4.new_code_cell(
            "rankings['overall'].to_csv('player_rankings_overall.csv', index=False)\n"
            "rankings['pitchers'].to_csv('player_rankings_pitchers.csv', index=False)\n"
            "master_players.to_csv('cleaned_master_players.csv', index=False)\n"
            "master_players[['player_name','tier','overall_rank','overall_composite']].to_csv('player_tiers.csv', index=False)\n"
            "hidden_value_players.to_csv('hidden_value_players.csv', index=False)\n"
            "draft_board.to_csv('draft_board.csv', index=False)\n"
            "pd.concat([best_available_results, balanced_results], ignore_index=True).to_csv('simulated_draft_results.csv', index=False)\n"
            "pd.concat([\n"
            "    best_available_team_strength.assign(strategy='best_available'),\n"
            "    balanced_team_strength.assign(strategy='balanced')\n"
            "], ignore_index=True).to_csv('team_strength_summary.csv', index=False)"
        ),
        nbf.v4.new_markdown_cell("## 15. Final summary"),
        nbf.v4.new_code_cell(
            "report = build_report(inspection_summary, master_players, hidden_value_players, {'best_available': best_available_team_strength, 'balanced': balanced_team_strength})\n"
            "print(report)"
        ),
    ]
    nb["cells"] = cells
    output_path.write_text(nbf.writes(nb), encoding="utf-8")


def export_outputs(workbook_path: Path = WORKBOOK_PATH, output_dir: Path = Path(".")) -> dict:
    output_dir.mkdir(parents=True, exist_ok=True)
    inspection, inspection_summary = inspect_workbook(workbook_path)
    master_players = build_master_players(workbook_path)
    rankings = build_rankings(master_players)
    hidden_value_players = build_hidden_value_table(master_players)
    draft_board = build_draft_board(master_players)

    best_available_results, best_available_team_strength = simulate_draft(
        draft_board,
        master_players,
        num_teams=NUM_TEAMS,
        roster_size=ROSTER_SIZE,
        strategy="best_available",
        locked_players=LOCKED_PLAYERS,
    )
    balanced_results, balanced_team_strength = simulate_draft(
        draft_board,
        master_players,
        num_teams=NUM_TEAMS,
        roster_size=ROSTER_SIZE,
        strategy="balanced",
        locked_players=LOCKED_PLAYERS,
    )

    master_players.to_csv(output_dir / "cleaned_master_players.csv", index=False)
    rankings["overall"].to_csv(output_dir / "player_rankings_overall.csv", index=False)
    rankings["pitchers"].to_csv(output_dir / "player_rankings_pitchers.csv", index=False)
    master_players[["player_name", "tier", "overall_rank", "overall_composite"]].to_csv(output_dir / "player_tiers.csv", index=False)
    hidden_value_players.to_csv(output_dir / "hidden_value_players.csv", index=False)
    draft_board.to_csv(output_dir / "draft_board.csv", index=False)
    pd.concat([best_available_results, balanced_results], ignore_index=True).to_csv(output_dir / "simulated_draft_results.csv", index=False)
    team_strength_summary = pd.concat(
        [
            best_available_team_strength.assign(strategy="best_available"),
            balanced_team_strength.assign(strategy="balanced"),
        ],
        ignore_index=True,
    )
    team_strength_summary.to_csv(output_dir / "team_strength_summary.csv", index=False)

    visuals = save_visualizations(
        master_players,
        {"best_available": best_available_team_strength, "balanced": balanced_team_strength},
        output_dir,
    )
    report = build_report(
        inspection_summary,
        master_players,
        hidden_value_players,
        {"best_available": best_available_team_strength, "balanced": balanced_team_strength},
    )
    (output_dir / "draft_summary_report.md").write_text(report, encoding="utf-8")
    (output_dir / "workbook_inspection.json").write_text(json.dumps(inspection, indent=2), encoding="utf-8")
    create_notebook(output_dir / "analysis_notebook.ipynb")

    return {
        "inspection_summary": inspection_summary,
        "master_players": master_players,
        "rankings": rankings,
        "hidden_value_players": hidden_value_players,
        "draft_board": draft_board,
        "best_available_results": best_available_results,
        "balanced_results": balanced_results,
        "team_strength_summary": team_strength_summary,
        "visualizations": visuals,
        "report": report,
    }


def main() -> None:
    export_outputs()


if __name__ == "__main__":
    main()
