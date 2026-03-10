from __future__ import annotations

import base64
import html
import math
import subprocess
from pathlib import Path

import pandas as pd


ROOT = Path(".")
HTML_OUT = ROOT / "README_google_docs_export.html"
DOCX_OUT = ROOT / "README_google_docs_export.docx"


def img_data_uri(path: Path) -> str:
    raw = path.read_bytes()
    encoded = base64.b64encode(raw).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def fmt(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float):
        if math.isclose(value, round(value)):
            return str(int(round(value)))
        return f"{value:.1f}"
    return str(value)


def df_to_html_table(df: pd.DataFrame) -> str:
    headers = "".join(f"<th>{html.escape(str(col))}</th>" for col in df.columns)
    rows = []
    for _, row in df.iterrows():
        cells = "".join(f"<td>{html.escape(fmt(v))}</td>" for v in row.tolist())
        rows.append(f"<tr>{cells}</tr>")
    return f"<table><thead><tr>{headers}</tr></thead><tbody>{''.join(rows)}</tbody></table>"


def section(title: str, body: str) -> str:
    return f"<section><h2>{html.escape(title)}</h2>{body}</section>"


def image_block(title: str, filename: str) -> str:
    uri = img_data_uri(ROOT / filename)
    return (
        f"<div class='chart-block'>"
        f"<h3>{html.escape(title)}</h3>"
        f"<img src='{uri}' alt='{html.escape(title)}' />"
        f"</div>"
    )


def build_html() -> str:
    board = pd.read_csv(ROOT / "draft_board.csv")
    hidden = pd.read_csv(ROOT / "hidden_value_players.csv")
    teams = pd.read_csv(ROOT / "suggested_balanced_7_teams.csv")

    team_summary = (
        teams.groupby("team")
        .agg(
            players=("player_name", "count"),
            overall_strength=("overall_composite", "sum"),
            pitching_strength=("pitching_score", "sum"),
            avg_overall_rank=("overall_rank", "mean"),
        )
        .reset_index()
    )

    team_sections = []
    for team_num in sorted(teams["team"].unique()):
        sub = (
            teams.loc[teams["team"] == team_num, ["player_name", "overall_rank", "pitcher_rank", "tier", "top_3_strengths"]]
            .sort_values("overall_rank")
            .reset_index(drop=True)
        )
        team_sections.append(section(f"Team {team_num}", df_to_html_table(sub)))

    visual_html = "".join(
        [
            image_block("Overall Score Distribution", "overall_score_distribution.png"),
            image_block("Pitching Score Distribution", "pitching_score_distribution.png"),
            image_block("Top 20 Overall Players", "top_20_overall_players.png"),
            image_block("Top Pitchers", "top_pitchers.png"),
            image_block("Tier Distribution", "tier_distribution.png"),
            image_block("Category Correlation Heatmap", "category_correlation_heatmap.png"),
            image_block("Suggested 7-Team Overall Strength", "suggested_7_teams_strength.png"),
            image_block("Suggested 7-Team Pitching Strength", "suggested_7_teams_pitching.png"),
        ]
    )

    top10 = board[["Player Name", "Overall Rank", "Pitcher Rank", "Tier", "Top 3 Strengths", "Risk / Weakness Flag"]].head(10)
    hidden10 = hidden[["player_name", "tier", "value_delta", "top_3_strengths", "risk_flag"]].head(10)

    workbook_structure = """
    <ul>
      <li>4 sheets total</li>
      <li>77 unique players in the full pool</li>
      <li>73 players with pitching-specific data</li>
      <li>4 players missing pitching-specific data: Bo Singerman, Elliott Cocke, Luca Di Nozzi, Shael Singerman</li>
      <li>The ranking sheets require custom parsing because they contain title rows and embedded headers</li>
    </ul>
    """

    model_summary = """
    <p>The analysis keeps pitching value separate from overall player value and builds four ranking lenses.</p>
    <ul>
      <li><b>Raw score model:</b> weighted direct scoring from athleticism, hitting, fielding, throwing, and pitching components</li>
      <li><b>Normalized score model:</b> z-score based comparison across categories</li>
      <li><b>Balanced-player model:</b> rewards fewer weak spots across the player profile</li>
      <li><b>Specialist/upside model:</b> rewards standout tools and ceiling</li>
    </ul>
    <p>Use Overall Rank for best available, Pitcher Rank for arm priority, Balanced Score for safer profiles, and Upside Score when chasing ceiling.</p>
    """

    html_doc = f"""
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8" />
      <title>VCB 13U House Draft Analysis</title>
      <style>
        body {{
          font-family: "Avenir Next", "Helvetica Neue", Arial, sans-serif;
          margin: 0;
          padding: 0;
          background: #f7f3ea;
          color: #152238;
          line-height: 1.45;
        }}
        .page {{
          max-width: 980px;
          margin: 0 auto;
          padding: 36px 42px 60px;
          background: linear-gradient(180deg, #fffdf8 0%, #f7f3ea 100%);
        }}
        h1 {{
          font-size: 34px;
          margin: 0 0 8px;
        }}
        .subtitle {{
          color: #6b7280;
          margin-bottom: 28px;
          font-size: 16px;
        }}
        h2 {{
          margin-top: 34px;
          border-bottom: 3px solid #c95c32;
          padding-bottom: 6px;
          font-size: 24px;
        }}
        h3 {{
          margin-top: 24px;
          margin-bottom: 10px;
          font-size: 18px;
        }}
        .hero {{
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin: 22px 0 30px;
        }}
        .card {{
          background: #fffdf8;
          border: 2px solid #d8d2c2;
          border-radius: 16px;
          padding: 16px;
        }}
        .card .label {{
          color: #6b7280;
          font-size: 12px;
          text-transform: uppercase;
          letter-spacing: 0.08em;
        }}
        .card .value {{
          font-size: 28px;
          font-weight: 700;
          margin-top: 8px;
        }}
        .charts {{
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 18px;
        }}
        .chart-block {{
          background: #fffdf8;
          border: 1px solid #d8d2c2;
          border-radius: 16px;
          padding: 14px;
          break-inside: avoid;
        }}
        .chart-block img {{
          width: 100%;
          border-radius: 10px;
        }}
        table {{
          width: 100%;
          border-collapse: collapse;
          margin-top: 12px;
          font-size: 12px;
          background: #fffdf8;
        }}
        th {{
          background: #152238;
          color: white;
          text-align: left;
          padding: 8px;
        }}
        td {{
          border-bottom: 1px solid #e8e2d4;
          padding: 8px;
          vertical-align: top;
        }}
        tr:nth-child(even) td {{
          background: #fbf8f1;
        }}
        ul {{
          padding-left: 20px;
        }}
        section {{
          break-inside: avoid;
        }}
      </style>
    </head>
    <body>
      <div class="page">
        <h1>VCB 13U House Draft Analysis</h1>
        <div class="subtitle">Google Docs-friendly export with embedded infographics, rankings, and a suggested balanced 7-team build.</div>

        <div class="hero">
          <div class="card"><div class="label">Players</div><div class="value">77</div></div>
          <div class="card"><div class="label">Pitchers Evaluated</div><div class="value">73</div></div>
          <div class="card"><div class="label">Suggested Teams</div><div class="value">7</div></div>
          <div class="card"><div class="label">Roster Size</div><div class="value">11</div></div>
        </div>

        {section("Workbook Structure", workbook_structure)}
        {section("Visual Overview", f"<div class='charts'>{visual_html}</div>")}
        {section("Draft Model Summary", model_summary)}
        {section("Top 10 Overall Players", df_to_html_table(top10))}
        {section("Hidden-Value And Role-Fit Players", df_to_html_table(hidden10))}
        {section("Suggested Balanced 7-Team Build", "<p>This build uses all 77 players as 7 teams of 11 and tries to minimize the spread in overall strength, pitching strength, and top-end talent concentration.</p>" + df_to_html_table(team_summary))}
        {''.join(team_sections)}
      </div>
    </body>
    </html>
    """
    return html_doc


def main() -> None:
    HTML_OUT.write_text(build_html(), encoding="utf-8")
    subprocess.run(
        [
            "textutil",
            "-convert",
            "docx",
            "-format",
            "html",
            str(HTML_OUT),
            "-output",
            str(DOCX_OUT),
        ],
        check=True,
    )


if __name__ == "__main__":
    main()
