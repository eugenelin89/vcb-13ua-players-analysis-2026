# VCB 13U House Draft Summary

## Workbook Structure
The workbook has four sheets: one overall ranking summary, one pitcher ranking summary, one detailed assessment sheet, and one detailed pitching sheet.
The two summary ranking sheets are not tidy tables at the top of the sheet. They include title rows and embedded header rows, so they need sheet-specific parsing before analysis.
The detailed assessment sheet covers 77 players and contains raw athletic, hitting, and fielding/throwing measurements.
The pitching sheets cover 73 players, which means four players appear to have no pitching evaluation recorded.
Player names are consistent across sheets after standardizing whitespace and capitalization, and there were no duplicate player records after cleanup.
The four players missing pitching-specific data are: Bo Singerman, Elliott Cocke, Luca Di Nozzi, Shael Singerman.

## Cleaning Steps
- Removed title rows and embedded header rows from the ranking sheets.
- Standardized player names by trimming whitespace, removing asterisks, and title-casing names.
- Converted all score and measurement columns to numeric values with explicit parsing of text-based numbers.
- Preserved all raw measurements plus published ranking columns.
- Left pitching fields blank for players without pitching data instead of imputing scouting value.

## Top Overall Players
Charles Love, George Morris, Sean Choi, Arek Aubuchon, Brody Pendergast, Milan Avinashi, Carter Fong, Mylo Mclaughlin, Ashton Chang, Arlo Tourigny

## Top Pitchers
Arek Aubuchon, Sean Choi, Mylo Mclaughlin, Brody Pendergast, Nolan Whitfield, Charles Love, George Morris, Robbie Mcmynn

## Balanced Players
Jack Cheyne, Hayden Gudewill, Bowie Scott, Henry Galer, Felix Vandenenden, Bo Singerman, Oliver Lewis, Liam Pearce

## Hidden-Value Players
Oliver Lewis, Armann Freimanis, River Yonge, Dylan Bansback, Shael Singerman, Gray Dickson, Francis De La Cruz, Mylo Mclaughlin

## Tier Structure
tier
Tier 1 - Impact Players         10
Tier 2 - Strong Starters        17
Tier 3 - Solid Contributors     37
Tier 4 - Development Players    13

## Draft Observations
- Overall board blends raw athletic/hitting/fielding data with published overall ranking data.
- Pitcher value is kept separate so coaches can decide whether to draft arms early or let overall talent drive the room.
- Balanced-player model pushes up players with fewer weak spots, while the upside model rewards high-end tools or standout pitching value.

## Simulated Draft Results
- Best-available simulation average team strength: 517.9
- Balanced-team simulation average team strength: 517.9
- Best-available spread (max-min): 17.2
- Balanced-team spread (max-min): 19.6

## Team Balance Observations
- Balanced-team simulation narrows gaps by steering pitching and all-around value toward weaker roster builds.
- Best-available simulation concentrates top-end talent faster, which can widen roster variance when early picks line up with pitching strength.

## Limitations
- The workbook does not contain defensive positions, so position-player rankings are role-agnostic.
- Pitching data is missing for four players, so they are excluded from pitcher-only rankings.
- Subjective ranking scales appear to treat lower numbers as better; this was inferred from the ranking sheets and converted accordingly.