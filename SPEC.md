# Game...Blouses Site Spec

Repo code is the source of truth. If this spec and the code disagree, flag it and ask.

## League
- Yahoo NBA league #86606, "Game...Blouses", 10 teams, head-to-head points, no divisions
- Season label: 2026-27 (`SEASON = 2026` in fetch_data.py)
- Regular season weeks 1-20; playoffs 4 teams, weeks 21-22 (ends Sun Apr 4, 2027), reseeding, 3rd place game (no payout)
- Standings order: win % then Points For
- Scoring (read live from the SCORING MODIFIERS tab): PTS 1, REB 1.2, AST 1.5, ST 3, BLK 3, TO -1
- Roster: G, G, F, F, C, Util, 1 BN, 5 IL

## Money
- Fee $75 x 10 = $750
- Champion $375, Runner-up $75, Regular season (best record, PF tiebreak) $300

## Owners (site shows owner names only, never Yahoo team names)
MARZ, KARDIAN, ADAM B, DOM, FUR, MEANY, PAT CLARK, PHIL, T-ROW, BUBBA
- Owner list and order come from the Weekly Scores tab (`Team` column)
- History names are uppercased as-is (Pat Clark -> PAT CLARK, T-Row -> T-ROW); inactive managers (ODED, MERGY, KEV, CONOR) are not merged with anyone

## Data sources
| Source | What | Notes |
|---|---|---|
| Stats workbook (published xlsx) | Weekly Scores, Consolidated Data, SCORING MODIFIERS | Weekly scores hand-entered by Marz |
| All-Time workbook (published xlsx) | Data tab (one row per manager per season) | No PF/PA history; Data Summary tab is not used |
| draft_2026.xlsx (in repo) | 70 picks, 7-round snake | Transcribed from the Yahoo draft results PDF |

- A week counts only when all 10 teams have Points For and Points Against entered
- Matchups are inferred from mirrored PF/PA (no opponent column)
- Consolidated Data = Yahoo team logs, season to date, active-slot stats only. Every player who ever scored for an owner stays listed after being dropped; players appear only after their first game
- Player string parsing: `Name` + optional status (O, D, INJ, GTD, DTD, SUSP, NA) + `Player Note` / `New Player Note` + `TEAM - POS` + optional game time
- Draft names match to game log names by normalized full name, then first initial + last name (handles "N. Alexander-Walker")

## Formulas
- Player FP = sum(stat x modifier)
- All-play: each week, wins vs every team with a lower score; Exp. W = all-play % x games
- Luck = 50 + 50 x (win % - all-play %)
- Performance = normal CDF of z-score of team avg PF vs league team averages, x 100
- Position group = LAST listed eligibility (G,F -> Forward; F,C -> Center); display order Center, Forward, Guard
- Draft value = draft rank within position group - current FP rank within group (all players with games). Drafted players with no games share last rank

## Pages
Standings (with Luck and strength, rank chart, Data check), Weekly, Stat Categories, Positions, Teams, Draft (Full board, By team, Best and worst, By position, Added players), League History, Rules

## Design
- Light page (#fbfafd), plum header (#3a1670), violet (#7b35d6), gold (#f4be1c), orange (#e9782a) used sparingly
- Fonts: Shrikhand (display), Barlow / Barlow Semi Condensed (body, tables)
- No Chappelle skit references in copy (colors only)
- All tables sortable; chart legend click highlights one series and greys the rest

## Build
- fetch_data.py -> data/gameblouses_data.json; build_site.py injects JSON into template.html -> index.html
- Never edit index.html. Change template.html or fetch_data.py and let the Action rebuild
- GitHub Action: daily 11:00 UTC, manual run, and on push of source files
- URL: gameblouses.redplanetanalytics.com (CNAME file in repo; Cloudflare CNAME `gameblouses` -> smarz1223.github.io, DNS only)
