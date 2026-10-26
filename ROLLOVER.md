# Game...Blouses Season Rollover Checklist

Do these before the first game of the new season.

## Google Sheets
- [ ] All-Time workbook, Data tab: add the finished season's 10 rows (Years, Manager, Status, 1st/2nd/3rd, Wins, Losses, Ties, Games, Playoffs)
- [ ] All-Time Lookup tab: add any new managers, mark departures Inactive
- [ ] Stats workbook: clear Weekly Scores PF/PA, update Team / TEAM NAME for any new owners or renamed teams
- [ ] Stats workbook: point each owner's IMPORTHTML team log at the new season's Yahoo URLs
- [ ] Confirm SCORING MODIFIERS still match Yahoo settings

## Repo
- [ ] fetch_data.py: update `SEASON`, `REG_WEEKS`, `PLAYOFF_WEEKS`, `FEE`, `PAYOUTS` if anything changed
- [ ] Replace draft_2026.xlsx with the new draft (same columns: Pick, Round, Owner, Player, NBA Team, Positions) and update `DRAFT_PATH` plus the workflow `paths` list
- [ ] template.html Rules page: update playoff end date, trade deadline, draft date, Yahoo league ID if new
- [ ] Run the Action manually and check the Data check panel on Standings

## Spec
- [ ] Update SPEC.md owners, money, dates
