# Game...Blouses

Fantasy basketball league site: https://gameblouses.redplanetanalytics.com

- `fetch_data.py` pulls the published Google Sheets and writes `data/gameblouses_data.json`
- `build_site.py` builds `index.html` from `template.html`
- `.github/workflows/refresh.yml` runs both every morning

See SPEC.md and ROLLOVER.md. Do not edit index.html by hand.
