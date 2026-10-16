# BIOGLOW™ live scores

Edit the workbook, push to `main`, and GitHub Actions rebuilds and publishes the site.

## One-time setup
1. Create a GitHub repo and push this folder to it.
2. Repo **Settings → Pages → Build and deployment → Source: GitHub Actions**.
3. Put the workbook (.xlsx or .xlsm) in `data/`. Keep only one file there.
4. Download the official logo from the FIRST BIOGLOW brand page and save it as
   `site/assets/bioglow-logo.png` (use the file as provided; do not recolor or crop).

## At the event
Enter scores, **save in Excel** (so formula results are stored), then:

    git add data && git commit -m "scores" && git push

The site updates in about a minute; open pages refresh themselves every 30 seconds.

## Notes
- The tab is found by its headers (`Team #`, `Team Name`, `Round 1`...). To force a tab: set env `SHEET_NAME`.
- Rank is by each team's highest round score (Raw Performance and GP columns are ignored); ties share a rank. Unscored teams show "–".
- Test locally: `pip install openpyxl && python scripts/build_site.py`, then open `_site/index.html`.
